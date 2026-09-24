// カード生成ツール ⇔ Claude Code 中継サーバー（列伝のAI強化用）
// ・ログイン済みの Claude Code（プロプラン枠）を `claude -p` で呼び出す。APIキーは使わない。
// ・Claude に許可するのは Web検索・Web閲覧のみ（ファイル操作やコマンド実行は不可）。
// ・プロプランの上限に達して断られた場合は、上限到達として記録し、呼び出し元は従来の列伝のまま使う。
const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');

const PORT = 3950;
const USAGE_FILE = path.join(__dirname, 'claude_bridge_usage.json');
const MODELS = { opus: 'opus', sonnet: 'sonnet' };
const TIMEOUT_MS = 240000;
// 空の作業フォルダで実行し、プロジェクトのファイルや設定を読ませない
const WORK_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'claude-bridge-'));

// ── 使用履歴 ─────────────────────────────────────────────
function loadUsage() {
  try { return JSON.parse(fs.readFileSync(USAGE_FILE, 'utf8')); }
  catch { return { calls: [], lastLimit: null }; }
}
function saveUsage(u) {
  const cutoff = Date.now() - 30 * 86400000;
  u.calls = u.calls.filter(c => c.t >= cutoff);
  fs.writeFileSync(USAGE_FILE, JSON.stringify(u, null, 1));
}
function summarize(u) {
  const now = Date.now();
  const sum = (since) => {
    const cs = u.calls.filter(c => c.t >= since);
    return {
      calls: cs.length,
      ok: cs.filter(c => c.ok).length,
      tokens: cs.reduce((a, c) => a + (c.inTok || 0) + (c.outTok || 0), 0),
      costUsd: Math.round(cs.reduce((a, c) => a + (c.costUsd || 0), 0) * 100) / 100,
    };
  };
  const midnight = new Date(); midnight.setHours(0, 0, 0, 0);
  const lim = u.lastLimit;
  const limitedNow = !!(lim && (lim.resetAt ? lim.resetAt > now : now - lim.at < 3600000));
  return {
    last5h: sum(now - 5 * 3600000),
    today: sum(midnight.getTime()),
    last7d: sum(now - 7 * 86400000),
    lastLimit: lim,
    limitedNow,
    busy,
  };
}

// ── 上限到達の判定 ─────────────────────────────────────────
function detectLimit(text) {
  if (!/(usage limit|limit reached|hit your limit|rate limit|out of (extra )?usage|5-hour limit|weekly limit)/i.test(text)) return null;
  let resetAt = null;
  const epoch = text.match(/limit reached\|(\d{10,13})/i);
  if (epoch) resetAt = epoch[1].length === 13 ? +epoch[1] : +epoch[1] * 1000;
  return { at: Date.now(), resetAt, message: text.slice(0, 200) };
}

// ── claude -p 実行 ──────────────────────────────────────────
function runClaude(prompt, model) {
  return new Promise((resolve) => {
    const env = { ...process.env };
    for (const k of Object.keys(env)) if (k === 'CLAUDECODE' || k.startsWith('CLAUDE_CODE_')) delete env[k];
    const args = ['-p', '--model', model, '--output-format', 'json', '--allowedTools', 'WebSearch,WebFetch'];
    // Windows の claude は .cmd のため shell 経由で起動（引数は固定値のみ、プロンプトは標準入力で渡す）
    const child = spawn('claude', args, { cwd: WORK_DIR, env, shell: process.platform === 'win32', windowsHide: true });
    let out = '', err = '';
    const timer = setTimeout(() => { child.kill(); }, TIMEOUT_MS);
    child.stdout.on('data', d => out += d);
    child.stderr.on('data', d => err += d);
    child.on('error', e => { clearTimeout(timer); resolve({ spawnError: e.message }); });
    child.on('close', code => { clearTimeout(timer); resolve({ code, out, err }); });
    child.stdin.end(prompt, 'utf8');
  });
}

function extractJson(text) {
  const s = text.indexOf('{'), e = text.lastIndexOf('}');
  if (s < 0 || e <= s) return null;
  try { return JSON.parse(text.slice(s, e + 1)); } catch { return null; }
}

let busy = false;
let queue = Promise.resolve();

async function enrich(body) {
  const model = MODELS[body.model] || 'opus';
  const usage = loadUsage();
  const r = await runClaude(String(body.prompt || ''), model);
  const entry = { t: Date.now(), model, ok: false };

  if (r.spawnError) {
    return { status: 500, json: { error: 'Claude Code を起動できませんでした: ' + r.spawnError } };
  }
  let meta = extractJson(r.out) || {};
  const resultText = typeof meta.result === 'string' ? meta.result : r.out;
  const limit = detectLimit(`${resultText}\n${r.err}`);
  if (meta.usage) {
    entry.inTok = (meta.usage.input_tokens || 0) + (meta.usage.cache_creation_input_tokens || 0) + (meta.usage.cache_read_input_tokens || 0);
    entry.outTok = meta.usage.output_tokens || 0;
  }
  if (typeof meta.total_cost_usd === 'number') entry.costUsd = meta.total_cost_usd;

  if (limit && (meta.is_error || r.code !== 0 || !meta.result)) {
    entry.limited = true;
    usage.calls.push(entry);
    usage.lastLimit = limit;
    saveUsage(usage);
    return { status: 429, json: { limited: true, message: limit.message, usage: summarize(usage) } };
  }
  if (r.code !== 0 || meta.is_error) {
    usage.calls.push(entry);
    saveUsage(usage);
    if (/authenticat|oauth|401|not logged in|\/login/i.test(`${resultText}\n${r.err}`)) {
      return { status: 401, json: { authError: true, error: 'Claude Code のログインが切れています。ターミナルで「claude」を起動し「/login」でログインし直してください。', usage: summarize(usage) } };
    }
    return { status: 502, json: { error: (resultText || r.err || 'unknown error').slice(0, 300), usage: summarize(usage) } };
  }
  const parsed = extractJson(resultText);
  entry.ok = !!(parsed && typeof parsed.retsuden === 'string');
  usage.calls.push(entry);
  if (usage.lastLimit && entry.ok) usage.lastLimit = null;
  saveUsage(usage);
  if (!entry.ok) return { status: 502, json: { error: '回答の形式が不正でした', usage: summarize(usage) } };
  return { status: 200, json: { retsuden: parsed.retsuden, sources: parsed.sources || [], model, usage: summarize(usage) } };
}

// ── HTTP ────────────────────────────────────────────────
// 他サイトから勝手に呼ばれないよう、ローカルで開いたページ（file:// / localhost）からのみ受け付ける
function originAllowed(origin) {
  if (!origin || origin === 'null') return true;
  return /^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin);
}

http.createServer((req, res) => {
  const origin = req.headers.origin;
  if (!originAllowed(origin)) { res.writeHead(403); res.end(); return; }
  const cors = {
    'Access-Control-Allow-Origin': origin || '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Private-Network': 'true',
  };
  const send = (status, obj) => {
    res.writeHead(status, { ...cors, 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify(obj));
  };
  if (req.method === 'OPTIONS') { res.writeHead(204, cors); res.end(); return; }

  if (req.method === 'GET' && req.url === '/usage') { send(200, summarize(loadUsage())); return; }

  if (req.method === 'POST' && req.url === '/retsuden') {
    let raw = '';
    req.on('data', d => { raw += d; if (raw.length > 200000) req.destroy(); });
    req.on('end', () => {
      let body;
      try { body = JSON.parse(raw); } catch { send(400, { error: 'bad json' }); return; }
      const usage = loadUsage();
      const s = summarize(usage);
      if (s.limitedNow) { send(429, { limited: true, message: usage.lastLimit.message, usage: s }); return; }
      queue = queue.then(async () => {
        busy = true;
        try { const r = await enrich(body); send(r.status, r.json); }
        catch (e) { send(500, { error: e.message }); }
        finally { busy = false; }
      });
    });
    return;
  }
  send(404, { error: 'not found' });
}).listen(PORT, '127.0.0.1', () => {
  console.log(`Claude 中継サーバー起動: http://localhost:${PORT}`);
  console.log('カード生成ツールの「AI列伝強化」をONにすると、ここ経由で Claude Code が呼ばれます。');
  console.log('このウィンドウは閉じずにそのままにしてください。');
});
