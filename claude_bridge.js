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
const crypto = require('crypto');
const USAGE_FILE = path.join(__dirname, 'claude_bridge_usage.json');
const CACHE_FILE = path.join(__dirname, 'claude_bridge_cache.json');
const MODELS = { opus: 'opus', sonnet: 'sonnet' };
// ログイン情報ファイル（中身は読まず、更新日時だけで「ログインし直したか」を判定する）
const CRED_FILE = path.join(os.homedir(), '.claude', '.credentials.json');
function credMtime() {
  try { return fs.statSync(CRED_FILE).mtimeMs; } catch { return 0; }
}
const AUTH_MSG = 'Claude Code のログインが切れています。ターミナルで「claude」を起動し「/login」でログインし直してください。';
// 実績がまだ無いときの所要時間の目安（秒）
const DEFAULT_EST_SEC = { opus: 75, sonnet: 50 };
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
  // 直近の成功5回の所要時間の中央値を目安にする
  const estSec = {};
  for (const m of Object.keys(MODELS)) {
    const ds = u.calls.filter(c => c.ok && c.model === m && c.durMs).slice(-5).map(c => c.durMs).sort((a, b) => a - b);
    estSec[m] = ds.length ? Math.round(ds[Math.floor(ds.length / 2)] / 1000) : DEFAULT_EST_SEC[m];
  }
  const midnight = new Date(); midnight.setHours(0, 0, 0, 0);
  const lim = u.lastLimit;
  const limitedNow = !!(lim && (lim.resetAt ? lim.resetAt > now : now - lim.at < 3600000));
  return {
    last5h: sum(now - 5 * 3600000),
    today: sum(midnight.getTime()),
    last7d: sum(now - 7 * 86400000),
    lastLimit: lim,
    limitedNow,
    authError: authErrorNow(u),
    busy,
    estSec,
  };
}

// ログイン切れ後、ログイン情報が更新されていなければ即座に「ログイン切れ」と返す（3分待たせない）
function authErrorNow(u) {
  return !!(u.lastAuthError && u.lastAuthError.credMtime === credMtime());
}

// ── 上限到達の判定 ─────────────────────────────────────────
function detectLimit(text) {
  if (!/(usage limit|limit reached|hit your limit|rate limit|out of (extra )?usage|5-hour limit|weekly limit)/i.test(text)) return null;
  let resetAt = null;
  const epoch = text.match(/limit reached\|(\d{10,13})/i);
  if (epoch) resetAt = epoch[1].length === 13 ? +epoch[1] : +epoch[1] * 1000;
  // 例: "You've hit your limit · resets 9pm (Etc/GMT-9)" / "resets 9:30am" → 次に来るその時刻（PCの現地時刻）
  const hm = !resetAt && text.match(/resets\s+(?:[A-Za-z]{3,9}\s+\d{1,2},?\s+)?(\d{1,2})(?::(\d{2}))?\s*(am|pm)/i);
  if (hm) {
    let h = parseInt(hm[1], 10) % 12;
    if (/pm/i.test(hm[3])) h += 12;
    const d = new Date(); d.setHours(h, hm[2] ? parseInt(hm[2], 10) : 0, 0, 0);
    if (d.getTime() <= Date.now()) d.setDate(d.getDate() + 1);
    resetAt = d.getTime();
  }
  return { at: Date.now(), resetAt, message: text.slice(0, 200) };
}

// ── claude -p 実行 ──────────────────────────────────────────
function runClaude(prompt, model, allowWeb = true) {
  return new Promise((resolve) => {
    const env = { ...process.env };
    for (const k of Object.keys(env)) if (k === 'CLAUDECODE' || k.startsWith('CLAUDE_CODE_')) delete env[k];
    // 高速化: 思考は軽め・MCP/スキル/セッション保存は読み込まない（校正工程ではWebも使わない）
    const args = ['-p', '--model', model, '--output-format', 'json',
      ...(allowWeb ? ['--allowedTools', 'WebSearch,WebFetch'] : []),
      '--effort', 'low', '--strict-mcp-config', '--disable-slash-commands', '--no-session-persistence'];
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

// Claudeの回答から {"retsuden": ...} を頑丈に取り出す。
// 前後の説明文・コードブロック・下書きと清書の2つのJSON・本文中の半角"による崩れに対応する。
function parseRetsudenJson(text) {
  if (!text) return null;
  const t = String(text).replace(/```(?:json)?/gi, '');
  // ① 括弧の対応で {...} の候補をすべて拾い、後ろ（清書）から順に試す
  const cands = [];
  for (let i = 0; i < t.length; i++) {
    if (t[i] !== '{') continue;
    let depth = 0, inStr = false, esc = false;
    for (let j = i; j < t.length; j++) {
      const ch = t[j];
      if (inStr) {
        if (esc) esc = false; else if (ch === '\\') esc = true; else if (ch === '"') inStr = false;
        continue;
      }
      if (ch === '"') inStr = true;
      else if (ch === '{') depth++;
      else if (ch === '}' && --depth === 0) { cands.push(t.slice(i, j + 1)); break; }
    }
  }
  for (const c of cands.reverse()) {
    try { const o = JSON.parse(c); if (o && typeof o.retsuden === 'string' && o.retsuden.trim()) return o; } catch {}
  }
  // ② JSONとして読めない場合：最後の "retsuden": "..." を直接取り出す（本文中の改行・半角"にも対応）
  const re = /"retsuden"\s*:\s*"([\s\S]*?)"\s*(?:,\s*"sources"|\}\s*(?:$|[^"]))/g;
  let m, last = null;
  while ((m = re.exec(t)) !== null) last = m;
  if (last) {
    const body = last[1].replace(/\\n/g, '\n').replace(/\\"/g, '"').replace(/\\\\/g, '\\').trim();
    if (body) return { retsuden: body, sources: [] };
  }
  return null;
}

// 追加の呼び出し（校正・整形）のトークン量を集計に加える
function addUsage(entry, meta) {
  if (!meta || !meta.usage) return;
  const u = meta.usage;
  entry.inTok = (entry.inTok || 0) + (u.input_tokens || 0) + (u.cache_creation_input_tokens || 0) + (u.cache_read_input_tokens || 0);
  entry.outTok = (entry.outTok || 0) + (u.output_tokens || 0);
}

// 読み取れなかった回答を後から原因調査できるよう記録する（直近20件）
const ERROR_LOG = path.join(__dirname, 'claude_bridge_errors.log');
function logBadResponse(kind, text) {
  try {
    const entry = `=== ${new Date().toLocaleString()} ${kind} ===\n${String(text).slice(0, 20000)}\n`;
    let old = '';
    try { old = fs.readFileSync(ERROR_LOG, 'utf8'); } catch {}
    const parts = (old + entry).split(/(?==== )/).slice(-20);
    fs.writeFileSync(ERROR_LOG, parts.join(''));
  } catch {}
}

// ── 結果の保存（同じ依頼は即時に返し、プロプラン枠を使わない） ──
function cacheKey(model, prompt) {
  return crypto.createHash('sha1').update(model + '\n' + prompt).digest('hex');
}
function loadCache() {
  try { return JSON.parse(fs.readFileSync(CACHE_FILE, 'utf8')); } catch { return {}; }
}
function saveCache(key, value) {
  const c = loadCache();
  c[key] = { ...value, savedAt: Date.now() };
  fs.writeFileSync(CACHE_FILE, JSON.stringify(c, null, 1));
}

let busy = false;
let queue = Promise.resolve();

async function enrich(body) {
  const model = MODELS[body.model] || 'opus';
  const prompt = String(body.prompt || '');
  const usage = loadUsage();
  const started = Date.now();
  const r = await runClaude(prompt, model);
  const entry = { t: Date.now(), model, ok: false, durMs: Date.now() - started };

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
    const isAuth = /authenticat|oauth|401|not logged in|\/login/i.test(`${resultText}\n${r.err}`);
    if (isAuth) usage.lastAuthError = { at: Date.now(), credMtime: credMtime() };
    saveUsage(usage);
    if (isAuth) return { status: 401, json: { authError: true, error: AUTH_MSG, usage: summarize(usage) } };
    return { status: 502, json: { error: (resultText || r.err || 'unknown error').slice(0, 300), usage: summarize(usage) } };
  }
  let parsed = parseRetsudenJson(resultText);
  if (!parsed) {
    // 形式が崩れて読み取れない場合：回答を記録し、本文だけをJSONに整え直す依頼を1回だけ行う（Webなし・短時間）
    logBadResponse(`draft-unparsable (${model})`, resultText);
    const f0 = Date.now();
    const fixPrompt = [
      '次の文章は、野球カードの「列伝」（選手紹介文）を書くよう依頼したときの回答です。',
      'この中から完成した列伝の本文だけを取り出し、次のJSONだけを出力してください（前後に説明文を付けない）。',
      '本文の内容・表現は変えない。本文中の半角の " は「」に置き換える。説明文・調査メモ・下書きは含めない。',
      '{"retsuden":"本文（文ごとに\\nで改行）","sources":[]}',
      '',
      '回答:',
      String(resultText).slice(0, 12000),
    ].join('\n');
    const fr = await runClaude(fixPrompt, model, false);
    const fm = extractJson(fr.out || '') || {};
    addUsage(entry, fm);
    entry.durMs += Date.now() - f0;
    parsed = typeof fm.result === 'string' ? parseRetsudenJson(fm.result) : null;
    if (parsed) entry.repaired = true;
    else logBadResponse(`repair-failed (${model})`, fm.result || fr.out || fr.err);
  }
  entry.ok = !!parsed;
  let text = entry.ok ? parsed.retsuden : '';

  // 校正工程（依頼があるときのみ）：事実を変えずに日本語だけを整える。失敗したら下書きをそのまま使う
  if (entry.ok && typeof body.polishPrompt === 'string' && body.polishPrompt.includes('{{TEXT}}')) {
    const p0 = Date.now();
    const pr = await runClaude(body.polishPrompt.replace('{{TEXT}}', text), model, false);
    const pm = extractJson(pr.out || '') || {};
    addUsage(entry, pm);
    entry.durMs += Date.now() - p0;
    const pj = typeof pm.result === 'string' ? parseRetsudenJson(pm.result) : null;
    const len = s => s.replace(/\s/g, '').length;
    if (pr.code === 0 && !pm.is_error && pj && typeof pj.retsuden === 'string' && len(pj.retsuden) >= len(text) * 0.9) {
      text = pj.retsuden;
      entry.polished = true;
    }
  }

  usage.calls.push(entry);
  if (usage.lastLimit && entry.ok) usage.lastLimit = null;
  usage.lastAuthError = null;
  saveUsage(usage);
  if (!entry.ok) return { status: 502, json: { error: '回答の形式が不正でした', usage: summarize(usage) } };
  const result = { retsuden: text, sources: parsed.sources || [], model, polished: !!entry.polished };
  saveCache(cacheKey(model, prompt), result);
  return { status: 200, json: { ...result, usage: summarize(usage) } };
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
      const model = MODELS[body.model] || 'opus';
      const hit = !body.refresh && loadCache()[cacheKey(model, String(body.prompt || ''))];
      if (hit) { send(200, { ...hit, cached: true, usage: s }); return; }
      if (s.authError) { send(401, { authError: true, error: AUTH_MSG, usage: s }); return; }
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
