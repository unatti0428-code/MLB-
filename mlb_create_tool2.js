'use strict';
const http      = require('http');
const https     = require('https');
const XLSX      = require('xlsx');
const ExcelJS   = require('exceljs');
const fs        = require('fs');
const path      = require('path');
const os        = require('os');
const crypto    = require('crypto');
const puppeteer = require('puppeteer-core');
const { spawnSync, spawn } = require('child_process');

const PORT    = 3941;
const OUT_DIR = __dirname;

// ── .env 読み込み ────────────────────────────────────────────────────────────
const ENV_PATH = path.join(__dirname, '.env');
if (fs.existsSync(ENV_PATH)) {
  fs.readFileSync(ENV_PATH, 'utf8').split(/\r?\n/).forEach(line => {
    const m = line.match(/^([A-Z_][A-Z0-9_]*)\s*=\s*(.+)$/);
    if (m) process.env[m[1]] = m[2].trim();
  });
}

// ── Chrome detection ──────────────────────────────────────────────────────────
function findChrome() {
  const lapp = process.env.LOCALAPPDATA || '';
  const pf   = process.env.ProgramFiles  || '';
  const candidates = [
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
    path.join(lapp, 'Google\\Chrome\\Application\\chrome.exe'),
    'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
    path.join(pf, 'Microsoft\\Edge\\Application\\msedge.exe'),
  ];
  return candidates.find(p => { try { return fs.existsSync(p); } catch { return false; } }) || null;
}

// ── PowerShell file browsers ──────────────────────────────────────────────────
function browseFileWithFilter(filter) {
  const r = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', `
[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
$d = New-Object System.Windows.Forms.OpenFileDialog
$d.Filter = "${filter}"
if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($d.FileName)
  [Convert]::ToBase64String($bytes)
}`], { encoding: 'buffer' });
  const b64 = (r.stdout || Buffer.alloc(0)).toString('ascii').trim();
  return b64 ? Buffer.from(b64, 'base64').toString('utf8') : '';
}
const browseFile = () => browseFileWithFilter('Excel Files (*.xlsx)|*.xlsx');

// ── MLB Stats API ─────────────────────────────────────────────────────────────
function mlbGet(url) {
  return new Promise((resolve, reject) => {
    https.get(url, {
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
    }, res => {
      let buf = '';
      res.on('data', c => buf += c);
      res.on('end', () => {
        try { resolve(JSON.parse(buf)); }
        catch (e) { reject(new Error('MLB API parse error: ' + buf.slice(0, 120))); }
      });
    }).on('error', reject);
  });
}

// ── Anthropic API (Claude + web search) ──────────────────────────────────────
function httpsPost(options, bodyStr) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, res => {
      let raw = '';
      res.on('data', c => raw += c);
      res.on('end', () => resolve({ status: res.statusCode, body: raw }));
    });
    req.on('error', reject);
    if (bodyStr) req.write(bodyStr);
    req.end();
  });
}

/**
 * Claude にウェブ検索させて投手の球種データを推定。
 * 戻り値: { pitches: [{name, speed, pct}, ...], note: '' } または null
 */
async function callClaudeForPitchData(apiKey, playerName, years) {
  const prompt =
`あなたはMLBの球種データ専門家です。以下の投手の球種情報をウェブ検索で調べ、JSONのみで回答してください。

投手名: ${playerName}
対象年度: ${years.join(', ')}（この期間を代表する球種レパートリー）

以下フォーマットのJSONのみを返してください（説明文・マークダウン・コードブロック不要）:
{"pitches":[{"name":"4-Seam Fastball","speed":92,"pct":55},{"name":"Slider","speed":83,"pct":30},{"name":"Changeup","speed":80,"pct":15}],"note":"データ根拠"}

・球種名は必ず次のいずれか: 4-Seam Fastball, Two-Seam Fastball, Sinker, Slider, Sweeper, Changeup, Circle Change, Curveball, 12-6 Curve, Cutter, Splitter, Forkball, Split Finger
・speed は実際の球速(mph)を整数で記載
・pct は投球割合(合計100になるよう整数で調整)
・球種は最大5種類、使用率5%未満は省略
・Sinker と Two-Seam Fastball の区別: 縦方向に大きく沈む球（日本のシンカー含む）→ Sinker。横変化中心の速球（カット系・右打者への食い込み重視）→ Two-Seam Fastball`;

  const messages = [{ role: 'user', content: prompt }];

  for (let turn = 0; turn < 8; turn++) {
    const body = JSON.stringify({
      model: 'claude-opus-4-5',
      max_tokens: 1024,
      tools: [{ type: 'web_search_20250305', name: 'web_search' }],
      messages,
    });
    let parsed;
    try {
      const { body: raw } = await httpsPost({
        hostname: 'api.anthropic.com',
        port: 443,
        path: '/v1/messages',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-beta': 'web-search-2025-03-05',
          'Content-Length': Buffer.byteLength(body),
        },
      }, body);
      parsed = JSON.parse(raw);
    } catch { break; }

    if (parsed.error) throw new Error(parsed.error.message || 'Claude API error');

    const content    = parsed.content || [];
    const stopReason = parsed.stop_reason;

    if (stopReason === 'end_turn') {
      const textBlock = content.find(b => b.type === 'text');
      if (!textBlock) break;
      const m = textBlock.text.trim().match(/\{[\s\S]*\}/);
      if (!m) break;
      try { return JSON.parse(m[0]); } catch { break; }
    }

    if (stopReason === 'tool_use') {
      messages.push({ role: 'assistant', content });
      const autoResults = content.filter(b => b.type === 'tool_result');
      if (autoResults.length > 0) {
        messages.push({ role: 'user', content: '検索完了。JSONのみ返してください。' });
        continue;
      }
      const toolUseBlocks = content.filter(b => b.type === 'tool_use');
      if (!toolUseBlocks.length) break;
      messages.push({ role: 'user', content: toolUseBlocks.map(b => ({ type: 'tool_result', tool_use_id: b.id, content: '' })) });
      continue;
    }

    const textBlock = content.find(b => b.type === 'text');
    if (textBlock) {
      const m = textBlock.text.match(/\{[\s\S]*\}/);
      if (m) { try { return JSON.parse(m[0]); } catch {} }
    }
    break;
  }
  return null;
}

/**
 * Claude にウェブ検索させてキャリアピーク球速プロファイルを取得。
 * 戻り値: { pitches: [{name, peakSpeed, avgPct}, ...], note: '' } または null
 * peakSpeed: キャリア最高球速（瞬間最速・レーダーガン最大値, mph）
 *            ※ 呼び出し元で -3.7mph（≈-6km/h）してシーズン平均相当に変換する
 */
async function callClaudeForPeakProfile(apiKey, playerName, debutYear) {
  const prompt =
`MLBピッチャー「${playerName}」（${debutYear}年デビュー）の球種・球速データをWikipediaやBaseballReference、FanGraphsなどで調べ、以下のJSONフォーマットのみで回答してください。

{"pitches":[{"name":"4-Seam Fastball","peakSpeed":102,"avgPct":55},{"name":"Slider","peakSpeed":94,"avgPct":35}],"note":"データ根拠URL等"}

ルール:
・peakSpeed: キャリアを通じた最高球速（瞬間最速・レーダーガン計測最大値）(mph)を整数で
  ※ シーズン平均ではなく、記録・文献・Wikipedia等で確認できる最高球速を記載すること
・avgPct: キャリア代表的な投球割合(%, 合計100の整数)
・球種名は必ず次のいずれか: 4-Seam Fastball, Two-Seam Fastball, Sinker, Slider, Sweeper, Changeup, Circle Change, Curveball, 12-6 Curve, Cutter, Splitter, Forkball, Split Finger
・最大5球種（使用率5%未満は省略）
・Sinker と Two-Seam Fastball の区別: 縦方向に大きく沈む球（日本のシンカー含む）→ Sinker。横変化中心の速球 → Two-Seam Fastball
・JSONのみ返答（説明文・コードブロック不要）`;

  const messages = [{ role: 'user', content: prompt }];

  for (let turn = 0; turn < 8; turn++) {
    const body = JSON.stringify({
      model: 'claude-opus-4-5',
      max_tokens: 512,
      tools: [{ type: 'web_search_20250305', name: 'web_search' }],
      messages,
    });
    let parsed;
    try {
      const { body: raw } = await httpsPost({
        hostname: 'api.anthropic.com',
        port: 443,
        path: '/v1/messages',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-beta': 'web-search-2025-03-05',
          'Content-Length': Buffer.byteLength(body),
        },
      }, body);
      parsed = JSON.parse(raw);
    } catch { break; }

    if (parsed.error) throw new Error(parsed.error.message || 'Claude API error');

    const content    = parsed.content || [];
    const stopReason = parsed.stop_reason;

    if (stopReason === 'end_turn') {
      const textBlock = content.find(b => b.type === 'text');
      if (!textBlock) break;
      const m = textBlock.text.trim().match(/\{[\s\S]*\}/);
      if (!m) break;
      try { return JSON.parse(m[0]); } catch { break; }
    }

    if (stopReason === 'tool_use') {
      messages.push({ role: 'assistant', content });
      const autoResults = content.filter(b => b.type === 'tool_result');
      if (autoResults.length > 0) {
        messages.push({ role: 'user', content: '検索完了。JSONのみ返してください。' });
        continue;
      }
      const toolUseBlocks = content.filter(b => b.type === 'tool_use');
      if (!toolUseBlocks.length) break;
      messages.push({ role: 'user', content: toolUseBlocks.map(b => ({ type: 'tool_result', tool_use_id: b.id, content: '' })) });
      continue;
    }

    const textBlock = content.find(b => b.type === 'text');
    if (textBlock) {
      const m = textBlock.text.match(/\{[\s\S]*\}/);
      if (m) { try { return JSON.parse(m[0]); } catch {} }
    }
    break;
  }
  return null;
}

/**
 * THT・MLB.com・Baseball Reference・Wikipedia・Baseball America 等の複数ソースで
 * FanGraphs BIS データの球種・球速・投球割合の誤りを検証・補正する情報を取得する。
 * 主に BIS 時代 (2002〜2014) に適用し、Statcast 実測がない年の品質を向上させる。
 *
 * @param {string}   apiKey      Anthropic API キー
 * @param {string}   playerName  選手名 (英語)
 * @param {string[]} targetYears 検証対象年 (e.g. ['2007','2008'])
 * @param {Object}   fgSummary   現在の FanGraphs 概要 {yr: {key: {speedMph, pct}}}
 * @returns {{ yearCorrections, careerPeakSpeeds, pitcherCharacteristics, note } | null}
 *   yearCorrections: [{ years:[], pitches:[{key, avgSpeedMph, pct, note}] }]
 *     avgSpeedMph = PITCHf/x 実測値または公式記録（BIS バイアスなし・ブースト不要）
 *   careerPeakSpeeds: { key: peakMph } キャリア最高球速（瞬間最速）→ 全年への速度キャップに使用
 */
async function callClaudeForFgCorrection(apiKey, playerName, targetYears, fgSummary) {
  const yearRange = targetYears.length > 0
    ? `${Math.min(...targetYears.map(Number))}〜${Math.max(...targetYears.map(Number))}年`
    : '不明';
  const fgJson = JSON.stringify(fgSummary, null, 2);

  const prompt =
`MLB投手「${playerName}」（${yearRange}活躍）の球種データについて、複数の一次資料で検証・補正してください。

【FanGraphs BIS から取得している現在のデータ（キー: ff=4シーム/sl=スライダー/ch=チェンジアップ/cu=カーブ/fc=カッター/si=シンカー/fs=スプリット）】
${fgJson}

【検索する資料（優先度順）】
1. The Hardball Times (tht.fangraphs.com) — 2007-2009年の PITCHf/x 分析記事（最優先）
2. MLB.com 公式選手ページ — 投球レパートリー・球速記録
3. Baseball Reference (baseball-reference.com) — 公式成績・投球スタイル記述
4. Wikipedia・Baseball Hall of Fame 公式 (baseballhall.org) — 文献・証言記録
5. Baseball America スカウトレポート — 球種・球速・特徴

【FanGraphs BIS の典型的な誤り（検証ポイント）】
・FB% に 4 シーム(ff)と 2 シーム/シンカー(si) が混在 → PITCHf/x や文献で実際の球種を確認
・BIS 時代（〜2007）の球速は実際より 2〜4mph 低く記録される場合がある
・投球割合が実際のレパートリーと大きく乖離するケースがある

以下のJSONフォーマットのみで返答してください（説明文・コードブロック・その他テキスト一切不要）:
{"yearCorrections":[{"years":["2007","2008"],"pitches":[{"key":"ff","avgSpeedMph":93,"pct":55,"note":"THT PITCHf/x 2008"},{"key":"sl","avgSpeedMph":83,"pct":35,"note":"THT"},{"key":"ch","avgSpeedMph":81,"pct":10,"note":"BR"}]}],"careerPeakSpeeds":{"ff":97,"sl":87},"pitcherCharacteristics":"速球/スライダー型右腕、2007年以前は主にシンカー系","note":"THT 2008, Baseball Reference"}

ルール:
・key は必ず ff/sl/ch/cu/fc/si/fs のいずれか
・avgSpeedMph は PITCHf/x 実測値またはスカウト記録の平均球速（mph 整数）。FanGraphs BIS より正確な値のみ記載
・pct は整数（年ごとに合計 100 になるよう調整）
・careerPeakSpeeds はキャリア通じた瞬間最大球速（レーダーガン計測最大値, mph 整数）
・yearCorrections は確実な資料根拠がある年・球種のみ記載（推測で補正しない）
・訂正なし・確認不能な場合は yearCorrections:[]、careerPeakSpeeds:{} を返す`;

  const messages = [{ role: 'user', content: prompt }];

  for (let turn = 0; turn < 10; turn++) {
    const body = JSON.stringify({
      model: 'claude-opus-4-5',
      max_tokens: 1024,
      tools: [{ type: 'web_search_20250305', name: 'web_search' }],
      messages,
    });
    let parsed;
    try {
      const { body: raw } = await httpsPost({
        hostname: 'api.anthropic.com',
        port: 443,
        path: '/v1/messages',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-beta': 'web-search-2025-03-05',
          'Content-Length': Buffer.byteLength(body),
        },
      }, body);
      parsed = JSON.parse(raw);
    } catch { break; }

    if (parsed.error) throw new Error(parsed.error.message || 'Claude API error');

    const content    = parsed.content || [];
    const stopReason = parsed.stop_reason;

    if (stopReason === 'end_turn') {
      const textBlock = content.find(b => b.type === 'text');
      if (!textBlock) break;
      const m = textBlock.text.trim().match(/\{[\s\S]*\}/);
      if (!m) break;
      try { return JSON.parse(m[0]); } catch { break; }
    }

    if (stopReason === 'tool_use') {
      messages.push({ role: 'assistant', content });
      const autoResults = content.filter(b => b.type === 'tool_result');
      if (autoResults.length > 0) {
        messages.push({ role: 'user', content: '検索完了。JSONのみ返してください。' });
        continue;
      }
      const toolUseBlocks = content.filter(b => b.type === 'tool_use');
      if (!toolUseBlocks.length) break;
      messages.push({ role: 'user', content: toolUseBlocks.map(b => ({ type: 'tool_result', tool_use_id: b.id, content: '' })) });
      continue;
    }

    const textBlock = content.find(b => b.type === 'text');
    if (textBlock) {
      const m = textBlock.text.match(/\{[\s\S]*\}/);
      if (m) { try { return JSON.parse(m[0]); } catch {} }
    }
    break;
  }
  return null;
}

async function searchPlayers(name) {
  const data = await mlbGet(
    `https://statsapi.mlb.com/api/v1/people/search?names=${encodeURIComponent(name)}&sportId=1`
  );
  return (data.people || []).map(p => ({
    id:       p.id,
    name:     p.fullName,
    position: p.primaryPosition?.abbreviation || '?',
    debut:    (p.mlbDebutDate || '').slice(0, 4) || '?',
  }));
}

// ── Rate formatting helpers ───────────────────────────────────────────────────
// .275 → "275"
function fmtAvg(val) {
  if (val == null || val === '--' || val === '') return '--';
  const s = String(val).trim();
  const n = parseFloat(s);
  if (isNaN(n)) return '--';
  if (n > 0 && n < 1.0) return String(Math.round(n * 1000));
  return s;
}

// 投球率: 5%未満マスク後の有効な球種の合計が100になるよう再分配して整数化する。
// Input : array of 7 raw pct values (形式: '--', "45.4", "45.4%", etc.)
// Output: array of 7 integer strings (or '--' for masked/missing entries)
// アルゴリズム: Largest Remainder Method
//   ① 有効値のみ合計 → スケール係数 = 100 / 合計
//   ② 各値を floor で切り捨て → 余り = 100 - floor合計
//   ③ 小数部の大きい順に +1 して合計をちょうど100にする
// 例) FF=45.4, SL=15.1, CU=18.7, FC=15.8 (合計95.0) →
//     スケール: FF=47.79, SL=15.89, CU=19.68, FC=16.63 →
//     floor:  FF=47,    SL=15,    CU=19,    FC=16   (合計97)
//     +1 を小数部大順3個: SL(0.89), FF(0.79), CU(0.68) → FF=48, SL=16, CU=20, FC=16 (合計100)
function normalizePctToSum100(pctVals) {
  // ─ Step 1: 有効値を解析 ─
  const parsed = pctVals.map(v => {
    if (!v || v === '--') return null;
    const s = String(v).replace('%', '').trim();
    const n = parseFloat(s);
    return (isNaN(n) || n <= 0) ? null : n;
  });

  const total = parsed.reduce((s, n) => s + (n ?? 0), 0);
  if (total <= 0) return pctVals.map(() => '--');

  // ─ Step 2: 100 にスケーリングして floor ─
  const scaled  = parsed.map(n => (n === null) ? null : (n * 100 / total));
  const floors  = scaled.map(n => (n === null) ? null : Math.floor(n));
  const floorSum = floors.reduce((s, n) => s + (n ?? 0), 0);
  let   toAdd   = 100 - floorSum;

  // ─ Step 3: 小数部が大きい順に +1 ─
  const order = parsed
    .map((n, i) => ({ i, frac: (n === null) ? -1 : (scaled[i] - floors[i]) }))
    .sort((a, b) => b.frac - a.frac);

  const result = [...floors];
  for (const { i } of order) {
    if (toAdd <= 0) break;
    if (result[i] !== null) { result[i]++; toAdd--; }
  }

  return result.map(n => (n === null) ? '--' : String(n));
}

// イニング文字列 → 小数: "200.1" → 200.333...
function parseIP(ipStr) {
  const s = String(ipStr || '').trim();
  if (!s || s === '--') return 0;
  const [whole, frac] = s.split('.');
  return (parseInt(whole) || 0) + (parseInt(frac || 0)) / 3;
}

// IP文字列 → アウト数（重み付け用）
function ipToOuts(ip) {
  const s = String(ip || '0');
  const [w, f] = s.split('.');
  return (parseInt(w) || 0) * 3 + (parseInt(f || 0));
}

// ── スタミナ計算式 (守備.ods AC3 の等価実装) ──────────────────────────────
// 守備.ods AC3 の数式:
//   IFERROR(IFS(V3>=230, ROUND(V3/W3*12.5), V3>=210, ROUND(V3/W3*13.1),
//               V3>=86,  ROUND(V3/W3*13.5), V3>=65,  ROUND(V3/W3*20),
//               V3>=50,  ROUND(V3/W3*21),   V3<=49,  ROUND(V3/W3*22)), "")
// V3 = 換算イニング(K列), W3 = 試合数(G列)
// GS補正: H列(GS)/G列(試合数) > 0.5 の場合、IP>=65 の係数を 20→13 に変更
function calcStaminaFromIP(ip, g, gs) {
  if (!ip || isNaN(ip) || ip < 0) return '';
  if (!g  || isNaN(g)  || g  <= 0) return '';
  const ratio = ip / g;
  if (ip >= 230) return Math.round(ratio * 12.5);
  if (ip >= 210) return Math.round(ratio * 13.1);
  if (ip >= 86)  return Math.round(ratio * 13.5);
  // IP 85以下: スタミナ上限 69（先発でも投球回数不足なら上限を設ける）
  if (ip >= 65) {
    const mult = (gs > 0 && (gs / g) > 0.5) ? 13 : 20;
    return Math.min(69, Math.round(ratio * mult));
  }
  if (ip >= 50)  return Math.min(69, Math.round(ratio * 21));
  return Math.min(69, Math.round(ratio * 22));  // ip <= 49
}

// ── 制球計算式 (守備.ods AC2 の等価実装) ─────────────────────────────────────
// 守備.ods AC2 の数式:
//   IFERROR(IFS(V2>=4.2, ROUND(60-(V2-4.2)/0.16),
//               V2>=1.2, ROUND(85-(V2-1.2)/0.12),
//               V2>=0,   ROUND(100-V2/0.08)), "")
// V2 = 四死球(O列) / 換算イニング(K列) × 9  (= BB9)
function calcSeikyuFromBB9(bb9) {
  if (bb9 == null || isNaN(bb9) || bb9 < 0) return '';
  if (bb9 >= 4.2) return Math.round(60 - (bb9 - 4.2) / 0.16);
  if (bb9 >= 1.2) return Math.round(85 - (bb9 - 1.2) / 0.12);
  return Math.round(100 - bb9 / 0.08);
}

// ── 精神計算式 (守備.ods AE2 の等価実装) ─────────────────────────────────────
// W2 = 防御率(F列)
function calcSeisinFromERA(era) {
  if (era == null || isNaN(era) || era < 0) return '';
  if (era >= 8.2) return Math.round(55 - (era - 8.2) / 0.35);
  if (era >= 6.6) return Math.round(60 - (era - 6.6) / 0.32);
  if (era >= 5.2) return Math.round(65 - (era - 5.2) / 0.28);
  if (era >= 4.0) return Math.round(70 - (era - 4.0) / 0.24);
  if (era >= 3.2) return Math.round(75 - (era - 3.2) / 0.16);
  if (era >= 2.5) return Math.round(80 - (era - 2.5) / 0.14);
  if (era >= 1.9) return Math.round(85 - (era - 1.9) / 0.12);
  if (era >= 1.4) return Math.round(90 - (era - 1.4) / 0.1);
  if (era >= 1.0) return Math.round(95 - (era - 1.0) / 0.08);
  return Math.round(100 - (era - 0.7) / 0.06);
}

// ── 奪三振計算式 (守備.ods AF2 の等価実装) ────────────────────────────────────
// X2 = 奪三振(P列) / 換算イニング × 9  (= K/9)
function calcSanshinFromK9(k9) {
  if (k9 == null || isNaN(k9) || k9 < 0) return '';
  if (k9 <= 6)  return Math.round(40 + (k9 - 6)  / 0.2);
  if (k9 <= 10) return Math.round(80 + (k9 - 10) / 0.1);
  if (k9 <= 30) return Math.round(100 + (k9 - 14) / 0.2);
  return '';
}

// ── 重さ計算式 (守備.ods AG2 の等価実装) ─────────────────────────────────────
// Y2 = 被本塁打(N列) / 換算イニング × 9  (= HR/9)
function calcOmosaFromHR9(hr9) {
  if (hr9 == null || isNaN(hr9) || hr9 < 0) return '';
  if (hr9 >= 2.2)  return Math.round(50  - (hr9 - 2.2)  / 0.1);
  if (hr9 >= 1.8)  return Math.round(55  - (hr9 - 1.8)  / 0.08);
  if (hr9 >= 1.5)  return Math.round(60  - (hr9 - 1.5)  / 0.06);
  if (hr9 >= 1.3)  return Math.round(65  - (hr9 - 1.3)  / 0.04);
  if (hr9 >= 1.0)  return Math.round(80  - (hr9 - 1.0)  / 0.02);
  if (hr9 >= 0.25) return Math.round(105 - (hr9 - 0.25) / 0.03);
  if (hr9 >= 0.1)  return Math.round(110 - (hr9 - 0.1)  / 0.03);
  return '';
}

// ── 対左計算式 (守備.ods AH2 の等価実装) ─────────────────────────────────────
// Z2 = 被打率(Q列) - 対左被打率(S列)  ※どちらも整数形式（.275→275）
function calcTaiHidariFromDiff(z) {
  if (z == null || isNaN(z)) return '';
  if (z < -60) return Math.round(-15 + (60 + z) / 8);
  if (z > 60)  return Math.round(15 + (z - 60) / 8);
  return Math.round(z / 4);
}

// ── 対盗塁計算式 (守備.ods AI2=AJ2+AK2 の等価実装) ──────────────────────────
// AA2=SB(T列), AA3=PK(U列), AB2=換算IP(K列), AB3=CS(V列)
function calcTaiTouruiFromSBData(sb, pk, ip, cs) {
  if (!ip || ip <= 0) return '';
  const sb9 = (sb / ip) * 9;
  let aj;
  if (sb9 >= 1)      aj = -7;
  else if (sb9 >= 0) aj = Math.round(11 - sb9 * 18);
  else               return '';
  const denom = sb + cs;
  // SB・CS・PK がすべて0で盗塁関連の計算ができない場合はデフォルト値 5 を返す
  if (denom <= 0) return 5;
  const ratio = (sb - pk) / denom;
  let ak;
  if (ratio >= 0.85)      ak = -10;
  else if (ratio <= 0.35) ak = 18;
  else ak = Math.round((0.65 - ratio) * 60);
  return aj + ak;
}

// ── Cell styling ──────────────────────────────────────────────────────────────
const PURPLE_FILL     = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7030A0' } };
const RED_PURPLE_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC3399' } };
function purpleCell(cell, value, fs) {
  cell.value = value;
  cell.fill  = { ...PURPLE_FILL };
  cell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: fs };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
}
function redPurpleCell(cell, value, fs) {
  cell.value = value;
  cell.fill  = { ...RED_PURPLE_FILL };
  cell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: fs };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
}

// ── 球種グループ (守備.ods 行14〜20 対応) ────────────────────────────────────
// ★ name はデータセクションの PITCH_NAMES_JA と完全に一致させること。
//   pitchNameOverrides がない場合（Savant以前の選手等）に this.name がフォールバックとして使われる。
//   ・idx=5 は 'シンカー' が正: PITCH_NAMES_JA[5]='シンカー' と一致。
//     現代のツーシーマー（Savant上 'Two Seamer'/'Two-Seam Fastball'）は
//     trackSubtype → SUBTYPE_DISPLAY_JA → pitchNameOverrides[5]='ツーシーム' で上書きされる。
const PITCH_GROUPS = [
  { idx: 0, name: 'フォーシーム',   startCol: 23 },
  { idx: 1, name: 'スライダー',     startCol: 27 },
  { idx: 2, name: 'チェンジアップ', startCol: 31 },
  { idx: 3, name: 'カーブ',         startCol: 35 },
  { idx: 4, name: 'カットボール',   startCol: 39 },
  { idx: 5, name: 'シンカー',       startCol: 43 },  // ← 'ツーシーム' から修正（PITCH_NAMES_JA[5] と統一）
  { idx: 6, name: 'スプリット',     startCol: 47 },
];
const PITCH_ABILITY_START_COL = 59;

function calcKyuSoku(velo) {
  if (velo == null || isNaN(velo) || velo < 10) return '';
  if (velo > 11) return Math.round(velo * 1.6 + 4);
  return '';
}
function calcAH_pitch(idx, p) {
  if (p == null || isNaN(p)) return '';
  switch (idx) {
    case 0:
      if (p >= 300) return Math.round(55  - (p-300)/4);
      if (p >= 250) return Math.round(80  - (p-250)/2);
      if (p >= 235) return Math.round(85  - (p-235)/3);
      if (p >= 215) return Math.round(90  - (p-215)/4);
      if (p >= 190) return Math.round(95  - (p-190)/5);
      if (p >= 150) return Math.round(100 - (p-150)/8);
      if (p >= 1)   return Math.round(105 - (p-70)/16);
      return '';
    case 1:
      if (p >= 300) return Math.round(55  - (p-300)/6);
      if (p >= 200) return Math.round(80  - (p-200)/4);
      if (p >= 150) return Math.round(90  - (p-150)/5);
      if (p >= 120) return Math.round(95  - (p-120)/6);
      if (p >= 1)   return Math.round(105 - (p-60)/6);
      return '';
    case 2: case 3:
      if (p >= 300) return Math.round(55  - (p-300)/6);
      if (p >= 200) return Math.round(80  - (p-200)/4);
      if (p >= 150) return Math.round(90  - (p-150)/5);
      if (p >= 120) return Math.round(95  - (p-120)/6);
      if (p >= 80)  return Math.round(100 - (p-80)/8);
      if (p >= 1)   return Math.round(100 - (p-80)/12);
      return '';
    case 4:
      if (p >= 290) return Math.round(55  - (p-290)/4);
      if (p >= 240) return Math.round(80  - (p-240)/2);
      if (p >= 225) return Math.round(85  - (p-225)/3);
      if (p >= 205) return Math.round(90  - (p-205)/4);
      if (p >= 180) return Math.round(95  - (p-180)/5);
      if (p >= 1)   return Math.round(105 - (p-80)/10);
      return '';
    case 5:
      if (p >= 330) return Math.round(50  - (p-330)/6);
      if (p >= 310) return Math.round(55  - (p-310)/4);
      if (p >= 260) return Math.round(80  - (p-260)/2);
      if (p >= 245) return Math.round(85  - (p-245)/3);
      if (p >= 220) return Math.round(90  - (p-220)/5);
      if (p >= 195) return Math.round(95  - (p-195)/5);
      if (p >= 150) return Math.round(100 - (p-150)/9);
      if (p >= 1)   return Math.round(105 - (p-70)/16);
      return '';
    case 6:
      if (p >= 285) return Math.round(55  - (p-285)/5);
      if (p >= 245) return Math.round(65  - (p-245)/4);
      if (p >= 200) return Math.round(80  - (p-200)/3);
      if (p >= 110) return Math.round(95  - (p-110)/6);
      if (p >= 1)   return Math.round(105 - (p-30)/8);
      return '';
    default: return '';
  }
}
function calcAI_pitch(idx, q) {
  if (q == null || isNaN(q)) return '';
  switch (idx) {
    case 0: case 4: case 5:
      if (q >= 570) return Math.round(60  - (q-570)/10);
      if (q >= 500) return Math.round(70  - (q-500)/7);
      if (q >= 300) return Math.round(90  - (q-300)/10);
      if (q >= 100) return Math.round(100 - (q-100)/20);
      if (q >= 1)   return Math.round(105 - (q-50)/10);
      return '';
    case 1: case 2: case 3:
      if (q >= 470) return Math.round(65  - (q-470)/8);
      if (q >= 420) return Math.round(70  - (q-420)/10);
      if (q >= 300) return Math.round(80  - (q-300)/12);
      if (q >= 250) return Math.round(90  - (q-250)/5);
      if (q >= 215) return Math.round(95  - (q-215)/7);
      if (q >= 180) return Math.round(100 - (q-180)/7);
      if (q >= 100) return Math.round(105 - (q-100)/16);
      if (q >= 1)   return Math.round(110 - (q-50)/10);
      return '';
    case 6:
      if (q >= 310) return Math.round(80  - (q-310)/6);
      if (q >= 270) return Math.round(85  - (q-270)/8);
      if (q >= 1)   return Math.round(102 - (q-32)/14);
      return '';
    default: return '';
  }
}
function calcAK_pitch(idx, aj, r) {
  if (aj === '' || aj == null || aj === 0) return 0;
  const typeA = (idx === 0 || idx === 4 || idx === 5);
  if (typeA) {
    if (aj >= 80) {
      if (r >= 65) return 8/3;  if (r >= 60) return 7/3;  if (r >= 55) return 6/3;
      if (r >= 50) return 5/3;  if (r >= 45) return 4/3;  if (r >= 40) return 3/3;
      if (r >= 35) return 2/3;  if (r >= 30) return 1/3;  if (r >= 18.5) return 0;
      if (r >= 18) return -1;   if (r >= 16) return -2;   if (r >= 14) return -4;
      if (r >= 12) return -6;   if (r >= 10) return -8;   if (r >= 8)  return -10;
    } else if (aj >= 70) {
      if (r >= 65) return 4;    if (r >= 60) return 3.5;  if (r >= 55) return 3;
      if (r >= 50) return 2.5;  if (r >= 45) return 2;    if (r >= 40) return 1.5;
      if (r >= 35) return 1;    if (r >= 30) return 0.5;  if (r >= 18.5) return 0;
      if (r >= 18) return -1/3; if (r >= 16) return -2/3; if (r >= 14) return -4/3;
      if (r >= 12) return -6/3; if (r >= 10) return -8/3; if (r >= 8)  return -10/3;
    } else if (aj >= 40) {
      if (r >= 65) return 8; if (r >= 60) return 7; if (r >= 55) return 6;
      if (r >= 50) return 5; if (r >= 45) return 4; if (r >= 40) return 3;
      if (r >= 35) return 2; if (r >= 30) return 1;
    }
  } else {
    if (aj >= 80) {
      if (r >= 65) return 6;    if (r >= 60) return 16/3; if (r >= 55) return 14/3;
      if (r >= 50) return 4;    if (r >= 45) return 10/3; if (r >= 40) return 8/3;
      if (r >= 35) return 2;    if (r >= 30) return 4/3;  if (r >= 25) return 2/3;
      if (r >= 18.5) return 0;  if (r >= 18) return -2;   if (r >= 16) return -4;
      if (r >= 14)  return -6;  if (r >= 12) return -8;   if (r >= 10) return -10;
      if (r >= 8)  return -12;
    } else if (aj >= 70) {
      if (r >= 65) return 9;    if (r >= 60) return 8;    if (r >= 55) return 7;
      if (r >= 50) return 6;    if (r >= 45) return 5;    if (r >= 40) return 4;
      if (r >= 35) return 3;    if (r >= 30) return 2;    if (r >= 25) return 1;
      if (r >= 18.5) return 0;  if (r >= 18) return -1;   if (r >= 16) return -2;
      if (r >= 14)  return -3;  if (r >= 12) return -4;   if (r >= 10) return -5;
      if (r >= 8)  return -6;
    } else if (aj >= 40) {
      if (r >= 65) return 18; if (r >= 60) return 16; if (r >= 55) return 14;
      if (r >= 50) return 12; if (r >= 45) return 10; if (r >= 40) return 8;
      if (r >= 35) return 6;  if (r >= 30) return 4;  if (r >= 25) return 2;
      if (r >= 20) return 1;
    }
  }
  return 0;
}
function calcKyuI(aj, ak, r) {
  if (aj === '' || aj == null || aj === 0) return '';
  const sum = aj + ak;
  if (r <= 8 && sum >= 85) return 85;
  return Math.ceil(sum);
}

// ── 緩急計算 (守備.ods 盗塁能シート参照) ─────────────────────────────────────
// 線形補外付きテーブル参照（範囲外は端点の傾きで外挿）
function _tblLookup(value, t, v) {
  if (value == null || isNaN(value)) return null;
  const n = t.length;
  const asc = t[1] > t[0];
  if (asc) {
    if (value < t[0])     return v[0] + (v[1] - v[0]) / (t[1] - t[0]) * (value - t[0]);
    if (value > t[n - 1]) return v[n-1] + (v[n-1] - v[n-2]) / (t[n-1] - t[n-2]) * (value - t[n-1]);
    let r = v[0]; for (let i = 0; i < n; i++) { if (value >= t[i]) r = v[i]; } return r;
  } else {
    if (value > t[0])     return v[0] + (v[1] - v[0]) / (t[1] - t[0]) * (value - t[0]);
    if (value < t[n - 1]) return v[n-1] + (v[n-1] - v[n-2]) / (t[n-1] - t[n-2]) * (value - t[n-1]);
    for (let i = 0; i < n; i++) { if (value >= t[i]) return v[i]; } return v[n - 1];
  }
}
// ① ERA → 緩急スコア (A13:M14)  ERA降順テーブル
function _kERA(era) {
  const t = [5.50, 5.00, 4.50, 4.00, 3.50, 3.00, 2.50, 2.00, 1.50, 1.00, 0.50, 0.00];
  const v = [  25,   28,   30,   33,   35,   38,   40,   43,   45,   48,   50,   53];
  return _tblLookup(era, t, v);
}
// ② 制球 → 緩急スコア (A16:N17)  昇順テーブル
function _kSeikyu(s) {
  const t = [ 50,  55,  60,  65,  70,  75,  80,  85,  90,  95, 100, 105, 110];
  const v = [ 25,  28,  30,  33,  35,  38,  40,  43,  45,  48,  50,  53,  55];
  return _tblLookup(s, t, v);
}
// ③-CH: チェンジアップ威力 → 緩急スコア (A19:N20)
function _kCH(p) {
  const t = [ 50,  55,  60,  65,  70,  75,  80,  85,  90,  95, 100, 105, 110];
  const v = [ 50,  55,  60,  65,  70,  75,  80,  85,  90,  95, 100, 105, 110];
  return _tblLookup(p, t, v);
}
// ③-CU: カーブ威力 → 緩急スコア (A22:N23)
function _kCU(p) {
  const t = [  50,   55,   60,   65,   70,   75,   80,   85,   90,   95,  100,  105,  110];
  const v = [  45, 49.5,   54, 58.5,   63, 67.5,   72, 76.5,   81, 85.5,   90, 94.5,   99];
  return _tblLookup(p, t, v);
}
// ③-SL/FS: スライダー・スプリット威力 → 緩急スコア (A25:N26)
function _kSLFS(p) {
  const t = [ 50,  55,  60,  65,  70,  75,  80,  85,  90,  95, 100, 105, 110];
  const v = [ 40,  44,  48,  52,  56,  60,  64,  68,  72,  76,  80,  84,  88];
  return _tblLookup(p, t, v);
}
// 緩急メイン: ROUND((①+②+③)÷2)
// eraStr = 防御率文字列、seikyu = 制球能力値(数値 or '')、pitchData = {[idx]:{ba,slg,pct}}
// kyuiMap: BA/SLG なし年（pre-Statcast）用フォールバック { [idx]: kyui値 }
function calcKankyuu(eraStr, seikyu, pitchData, kyuiMap = {}) {
  const v1 = _kERA(parseFloat(String(eraStr || '').trim()));
  if (v1 === null) return '';
  const v2 = _kSeikyu(seikyu !== '' && seikyu != null ? Number(seikyu) : NaN);
  if (v2 === null) return '';

  // 各変化球の球威を計算してテーブルで変換、最大値を③に採用
  // BA/SLG が '--' の場合は kyuiMap（showKyuiMap の年分）をフォールバックとして使用
  const kyuiOf = (idx, pd) => {
    if (!pd) return null;
    const pctStr = String(pd.pct ?? '').trim();
    const pctNum = Number(pctStr);
    if (!pctStr || pctStr === '--' || isNaN(pctNum) || pctNum <= 0) return null;
    const baNum  = pd.ba  != null && String(pd.ba)  !== '--' ? Number(pd.ba)  : NaN;
    const slgNum = pd.slg != null && String(pd.slg) !== '--' ? Number(pd.slg) : NaN;
    const ah = calcAH_pitch(idx, baNum);
    const ai = calcAI_pitch(idx, slgNum);
    if (ah !== '' && ai !== '') {
      const aj = (Number(ah) + Number(ai)) / 2;
      const ki = calcKyuI(aj, calcAK_pitch(idx, aj, pctNum), pctNum);
      return ki !== '' ? Number(ki) : null;
    }
    // フォールバック①: showKyuiMap（The Show / FanGraphs / 推定データ）
    const override = kyuiMap[idx];
    if (override !== undefined) return Number(override);
    // フォールバック②: 通算行など kyuiMap にキーがない場合 → 球速から推定
    const veloN = parseFloat(String(pd.velo || ''));
    if (!isNaN(veloN) && veloN > 0) {
      const est = calcKyuiPreStatcast(veloN, idx, pctNum);
      return est !== '' ? Number(est) : null;
    }
    return null;
  };

  const cands = [];
  const chK = kyuiOf(2, pitchData[2]); if (chK !== null) { const s = _kCH(chK);   if (s !== null) cands.push(s); }
  const cuK = kyuiOf(3, pitchData[3]); if (cuK !== null) { const s = _kCU(cuK);   if (s !== null) cands.push(s); }
  const slK = kyuiOf(1, pitchData[1]); if (slK !== null) { const s = _kSLFS(slK); if (s !== null) cands.push(s); }
  const fsK = kyuiOf(6, pitchData[6]); if (fsK !== null) { const s = _kSLFS(fsK); if (s !== null) cands.push(s); }
  if (cands.length === 0) return '';
  return Math.round((v1 + v2 + 2 * Math.max(...cands)) / 3);
}

// ── 能力値列を追加 ────────────────────────────────────────────────────────────
// AY (Col 51) = スタミナ   AZ (Col 52) = 制球   BA (Col 53) = 緩急
// BB (Col 54) = 精神       BC (Col 55) = 奪三振
// BD (Col 56) = 重さ       BE (Col 57) = 対左   BF (Col 58) = 対盗塁
const STAMINA_COL   = 51;
const SEIKYU_COL    = 52;
const KANKYUU_COL   = 53;
const SEISIN_COL    = 54;
const SANSHIN_COL   = 55;
const OMOSA_COL     = 56;
const TAILEFT_COL   = 57;
const TAITOURUI_COL = 58;

async function addAbilityToFile(xlsxPath, showKyuiMap = {}, pitchNameOverrides = {}) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(xlsxPath);
  const ws = wb.worksheets[0];
  const fontSize = ws.getCell(1, 1).font?.size || 11;

  // ヘッダー (Row 1)
  purpleCell(ws.getCell(1, STAMINA_COL),   'スタミナ', fontSize);
  purpleCell(ws.getCell(1, SEIKYU_COL),    '制球',     fontSize);
  purpleCell(ws.getCell(1, KANKYUU_COL),   '緩急',     fontSize);
  purpleCell(ws.getCell(1, SEISIN_COL),    '精神',     fontSize);
  purpleCell(ws.getCell(1, SANSHIN_COL),   '奪三振',   fontSize);
  purpleCell(ws.getCell(1, OMOSA_COL),     '重さ',     fontSize);
  purpleCell(ws.getCell(1, TAILEFT_COL),   '対左',     fontSize);
  purpleCell(ws.getCell(1, TAITOURUI_COL), '対盗塁',   fontSize);

  const dataRows = [];
  const pitchActiveSet = new Set();
  ws.eachRow((row, rn) => {
    if (rn <= 2) return;
    const yr = row.getCell(2).value;
    if (!yr) return;
    const ipRaw = row.getCell(11).value;
    if (ipRaw == null || ipRaw === '' || ipRaw === '--') return;
    const g      = Number(row.getCell(7).value)  || 0;
    const gs     = Number(row.getCell(8).value)  || 0;
    const bb     = Number(row.getCell(15).value) || 0;
    const eraRaw = row.getCell(6).value;
    const hr     = Number(row.getCell(14).value) || 0;
    const so     = Number(row.getCell(16).value) || 0;
    const avgRaw = row.getCell(17).value;
    const vsLRaw = row.getCell(19).value;
    const sb     = Number(row.getCell(20).value) || 0;
    const pk     = Number(row.getCell(21).value) || 0;
    const cs     = Number(row.getCell(22).value) || 0;
    const pitchData = {};
    for (const pg of PITCH_GROUPS) {
      const velo = row.getCell(pg.startCol + 0).value;
      const ba   = row.getCell(pg.startCol + 1).value;
      const slg  = row.getCell(pg.startCol + 2).value;
      const pct  = row.getCell(pg.startCol + 3).value;
      const pctStr = String(pct == null ? '' : pct).trim();
      const pctNum = Number(pctStr);
      if (pctStr && pctStr !== '--' && !isNaN(pctNum) && pctNum > 0) pitchActiveSet.add(pg.idx);
      pitchData[pg.idx] = { velo, ba, slg, pct };
    }
    dataRows.push({ rn, yr, ipRaw, g, gs, bb, eraRaw, hr, so, avgRaw, vsLRaw, sb, pk, cs, pitchData });
  });

  const activePitchList = PITCH_GROUPS.filter(pg => pitchActiveSet.has(pg.idx));

  activePitchList.forEach((pg, i) => {
    const base = PITCH_ABILITY_START_COL + i * 3;
    // pitchNameOverrides に登録がある場合はサブタイプ優先表示名を使用
    const displayName = pitchNameOverrides[pg.idx] ?? pg.name;
    redPurpleCell(ws.getCell(1, base), displayName, fontSize);
    try { ws.mergeCells(1, base, 1, base + 2); } catch {}
    redPurpleCell(ws.getCell(2, base + 0), '球速', fontSize);
    redPurpleCell(ws.getCell(2, base + 1), '球威', fontSize);
    redPurpleCell(ws.getCell(2, base + 2), '割合', fontSize);
  });

  let count = 0;
  for (const { rn, yr, ipRaw, g, gs, bb, eraRaw, hr, so, avgRaw, vsLRaw, sb, pk, cs, pitchData } of dataRows) {
    const ip = parseIP(ipRaw);
    if (!ip) continue;

    const stamina = calcStaminaFromIP(ip, g, gs);
    if (stamina !== '') purpleCell(ws.getCell(rn, STAMINA_COL), stamina, fontSize);

    const seikyu = calcSeikyuFromBB9(bb / ip * 9);
    if (seikyu !== '') purpleCell(ws.getCell(rn, SEIKYU_COL), seikyu, fontSize);

    const eraStr = String(eraRaw || '').trim();

    // パフォーマンスデータ (BA/SLGなし年の球威補正用)
    const perfEra    = parseFloat(eraStr);
    const _avgStr    = String(avgRaw == null ? '' : avgRaw).trim();
    const perfBaa1000 = (_avgStr && _avgStr !== '--') ? Number(_avgStr) : NaN;
    const perfHr9    = ip > 0 ? hr / ip * 9 : NaN;

    // 緩急 (①ERA + ②制球 + ③変化球威力MAX) ÷ 2
    // showKyuiMap[yr] を渡すことで BA/SLG なし年（FanGraphs/推定データ）でも計算可能
    const kankyuu = calcKankyuu(eraStr, seikyu, pitchData, showKyuiMap[yr] || {});
    if (kankyuu !== '') purpleCell(ws.getCell(rn, KANKYUU_COL), kankyuu, fontSize);

    // 精神
    if (eraStr && eraStr !== '--') {
      const seisin = calcSeisinFromERA(Number(eraStr));
      if (seisin !== '') purpleCell(ws.getCell(rn, SEISIN_COL), seisin, fontSize);
    }

    const sanshin = calcSanshinFromK9(so / ip * 9);
    if (sanshin !== '') purpleCell(ws.getCell(rn, SANSHIN_COL), sanshin, fontSize);

    // 被本塁打0の場合はイニング数に応じたデフォルト値を使用
    // (hr9=0 は通常の計算式で範囲外になるため IP帯別に固定値を返す)
    let omosa;
    if (hr === 0 && ip > 0) {
      if      (ip <= 20) omosa = 85;
      else if (ip <= 40) omosa = 90;
      else if (ip <= 60) omosa = 95;
      else               omosa = 100;
    } else {
      omosa = calcOmosaFromHR9(ip > 0 ? hr / ip * 9 : NaN);
    }
    if (omosa !== '') purpleCell(ws.getCell(rn, OMOSA_COL), omosa, fontSize);

    const avgStr = String(avgRaw || '').trim();
    const vsLStr = String(vsLRaw || '').trim();
    if (avgStr && avgStr !== '--' && vsLStr && vsLStr !== '--') {
      const taileft = calcTaiHidariFromDiff(Number(avgStr) - Number(vsLStr));
      if (taileft !== '') purpleCell(ws.getCell(rn, TAILEFT_COL), taileft, fontSize);
    }

    const taitourui = calcTaiTouruiFromSBData(sb, pk, ip, cs);
    if (taitourui !== '') purpleCell(ws.getCell(rn, TAITOURUI_COL), taitourui, fontSize);

    // 球種能力値
    activePitchList.forEach((pg, i) => {
      const pd = pitchData[pg.idx];
      const pctStr = String(pd.pct == null ? '' : pd.pct).trim();
      const pctNum = Number(pctStr);
      if (!pctStr || pctStr === '--' || isNaN(pctNum) || pctNum <= 0) return;

      const veloNum = parseFloat(String(pd.velo || ''));
      const baStr   = String(pd.ba  == null ? '' : pd.ba).trim();
      const slgStr  = String(pd.slg == null ? '' : pd.slg).trim();
      const baNum   = (baStr  && baStr  !== '--') ? Number(baStr)  : NaN;
      const slgNum  = (slgStr && slgStr !== '--') ? Number(slgStr) : NaN;

      const base = PITCH_ABILITY_START_COL + i * 3;

      const kyusoku = calcKyuSoku(veloNum);
      if (kyusoku !== '') redPurpleCell(ws.getCell(rn, base + 0), kyusoku, fontSize);

      const ah = calcAH_pitch(pg.idx, baNum);
      const ai = calcAI_pitch(pg.idx, slgNum);
      let kyui = '';
      if (ah !== '' && ai !== '') {
        const aj = (Number(ah) + Number(ai)) / 2;
        const ak = calcAK_pitch(pg.idx, aj, pctNum);
        kyui = calcKyuI(aj, ak, pctNum);
      } else {
        // BA/SLG なし年: パフォーマンス補正値を算出
        const perfBoost = calcPerfBoost(perfEra, perfBaa1000, perfHr9);
        // フォールバック①: showKyuiMap（The Show / FanGraphs / 推定データ）+ 補正
        const showKyui = showKyuiMap[yr]?.[pg.idx];
        if (showKyui !== undefined) {
          kyui = Math.max(30, Math.min(110, Number(showKyui) + perfBoost));
        } else if (!isNaN(veloNum) && veloNum > 0) {
          // フォールバック②: 通算行など showKyuiMap にキーがない場合 → 球速 + 補正で推定
          const est = calcKyuiPreStatcast(veloNum, pg.idx, pctNum, perfEra, perfBaa1000, perfHr9);
          if (est !== '') kyui = est;
        }
        // ── 球威キャップ（③以降推定年: 投球回数・防御率による上限制約）─────────────────
        // ①②(Baseball Savant 実測) は BA/SLG あり → if(ah!==''&&ai!=='') ブランチを通るため制約なし。
        // ③(推定)/④(FanGraphs)/⑤(Claude+AgingCurve) は BA/SLG='--' → このブランチで制約適用。
        // 通算行はスキップ（yr==='通算' は年度別IPが不明のため）。
        //
        // IP 100-150: ERA≧3.00 → ×0.4, 2.99-2.00 → ×0.5, 1.99-1.00 → ×0.6  (閾値95, ベース95)
        // IP 75-99:   ERA≧2.80 → ×0.4, 2.79-1.80 → ×0.5, 1.79-0.80 → ×0.6  (閾値95, ベース95)
        // IP 35-74:   ERA≧2.50 → ×0.4, 2.49-1.50 → ×0.5, 1.49-0.50 → ×0.6  (閾値95, ベース95)
        // IP 0-34:    ERA≧1.50 → ×0.4, 1.49-0.00 → ×0.5                     (閾値90, ベース90)
        if (kyui !== '' && yr !== '通算') {
          const eraNum = !isNaN(perfEra) ? perfEra : NaN;
          if (!isNaN(eraNum) && eraNum >= 0) {
            const A = Number(kyui);
            let base = 95, threshold = 95, mult = null;

            if (ip >= 100 && ip <= 150) {
              if      (eraNum >= 3.00)              { threshold = 95; mult = 0.4; }
              else if (eraNum >= 2.00)              { threshold = 95; mult = 0.5; }
              else if (eraNum >= 1.00)              { threshold = 95; mult = 0.6; }
            } else if (ip >= 75 && ip < 100) {
              if      (eraNum >= 2.80)              { threshold = 95; mult = 0.4; }
              else if (eraNum >= 1.80)              { threshold = 95; mult = 0.5; }
              else if (eraNum >= 0.80)              { threshold = 95; mult = 0.6; }
            } else if (ip >= 35 && ip < 75) {
              if      (eraNum >= 2.50)              { threshold = 95; mult = 0.4; }
              else if (eraNum >= 1.50)              { threshold = 95; mult = 0.5; }
              else if (eraNum >= 0.50)              { threshold = 95; mult = 0.6; }
            } else if (ip >= 0 && ip < 35) {
              base = 90;
              if      (eraNum >= 1.50)              { threshold = 90; mult = 0.4; }
              else if (eraNum >= 0.00)              { threshold = 90; mult = 0.5; }
            }

            if (mult !== null && A >= threshold) {
              kyui = Math.min(110, base + Math.round((A - base) * mult));
            }
          }
        }
      }
      // ── ナックルボール球威 +5 補正 ───────────────────────────────────────────
      // ナックルボールは速度と球質が無相関（遅くても打ちにくい）ため、算出値に+5する。
      // 判定: fs バケット(idx=6) かつ pitchNameOverrides に 'ナックル' 登録あり
      // ★ 旧判定②「球速 ≤ 83mph」は廃止: スプリット/フォークも 83mph 以下になりうるため
      //   長谷川滋利など 'スプリット' 投手への誤適用が発生した。
      //   真のナックルボーラー（ディッキー 72-77mph / ウェイクフィールド 65-70mph）は
      //   pitchNameOverrides で 'ナックル' が正しく設定されるため①のみで十分。
      if (kyui !== '' && pg.idx === 6) {
        const ovName = pitchNameOverrides[pg.idx] ?? '';
        const isKnuckleball = ovName === 'ナックル';
        if (isKnuckleball) kyui = Math.max(30, Math.min(110, Number(kyui) + 5));
      }
      if (kyui !== '') redPurpleCell(ws.getCell(rn, base + 1), kyui, fontSize);

      redPurpleCell(ws.getCell(rn, base + 2), pctNum, fontSize);
    });

    count++;
  }

  await wb.xlsx.writeFile(xlsxPath);
  return count;
}

// ── チーム略称の正規化テーブル ────────────────────────────────────────────────
// MLB Stats API が返す略称はシーズン・移転により揺れがあるため統一する
const TEAM_ABBR_NORMALIZE = {
  // アメリカンリーグ東
  'TBD': 'TB',  'TBR': 'TB',                    // Devil Rays / Rays 表記揺れ
  'KCR': 'KC',                                    // Royals alt
  'CHW': 'CWS',                                   // White Sox alt
  // アメリカンリーグ西
  'ANA': 'LAA',  'CAL': 'LAA',                   // Angels 旧名
  'OAK': 'ATH',                                   // Athletics（Oakland→Sacramento移転後）
  // ナショナルリーグ東
  'FLA': 'MIA',                                   // Florida Marlins → Miami Marlins
  'MON': 'WSH',  'WSN': 'WSH',                   // Montreal Expos → Washington Nationals
  // ナショナルリーグ西
  'SDP': 'SD',                                    // Padres alt
  'SFG': 'SF',                                    // Giants alt
};
function normalizeTeamAbbr(raw) {
  return TEAM_ABBR_NORMALIZE[raw] || raw;
}

// ── Pitching stats fetch ──────────────────────────────────────────────────────
async function fetchPitchingStats(id, y1, y2) {
  // yearByYear と career を並列取得（career は year list 確定前に先行フェッチ）
  const [yby, careerData] = await Promise.all([
    mlbGet(`https://statsapi.mlb.com/api/v1/people/${id}/stats?stats=yearByYear&group=pitching&sportId=1`),
    mlbGet(`https://statsapi.mlb.com/api/v1/people/${id}/stats?stats=career&group=pitching&sportId=1`),
  ]);

  const allSplits = (yby.stats[0]?.splits || []).filter(s => s.sport?.id === 1);
  const byYear = {};
  for (const s of allSplits) {
    const yr = s.season;
    if (!byYear[yr]) byYear[yr] = [];
    byYear[yr].push(s);
  }
  const years = Object.keys(byYear).filter(y => +y >= y1 && +y <= y2).sort();
  if (!years.length) throw new Error(`ID ${id} に ${y1}〜${y2} の投手成績データがありません`);

  const basic = {};
  for (const yr of years) {
    const rows = byYear[yr];
    const row  = rows.find(r => !r.team) || rows[0];
    let teamStr;
    if (rows.length > 1) {
      const named   = rows.filter(r => r.team);
      const primary = named.reduce((a, b) => (a.stat.gamesPitched >= b.stat.gamesPitched ? a : b));
      teamStr = normalizeTeamAbbr(primary.team?.abbreviation || primary.team?.name?.slice(0,3)?.toUpperCase() || '???') + named.length;
    } else {
      teamStr = normalizeTeamAbbr(row.team?.abbreviation || row.team?.name?.slice(0,3)?.toUpperCase() || '???');
    }
    const st = row.stat;
    basic[yr] = {
      team: teamStr,
      w: st.wins, l: st.losses, era: st.era,
      g: st.gamesPitched, gs: st.gamesStarted,
      hld: st.holds || 0, sv: st.saves,
      ip: st.inningsPitched,
      h: st.hits, er: st.earnedRuns, hr: st.homeRuns,
      bb: (st.baseOnBalls || 0) + (st.hitBatsmen || 0),
      so: st.strikeOuts,
      avg: st.avg, whip: st.whip,
      sb: st.stolenBases || 0, pk: st.pickoffs || 0, cs: st.caughtStealing || 0,
    };
  }

  const cs2 = careerData.stats[0]?.splits[0]?.stat || {};
  basic['通算'] = {
    team: basic[years[years.length - 1]]?.team?.replace(/\d+$/, '') || '---',
    w: cs2.wins, l: cs2.losses, era: cs2.era,
    g: cs2.gamesPitched, gs: cs2.gamesStarted,
    hld: cs2.holds || 0, sv: cs2.saves,
    ip: cs2.inningsPitched,
    h: cs2.hits, er: cs2.earnedRuns, hr: cs2.homeRuns,
    bb: (cs2.baseOnBalls || 0) + (cs2.hitBatsmen || 0),
    so: cs2.strikeOuts,
    avg: cs2.avg, whip: cs2.whip,
    sb: cs2.stolenBases || 0, pk: cs2.pickoffs || 0, cs: cs2.caughtStealing || 0,
  };

  // vsLeft（年別 + 通算）をすべて並列取得
  const vsLeftByYear = {};
  await Promise.all([
    ...years.map(async yr => {
      try {
        const vl = await mlbGet(
          `https://statsapi.mlb.com/api/v1/people/${id}/stats?stats=statSplits&group=pitching&sportId=1&sitCodes=vl&season=${yr}`
        );
        vsLeftByYear[yr] = vl.stats[0]?.splits[0]?.stat?.avg || '--';
      } catch { vsLeftByYear[yr] = '--'; }
    }),
    (async () => {
      try {
        const carVL = await mlbGet(
          `https://statsapi.mlb.com/api/v1/people/${id}/stats?stats=careerStatSplits&group=pitching&sportId=1&sitCodes=vl`
        );
        vsLeftByYear['通算'] = carVL.stats[0]?.splits[0]?.stat?.avg || '--';
      } catch { vsLeftByYear['通算'] = '--'; }
    })(),
  ]);

  return { years, basic, vsLeftByYear };
}

// ── Pitch type config ─────────────────────────────────────────────────────────
const PITCH_KEYS     = ['ff', 'sl', 'ch', 'cu', 'fc', 'si', 'fs'];
const PITCH_NAMES_JA = ['4シーム', 'スライダー', 'チェンジアップ', 'カーブ', 'カット', 'シンカー', 'スプリット'];

// サブタイプ → 日本語表示名（デフォルト PITCH_NAMES_JA を上書きする球種のみ定義）
const SUBTYPE_DISPLAY_JA = {
  // sl バケット
  'Sweeper':          'スイーパー',
  'Slurve':           'スライダー',
  'Hard Slider':      'スライダー',
  // ch バケット
  'Circle Change':    'チェンジアップ',
  'Vulcan Change':    'チェンジアップ',
  'Eephus':           'イーファス',
  // cu バケット
  'Knuckle Curve':    'ナックルカーブ',
  'Knuckle-Curve':    'ナックルカーブ',
  'Slow Curve':       'スローカーブ',
  '12-6 Curve':       'カーブ',
  'Power Curve':      'カーブ',
  // si バケット
  'Sinker':           'シンカー',
  'Two-Seam Fastball':'ツーシーム',
  '2-Seam Fastball':  'ツーシーム',
  'Two Seamer':       'ツーシーム',
  // fs バケット
  'Knuckleball':      'ナックル',
  'Forkball':         'フォーク',
  'Splitter':         'スプリット',
  'Split-Finger':     'スプリット',
  'Split Finger':     'スプリット',
};

// Baseball Savant pitch name → 球種キー
const PITCH_MAP_P    = {
  '4-Seam Fastball': 'ff', '4-seam Fastball': 'ff', 'Four-Seam Fastball': 'ff',
  'Four Seamer': 'ff', 'Four-Seamer': 'ff', '4-Seamer': 'ff', '4 Seamer': 'ff',
  'Fastball': 'ff',        // 初期Statcast年代の汎用分類
  'Riding Fastball': 'ff', 'Rising Fastball': 'ff',
  // sl: スライダー系（スイーパー・スラーブ含む）
  'Slider': 'sl', 'Sweeper': 'sl', 'Hard Slider': 'sl', 'Slurve': 'sl',
  // ch: チェンジアップ系（イーファス含む）
  'Changeup': 'ch', 'Change-up': 'ch', 'Circle Change': 'ch', 'Eephus': 'ch',
  // cu: カーブ系（ナックルカーブ・スローカーブ含む）
  'Curveball': 'cu', 'Knuckle Curve': 'cu', 'Slow Curve': 'cu',
  '12-6 Curve': 'cu', 'Power Curve': 'cu',
  'Cutter': 'fc',
  // si: シンカー系（ツーシーム合算）
  'Sinker': 'si', 'Two-Seam Fastball': 'si', '2-Seam Fastball': 'si', 'Two Seamer': 'si',
  // fs: スプリット系（ナックルボール・フォーク含む）
  'Split-Finger': 'fs', 'Splitter': 'fs', 'Split Finger': 'fs',
  'Forkball': 'fs', 'Knuckleball': 'fs',
};

// ※ brooksbaseball.net は廃止のため PITCH_MAP_B 削除済み (2025-05)

// Baseball Savant JSON API pitch_type コード → 球種キー
const PITCH_TYPE_JSON = {
  'FF': 'ff', 'FA': 'ff',                    // 4-Seam / generic
  'FT': 'si',                                 // Two-Seam → Sinker系に統合
  'SI': 'si',                                 // Sinker
  'SL': 'sl', 'ST': 'sl', 'SV': 'sl',       // Slider / Sweeper / Slurve
  'CH': 'ch', 'SC': 'ch', 'EP': 'ch',       // Changeup / Screwball / Eephus
  'CU': 'cu', 'KC': 'cu', 'CS': 'cu',       // Curveball / Knuckle-curve / Slow Curve
  'FC': 'fc',                                 // Cutter
  'FS': 'fs', 'FO': 'fs', 'KN': 'fs',       // Split-finger / Forkball / Knuckleball
};

const emptyPitchP = () => Object.fromEntries(
  PITCH_KEYS.map(k => [k, { velo: '--', ba: '--', slg: '--', pct: '--' }])
);

// ── ユーティリティ ──────────────────────────────────────────────────────────
const sleep = ms => new Promise(r => setTimeout(r, ms));

// ── MLB The Show pitch data (pre-2017 年用) ──────────────────────────────────
// MLB The Show カード球種名 → pitch idx (0=FF,1=SL,2=CH,3=CU,4=FC,5=SI,6=FS)
const PITCH_MAP_SHOW = {
  '4-Seam Fastball': 0, 'Fastball': 0, 'Rising Fastball': 0, 'Running Fastball': 0,
  'Two-Seam Fastball': 5, 'Sinker': 5,
  'Slider': 1, 'Sweeper': 1, 'Slurve': 1, 'Hard Slider': 1,
  'Changeup': 2, 'Circle Change': 2, 'Vulcan Change': 2, 'Eephus': 2,
  'Curveball': 3, '12-6 Curve': 3, 'Slow Curve': 3, 'Knuckle-Curve': 3, 'Power Curve': 3,
  'Cutter': 4,
  'Splitter': 6, 'Forkball': 6, 'Split-Finger': 6, 'Split Finger': 6, 'Knuckleball': 6,
};

// ── Wikipedia 球種テキスト解析パターン ──────────────────────────────────────────
// 英語テキストから球種を識別するためのキーワードマップ。
// 登場順 → primaryPitch 判定に使用（最初に言及される球種がその投手のメイン球種）。
const WIKI_PITCH_PATTERNS = [
  { key: 'ff', words: ['four-seam fastball','four seam fastball','4-seam fastball','4-seamer','four-seamer','straight fastball','rising fastball'] },
  { key: 'si', words: ['sinker','sinking fastball','two-seam fastball','two seam fastball','2-seam fastball','2-seamer','two-seamer','tailing fastball','sinkball','sinking two-seamer'] },
  { key: 'sl', words: ['slider','sweeper','hard slider','sharp slider'] },
  { key: 'ch', words: ['changeup','change-up','change up','circle changeup','circle change','palmball','screwball','fading changeup'] },
  { key: 'cu', words: ['curveball','curve ball','12-6 curveball','12-6 curve','overhand curve','knuckle curve','biting curveball','power curve','big curve','downer'] },
  { key: 'fc', words: ['cutter','cut fastball','cutting fastball','cut-fastball'] },
  { key: 'fs', words: ['splitter','split-finger fastball','split finger fastball','split-fingered fastball','forkball','fork ball','knuckleball','knuckle ball','knuckler'] },
];

/**
 * 英語テキストから球種・球速情報を抽出する。
 * @param {string} text  Wikipedia などから取得した英語テキスト（8000文字以内推奨）
 * @returns {{ pitchKeys:string[], primaryKey:string|null, pitchCounts:{[key]:number}, veloMentions:number[] } | null}
 */
function parsePitchProfile(text) {
  if (!text || typeof text !== 'string') return null;
  const lower = text.toLowerCase();

  // 球種の初出位置と出現回数を記録
  const pitchOrder  = []; // { key, firstPos } 初出位置順
  const pitchCounts = {}; // key → 出現回数

  for (const { key, words } of WIKI_PITCH_PATTERNS) {
    let firstPos = Infinity;
    let count = 0;
    for (const word of words) {
      let pos = 0;
      while ((pos = lower.indexOf(word, pos)) !== -1) {
        count++;
        if (pos < firstPos) firstPos = pos;
        pos += word.length;
      }
    }
    if (count > 0) {
      pitchCounts[key] = (pitchCounts[key] || 0) + count;
      if (!pitchOrder.find(p => p.key === key)) pitchOrder.push({ key, firstPos });
    }
  }

  pitchOrder.sort((a, b) => a.firstPos - b.firstPos);
  const pitchKeys  = pitchOrder.map(p => p.key);
  const primaryKey = pitchKeys[0] ?? null;

  // 球速の抽出: 80-106mph の数値（投手の実用球速範囲）
  const veloMentions = [];
  for (const m of text.matchAll(/(\d{2,3})\s*(?:[-–to]+\s*(\d{2,3}))?\s*mph/gi)) {
    const v1 = parseInt(m[1]), v2 = m[2] ? parseInt(m[2]) : null;
    if (v1 >= 80 && v1 <= 106) veloMentions.push(v1);
    if (v2 && v2 >= 80 && v2 <= 106) veloMentions.push(v2);
  }

  if (pitchKeys.length === 0 && veloMentions.length === 0) return null;
  return { pitchKeys, primaryKey, pitchCounts, veloMentions };
}

/**
 * Wikipedia API を使って投手の球種・球速プロファイルを無料で取得する。
 * Anthropic API 不要。英語選手名から Wikipedia 記事を検索・解析する。
 * @param {string} searchName  英語選手名 (例: "Randy Johnson", "Scot Shields")
 * @returns {{ pitchKeys, primaryKey, pitchCounts, veloMentions, pageTitle } | null}
 */
async function fetchWikipediaPitchProfile(searchName) {
  if (!searchName || !searchName.trim()) return null;

  const wikiGet = url => new Promise((resolve, reject) => {
    const mod = url.startsWith('https') ? require('https') : require('http');
    mod.get(url, {
      headers: {
        'User-Agent': 'MLB-PitchTool/1.0 (Node.js; baseball stats research)',
        'Accept': 'application/json',
      },
    }, res => {
      // Wikipedia は 301 リダイレクトする場合がある
      if (res.statusCode === 301 || res.statusCode === 302) {
        const loc = res.headers.location;
        if (loc) return wikiGet(loc).then(resolve).catch(reject);
        return reject(new Error('redirect without location'));
      }
      let buf = '';
      res.on('data', c => buf += c);
      res.on('end', () => {
        try { resolve(JSON.parse(buf)); } catch (e) { reject(e); }
      });
    }).on('error', reject);
  });

  try {
    // ── Step 1: opensearch で記事タイトル候補を検索 ──
    const q1 = encodeURIComponent(searchName + ' baseball pitcher');
    const searchRes = await wikiGet(
      `https://en.wikipedia.org/w/api.php?action=opensearch&search=${q1}&limit=5&namespace=0&format=json&origin=*`
    ).catch(() => null);

    let titles = searchRes?.[1] ?? [];
    let descs  = searchRes?.[2] ?? [];

    // ヒットなし → " baseball pitcher" なしで再検索
    if (!titles.length) {
      const q2 = encodeURIComponent(searchName);
      const fb = await wikiGet(
        `https://en.wikipedia.org/w/api.php?action=opensearch&search=${q2}&limit=5&namespace=0&format=json&origin=*`
      ).catch(() => null);
      titles = fb?.[1] ?? [];
      descs  = fb?.[2] ?? [];
    }
    if (!titles.length) return null;

    // baseball / pitcher に関連するタイトルを優先
    let bestTitle = titles[0];
    for (let i = 0; i < titles.length; i++) {
      const d = (descs[i] || '').toLowerCase();
      if (d.includes('pitcher') || d.includes('baseball') || d.includes('mlb')) {
        bestTitle = titles[i]; break;
      }
    }

    // ── Step 2: 記事テキストを取得（最大 8000 文字に制限）──
    const qt = encodeURIComponent(bestTitle);
    const extractRes = await wikiGet(
      `https://en.wikipedia.org/w/api.php?action=query&titles=${qt}&prop=extracts&explaintext=true&format=json&origin=*`
    ).catch(() => null);

    const pages = extractRes?.query?.pages;
    if (!pages) return null;
    const page = Object.values(pages)[0];
    if (!page || page.missing !== undefined) return null;

    const fullText = (page.extract || '').slice(0, 8000);
    if (!fullText) return null;

    const profile = parsePitchProfile(fullText);
    if (!profile) return null;
    return { ...profile, pageTitle: page.title };
  } catch {
    return null;
  }
}

// 球威計算 (MLB The Show ゲームデータ基準)
// speed: 実際mph(100mph=100%), control/movement: 0〜99スケール
// 平均>=90%: 90+(avg-90)*2 → 90%=球威90, 95%=球威100, 100%=球威110
// 平均<90% : (speedPct+movementPct)/2 を直接球威値として使用
function calcKyuiFromShow(speed, control, movement) {
  if (!speed || !control || !movement) return '';
  const speedPct    = Math.min(speed / 100 * 100, 105);
  const controlPct  = control / 99 * 100;
  const movementPct = movement / 99 * 100;
  const avg3 = (speedPct + controlPct + movementPct) / 3;
  if (avg3 >= 90) return Math.min(Math.round(90 + (avg3 - 90) * 2), 110);
  return Math.round((speedPct + movementPct) / 2);
}

// ── Pre-Statcast (pre-2008) 球威推定 ─────────────────────────────────────────
// 球速・球種・投球割合から BA/SLG を推計し、既存の calcAH/AI/AK/KyuI ロジックを適用。
// ベースライン: 参考選手データ（ボックスバーガー/レスター/リンスカム/コール/ウィリアムズ/
//              スクーバル/田中/シース/コービン）から逆算した「平均的MLB投手」の想定値。
// [平均球速(mph), 平均BA(×1000), 平均SLG(×1000), BA変化量/mph, SLG変化量/mph]
const PRE08_PITCH_BASELINES = [
  [91, 265, 420, 10, 15],  // 0: FF  (4-seam)
  [84, 220, 350, 10, 15],  // 1: SL  (slider) ※ハードスライダー基準: 低BA/SLGに調整
  [81, 260, 400,  8, 12],  // 2: CH  (changeup)
  [76, 240, 375,  7, 10],  // 3: CU  (curve)
  [86, 265, 415,  8, 12],  // 4: FC  (cutter)
  [89, 265, 400, 10, 15],  // 5: SI  (sinker)
  [83, 220, 340,  8, 12],  // 6: FS  (split)
];

// ── 変化量推定ボーナス ──────────────────────────────────────────────────────────
// Baseball Savant 実測がない年において、PRE08_PITCH_BASELINES の速度→球威換算では
// 捉えられない「球種固有の変化量・球質」が球威に与える寄与を補正する固定値。
// FF/SI は球速依存度が高いためボーナス小。CU/CH/FS は変化量依存度が高いため大きい値。
// キャリブレーション参考:
//   CU +5: 良いカーブは速度が低くても大きな縦落差で球威を発揮（コービン2.24 ERA的水準）
//   CH +4: アームサイドのフェードと緩急が球威に寄与（サンタナ / チェンジアップ系エース）
//   FS +4: 急激な落差（スプリットフィンガーの鋭い変化）
//   SL +3: 横の変化・スイープが球威に寄与（縦カーブほど速度から予測しにくくはない）
//   SI +2: グラウンドボール誘発（フォーシームより変化が少し多い）
//   FC +1: 微細なカット成分（ほぼ速度依存）
//   FF  0: 球速依存のためボーナスなし
const PITCH_MOVEMENT_BONUS = [
  0,  // 0: FF  - 球速依存
  3,  // 1: SL  - 横変化
  4,  // 2: CH  - フェード＋緩急
  5,  // 3: CU  - 縦落差（最も速度から推定しにくい）
  1,  // 4: FC  - 微細なカット
  2,  // 5: SI  - グラウンドボール変化
  4,  // 6: FS  - 急激な落差
];

/**
 * 緩急差ボーナス: 同年の主速球との球速差が大きいほど変化球はタイミングを外しやすい。
 * 速球（ff/si）には適用しない（緩急差の「与え手」のため）。
 * @param {number} speed    この球種の球速 (mph)
 * @param {number} idx      球種インデックス (0-6)
 * @param {number} ffSpeed  同年の主速球速度 (mph)
 * @returns {number} 緩急差ボーナス (0〜+7)
 * キャリブレーション参考 (CH idx=2, scale=1.2):
 *   diff 10mph → raw1 × 1.2 → +1  (平均的ピッチャーの速球-CH差)
 *   diff 15mph → raw2 × 1.2 → +2  (Santana 93-78mph)
 *   diff 20mph → raw4 × 1.2 → +5  (高津 82-62mph)
 *   diff 25mph → raw6 × 1.2 → +7  (Wake/Dickey型ナックル)
 */
function calcKakkyoSaBonus(speed, idx, ffSpeed) {
  if (isNaN(ffSpeed) || ffSpeed <= 0 || isNaN(speed) || speed <= 0) return 0;
  // idx 別スケール: CH が最大（緩急依存度最高）、SI/FF は与え手のため 0
  const scale = [0.0, 0.6, 1.2, 0.8, 0.3, 0.0, 1.0];
  if ((scale[idx] ?? 0) === 0) return 0;
  const diff = ffSpeed - speed;
  if (diff < 8) return 0;
  // diff  8-12mph → raw 1
  // diff 13-17mph → raw 2
  // diff 18-22mph → raw 4
  // diff   23mph+ → raw 6
  const raw = diff >= 23 ? 6 : diff >= 18 ? 4 : diff >= 13 ? 2 : 1;
  return Math.round(raw * scale[idx]);
}

/**
 * rawPitch の単一年データから主速球速度 (ff/si のうち投球割合が高い方) を返す。
 * @param {Object} rawPitchYr  rawPitch[yr] の値
 * @returns {number} 主速球速度 (mph)。取得できない場合は NaN。
 */
function getPrimaryFfMph(rawPitchYr) {
  if (!rawPitchYr) return NaN;
  let best = NaN, bestPct = 0;
  for (const fk of ['ff', 'si']) {
    const d = rawPitchYr[fk];
    if (!d || d.velo === '--') continue;
    const v = parseFloat(d.velo);
    const p = parseFloat(String(d.pct ?? '').replace('%', ''));
    if (isNaN(v) || v <= 0 || isNaN(p) || p <= 0) continue;
    if (p > bestPct) { bestPct = p; best = v; }
  }
  return best;
}

/**
 * パフォーマンス補正値を計算 (球威へのボーナス/ペナルティ)
 * ERA・被打率・HR/9 から投手の実際の支配力を評価する。
 * 速い球でも成績が良ければ追加補正、悪ければ減点。
 * @param {number} era      防御率 (例: 2.60)
 * @param {number} baa1000  被打率×1000の整数 (例: 197 = .197)
 * @param {number} hr9      9イニング換算被本塁打数
 * @returns {number} 球威補正値 (整数、-10〜+10)
 * キャリブレーション例:
 *   RaジョンソN 2004 (ERA2.60, BAA.197, HR/9 0.49) → +10
 *   平均的投手   2004 (ERA4.00, BAA.265, HR/9 1.00) → -5
 *   不振投手         (ERA5.50, BAA.310, HR/9 1.80) → -10 (上限)
 */
function calcPerfBoost(era, baa1000, hr9) {
  let boost = 0;
  if (!isNaN(era) && era > 0) {
    // ERA 2.0→+6, 2.5→+4, 3.0→+2, 3.5→0, 4.0→-2, 4.5→-4, 5.5→-8
    boost += Math.round((3.5 - era) * 4);
  }
  if (!isNaN(baa1000) && baa1000 > 0) {
    // BAA .200→+5, .225→+2, .250→0, .275→-2, .300→-5
    boost += Math.round((250 - baa1000) / 10);
  }
  if (!isNaN(hr9) && hr9 >= 0) {
    // HR/9 0.3→+2, 0.6→+1, 0.8→0, 1.2→-2, 1.5→-3
    boost += Math.round((0.8 - hr9) * 4);
  }
  return Math.max(-10, Math.min(10, boost));
}

/**
 * Pre-Statcast 球威推定
 * @param {number} speed    実際の球速 (mph)
 * @param {number} idx      球種インデックス (0=FF,1=SL,2=CH,3=CU,4=FC,5=SI,6=FS)
 * @param {number} pctNum   投球割合 (0〜100の整数)
 * @param {number} [era]    防御率 (省略可)
 * @param {number} [baa1000] 被打率×1000整数 (省略可)
 * @param {number} [hr9]    被本塁打/9回 (省略可)
 * @returns {number|string} 球威 (整数) または '' (計算不能)
 */
function calcKyuiPreStatcast(speed, idx, pctNum, era = NaN, baa1000 = NaN, hr9 = NaN) {
  if (!speed || isNaN(speed)) return '';

  // ── ナックルボール特別処理 ──────────────────────────────────────────────────
  // ナックルボール(idx=6 かつ speed≤83mph) は球速と球質が無相関。
  // 速度依存の BA/SLG 推計を使わず、ERA/BAA/HR9 の成績ベースで球威を決定する。
  // ベース70 = 平均的ナックルボーラー、最良で+10→80、最悪で-10→60
  if (idx === 6 && speed > 0 && speed <= 83) {
    const boost = calcPerfBoost(era, baa1000, hr9);
    return Math.max(40, Math.min(110, 70 + boost));
  }

  // ── 超低速チェンジアップ特別処理（高津臣吾型・日本式シンカー/スクリューボール）──
  // CH(idx=2) で球速 ≤ 70mph(≈112km/h) の場合、通常の速度→BA/SLG推計は機能しない。
  // PRE08_PITCH_BASELINESのベースラインは81mphのため、63mphでは
  //   spd=-18 → estBA=380(上限), estSLG=615 → 「極めて打ちやすい球」と誤判定する。
  // このタイプは「速球との緩急差」が武器であり、球速と被打率は無相関。
  // ナックルボールと同様に ERA/BAA/HR9 の成績ベースで球威を決定する。
  // ベース73 = 「緩急差で機能する変化球」水準、最良で+10→83、最悪で-10→63
  if (idx === 2 && speed > 0 && speed <= 70) {
    const boost = calcPerfBoost(era, baa1000, hr9);
    return Math.max(40, Math.min(110, 73 + boost));
  }

  const bl = PRE08_PITCH_BASELINES[idx] || PRE08_PITCH_BASELINES[0];
  const spd = speed - bl[0];
  // 球速偏差から BA/SLG を推計（速いほど低 BA/SLG → 高球威）
  const estBA  = Math.max(80,  Math.min(380, Math.round(bl[1] - spd * bl[3])));
  const estSLG = Math.max(120, Math.min(620, Math.round(bl[2] - spd * bl[4])));
  const ah = calcAH_pitch(idx, estBA);
  const ai = calcAI_pitch(idx, estSLG);
  const base = (ah === '' || ai === '')
    ? Math.max(30, Math.min(110, Math.round(75 + spd * 2)))
    : (() => {
        const aj = (Number(ah) + Number(ai)) / 2;
        const ak = calcAK_pitch(idx, aj, pctNum || 20);
        const kyui = calcKyuI(aj, ak, pctNum || 20);
        return kyui !== '' ? Number(kyui) : Math.max(30, Math.min(110, Math.round(75 + spd * 2)));
      })();
  // パフォーマンス補正: ERA/被打率/HR9 が渡された場合に適用
  const boost = calcPerfBoost(era, baa1000, hr9);
  return Math.max(30, Math.min(110, base + boost));
}

// 球種数 → 推定投球割合
function estimateShowUsagePct(n) {
  const tables = { 1:[100], 2:[62,38], 3:[50,30,20], 4:[42,28,18,12], 5:[35,25,20,12,8] };
  return tables[Math.min(n, 5)] || tables[5];
}

// ── ODT分析ドキュメント プロファイル (優先度③) ────────────────────────────────
// MLB_PitchAnalysis_3Players.odt の年度別解析データをハードコード。
// 優先度: ① Savant > ② Brooks(廃止) > ③ このプロファイル > ④ FanGraphs > ⑤ Claude
//
// pitchMaxKmh[key] = 0  : その球種は投球しない → rawPitch から削除
// pitchMaxKmh[key] = N  : 最大球速キャップ (FGブーストで超過した場合に上限適用)
// phases              : 年度区間別の平均球速 (km/h) ※ km/h → mph 変換: (kmh-4)/1.6
// yearPcts            : FanGraphs未取得年の年度別投球割合 (合計≈100%)
// 各プロファイルは複数名義で参照できるよう定数として外出し
const _SHIELDS_PROFILE = {
  // ODT §2: シンカー/2シームが最大の武器。フォーシームは投げない。
  // PITCHf/x実測(2007-2010)は si≈92mph(148km/h) が正確値。
  // FanGraphsブースト(+3~+6mph)を適用すると 157-161km/h になるバグを pitchMaxKmh でキャップ。
  pitchMaxKmh: {
    ff: 0,   // フォーシームなし（ODT明記）
    si: 153, // シンカー最大153km/h (≈95mph)
    ch: 138, // チェンジアップ最大138km/h
    cu: 140, // カーブ/スラーブ最大140km/h
    fc: 0,   // カットボールなし
    sl: 143, // スライダー最大143km/h
    fs: 0,   // スプリットなし
  },
  // 年度区間別の平均球速 (km/h) ― FanGraphs未取得年(2001-2006)に使用
  phases: [
    { from: 2001, to: 2006, si: 150, cu: 137, sl: 140, ch: 136 },
    { from: 2007, to: 2008, si: 148, cu: 135, sl: 140, ch: 136 },
    { from: 2009, to: 2010, si: 148, cu: 130, sl: 141, ch: 138 },
  ],
  // FanGraphs未取得年の年度別投球割合
  yearPcts: {
    '2001': { si: 65, cu: 20, sl: 10, ch: 5 },
    '2002': { si: 63, cu: 21, sl: 11, ch: 5 },
    '2003': { si: 62, cu: 22, sl: 11, ch: 5 },
    '2004': { si: 60, cu: 23, sl: 12, ch: 5 },
    '2005': { si: 58, cu: 25, sl: 12, ch: 5 },
    '2006': { si: 57, cu: 26, sl: 12, ch: 5 },
    // 2007-2010: FanGraphs/Savant実測あり → yearPcts に含めない (velocity capのみ適用)
  },
};

const _TAKATSU_PROFILE = {
  // ODT §1: サイドスロー。日本式「シンカー」はMLB分類でChangeup(ch)。
  // ffは130-138km/h(81-86mph)と遅く、FGブースト後に143-151km/hへ膨張するため要キャップ。
  // si/fc/fsはODT明記で投球なし。ch上限110km/h(100km/h未満の超スローボールも混在)。
  //
  // overrideFgPitch: FanGraphs実測年でもODTの球種分類を優先する。
  // 理由: FanGraphs BIS の「FB→SI再分類ロジック(≦90.5mph かつ >40%)」が高津の
  //   遅いサイドスローffをsiに誤再分類し、その後 pitchMaxKmh.si=0 で削除 → ff消失する。
  //   PITCHf/x以前(2004-2005)のBISデータは球種分類精度が低いため、ODT優先が適切。
  overrideFgPitch: true,
  pitchMaxKmh: {
    ff: 138, // ファストボール(サイドスロー)最大138km/h（ODT: 130-138、"140km/h未満"）
    si: 0,   // MLB式シンカーなし（日本式シンカー=Changeupに分類済み）
    ch: 110, // チェンジアップ(=日本式シンカー)最大110km/h
    cu: 110, // カーブ最大110km/h
    fc: 0,   // カットボールなし
    sl: 125, // スライダー最大125km/h
    fs: 0,   // スプリットなし
  },
  // 年度区間別の平均球速 (km/h) ― FanGraphsが球種データを持たない場合に使用
  phases: [
    { from: 2004, to: 2004, ff: 134, ch: 105, sl: 120, cu: 105 },
    { from: 2005, to: 2005, ff: 131, ch: 101, sl: 118, cu: 101 }, // 2005は若干球速低下
  ],
  // 年度別投球割合 (FanGraphs BIS に pitch mix データがない年に使用)
  yearPcts: {
    '2004': { ff: 35, ch: 35, sl: 25, cu: 5 },
    '2005': { ff: 35, ch: 32, sl: 25, cu: 5 },
  },
};

const DOCX_PLAYER_PROFILES = {
  // Scot Shields
  'Scot Shields': _SHIELDS_PROFILE,
  // 高津臣吾 (名義ゆれ対応: 日本語名・英語表記2種)
  '高津臣吾':      _TAKATSU_PROFILE,
  'Shingo Takatsu': _TAKATSU_PROFILE,
  'Shinji Takatsu': _TAKATSU_PROFILE,  // ODT表記
};

// ── FanGraphs pitch data (2002+、API key不要) ────────────────────────────────
// BIS era (pre-2008): FB%1/FBv, SL%/SLv, CH%/CHv, CB%/CBv, CT%/CTv, SF%/SFv, KN%/KNv
// PITCHf/x era (2008+): pfxFA%/pfxvFA, pfxSL%/pfxvSL, ..., pfxKN%/pfxvKN
// pct は 0〜1 の小数（0.51 = 51%）
const FG_BIS_PITCH_MAP = [
  { pct: 'FB%1', velo: 'FBv',  idx: 0 },
  { pct: 'SL%',  velo: 'SLv',  idx: 1 },
  { pct: 'CH%',  velo: 'CHv',  idx: 2 },
  { pct: 'CB%',  velo: 'CBv',  idx: 3 },
  { pct: 'CT%',  velo: 'CTv',  idx: 4 },
  { pct: 'SF%',  velo: 'SFv',  idx: 6 },
  { pct: 'KN%',  velo: 'KNv',  idx: 6 },  // Knuckleball (BIS era)
];
const FG_PFX_PITCH_MAP = [
  { pct: 'pfxFA%', velo: 'pfxvFA', idx: 0 },
  { pct: 'pfxSL%', velo: 'pfxvSL', idx: 1 },
  { pct: 'pfxST%', velo: 'pfxvST', idx: 1 },  // Sweeper → SL
  { pct: 'pfxCH%', velo: 'pfxvCH', idx: 2 },
  { pct: 'pfxCU%', velo: 'pfxvCU', idx: 3 },
  { pct: 'pfxKC%', velo: 'pfxvKC', idx: 3 },  // Knuckle-curve → CU
  { pct: 'pfxFC%', velo: 'pfxvFC', idx: 4 },
  { pct: 'pfxSI%', velo: 'pfxvSI', idx: 5 },
  { pct: 'pfxFT%', velo: 'pfxvFT', idx: 5 },  // 2-seam → SI
  { pct: 'pfxFS%', velo: 'pfxvFS', idx: 6 },
  { pct: 'pfxFO%', velo: 'pfxvFO', idx: 6 },  // Forkball → FS
  { pct: 'pfxKN%', velo: 'pfxvKN', idx: 6 },  // Knuckleball (PITCHf/x era, 2008+)
];

function fgGet(url) {
  return new Promise((resolve, reject) => {
    https.get(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'application/json',
        'Referer': 'https://www.fangraphs.com/',
      }
    }, res => {
      let buf = '';
      res.on('data', c => buf += c);
      res.on('end', () => {
        try { resolve(JSON.parse(buf)); }
        catch (e) { reject(new Error('FanGraphs parse error: ' + buf.slice(0, 120))); }
      });
    }).on('error', reject);
  });
}

// FanGraphs 選手検索 → FG numeric ID を返す
async function fetchFanGraphsId(playerName) {
  if (!playerName) return null;
  const lastName = playerName.trim().split(/\s+/).pop();
  try {
    const data = await fgGet(`https://www.fangraphs.com/api/search/players/?search=${encodeURIComponent(lastName)}`);
    const hits  = Array.isArray(data) ? data : (data.hits || []);
    const parts = playerName.trim().toLowerCase().split(/\s+/);
    const match = hits.find(h => {
      const levels = Array.isArray(h.level) ? h.level : [h.level || ''];
      if (!levels.includes('mlb')) return false;
      const hName = (h.name || '').toLowerCase();
      return parts.every(p => hName.includes(p));
    });
    if (match && match.id) {
      console.log(`[FanGraphs] 発見: ${match.name} (ID: ${match.id})`);
      return String(match.id);
    }
  } catch (e) {
    console.log(`[FanGraphs] 検索エラー: ${e.message}`);
  }
  return null;
}

// FanGraphs の1年分データを rawPitch / showKyuiMap に反映
function applyFgRow(row, yr, rawPitch, showKyuiMap, knuckleFsYears) {
  const usePfx = FG_PFX_PITCH_MAP.some(m => {
    const v = parseFloat(row[m.pct]);
    return !isNaN(v) && v > 0;
  });
  const map = usePfx ? FG_PFX_PITCH_MAP : FG_BIS_PITCH_MAP;

  // ── ナックルボール検出（FanGraphs KN% 列が実際に使われた年のみ） ──────────
  // 球速ではなく列名で判定する。KN%/pfxKN% が > 0 の場合のみ knuckleFsYears に追加。
  // ※ SF%（スプリット）が偶然 83mph 以下でも絶対にナックルとは判定しない。
  if (knuckleFsYears) {
    const knPct = usePfx ? parseFloat(row['pfxKN%']) : parseFloat(row['KN%']);
    if (!isNaN(knPct) && knPct > 0) knuckleFsYears.add(yr);
  }

  // idx ごとに集計（pct 加算、velo は加重平均）
  const agg = {};
  for (const m of map) {
    const pctRaw  = parseFloat(row[m.pct]);
    const veloRaw = parseFloat(row[m.velo]);
    if (isNaN(pctRaw) || pctRaw <= 0) continue;
    if (!agg[m.idx]) agg[m.idx] = { pctSum: 0, veloNum: 0, veloDen: 0 };
    agg[m.idx].pctSum  += pctRaw;
    if (!isNaN(veloRaw) && veloRaw > 0) {
      agg[m.idx].veloNum += pctRaw * veloRaw;
      agg[m.idx].veloDen += pctRaw;
    }
  }

  // BIS era: FB%1 は4シームと2シーム/シンカーを含む汎用バケット。
  // 球速≦90.5mph かつ主要球種(>40%)の場合はシンカー系(idx:5)として再分類。
  if (agg[0] && !agg[5]) {
    const veloAvg = agg[0].veloDen > 0 ? agg[0].veloNum / agg[0].veloDen : 0;
    const pctPct  = agg[0].pctSum * 100;
    if (veloAvg > 0 && veloAvg <= 90.5 && pctPct > 40) {
      agg[5] = agg[0];
      delete agg[0];
      console.log(`[FanGraphs] BIS FB→SI 再分類 (${Math.round(veloAvg)}mph, ${Math.round(pctPct)}%)`);
    }
  }

  let found = false;
  for (const [idxStr, { pctSum, veloNum, veloDen }] of Object.entries(agg)) {
    const idx    = Number(idxStr);
    const pctVal = Math.round(pctSum * 100);  // FG は 0〜1 の小数
    if (pctVal < 5) continue;
    const veloVal = veloDen > 0 ? Math.round(veloNum / veloDen) : null;
    const key     = PITCH_KEYS[idx];
    // Savant 実測値（ba/slg/velo）があれば保持し、FanGraphs は pct（と velo 未取得時）のみ補完
    const existing = rawPitch[yr]?.[key] || {};
    rawPitch[yr][key] = {
      velo: (existing.velo && existing.velo !== '--') ? existing.velo : (veloVal ? String(veloVal) : '--'),
      ba:   (existing.ba   && existing.ba   !== '--') ? existing.ba   : '--',
      slg:  (existing.slg  && existing.slg  !== '--') ? existing.slg  : '--',
      pct:  String(pctVal),
    };
    if (veloVal) {
      const kyui = calcKyuiPreStatcast(veloVal, idx, pctVal);
      if (kyui !== '') {
        if (!showKyuiMap[yr]) showKyuiMap[yr] = {};
        showKyuiMap[yr][idx] = kyui;
      }
    }
    found = true;
  }
  return found;
}

// FanGraphs から複数年分の球種データを一括取得
async function fetchFanGraphsPitchData(fgId, targetYears, rawPitch, showKyuiMap, knuckleFsYears) {
  // 年別リクエストを並列取得（最大4並列でFanGraphsのレート制限を回避）
  const CONCURRENCY = 4;
  let count = 0;
  for (let i = 0; i < targetYears.length; i += CONCURRENCY) {
    const batch = targetYears.slice(i, i + CONCURRENCY);
    const results = await Promise.allSettled(batch.map(yr => {
      const url = `https://www.fangraphs.com/api/leaders/major-league/data?pos=P&stats=pit&lg=all&qual=0&season=${yr}&season1=${yr}&type=4&players=${fgId}&pageitems=1&pagenum=1`;
      return fgGet(url).then(data => ({ yr, rows: data.data || [] }));
    }));
    for (const r of results) {
      if (r.status === 'fulfilled' && r.value.rows.length) {
        if (applyFgRow(r.value.rows[0], r.value.yr, rawPitch, showKyuiMap, knuckleFsYears)) count++;
      }
    }
  }
  return count;
}

// MLB The Show API から選手カードを検索
// ※ items API は name フィルタが効かないためバイナリサーチ（最大 log2(146)≈7 回）で探索
async function fetchMLBTheShowCard(playerName) {
  if (!playerName) return null;
  const nameLower = playerName.trim().toLowerCase();
  const lastName  = nameLower.split(/\s+/).pop();

  const showGet = async (yr, page) =>
    mlbGet(`https://mlb${yr}.theshow.com/apis/items.json?type=mlb_card&page=${page}`);

  for (const yr of [25, 24]) {
    try {
      const first = await showGet(yr, 1);
      const total = first.total_pages || 0;
      if (!total) continue;

      // バイナリサーチ: カードはアルファベット順 (full name)
      let lo = 1, hi = total;
      let found = null;
      while (lo <= hi && !found) {
        const mid = Math.floor((lo + hi) / 2);
        const pg  = mid === 1 ? first : await showGet(yr, mid);
        const items = pg.items || [];
        if (!items.length) break;

        // このページに対象選手がいるか確認
        found = items.find(item =>
          !item.is_hitter &&
          (item.name || '').toLowerCase().includes(lastName) &&
          Array.isArray(item.pitches) && item.pitches.length > 0
        );
        if (found) break;

        // ソート位置による進行方向判定
        const midName = (items[0].name || '').toLowerCase();
        if (midName < lastName) lo = mid + 1;
        else                    hi = mid - 1;
      }
      if (found) {
        console.log(`[The Show MLB${yr}] カード発見: ${found.name} (${found.rarity})`);
        return found;
      }
    } catch (_) {}
  }
  return null;
}

// ── キャリア速度カーブ (aging curve) ─────────────────────────────────────────
// debutYear: MLBデビュー年, careerLength: キャリア稼働年数
// 返値:
//   toPeak(yr, refVelo)  → refVelo が yr 年時点の観測値として、ピーク球速を逆算
//   fromPeak(yr, peakVelo) → ピーク球速からyr年の推定球速を算出
function buildVeloCurve(debutYear, careerLength) {
  const PEAK_CAREER_YEAR = 0;   // デビュー時がピーク（上昇フェーズなし）
  const DECLINE_RATE     = 0.35; // mph/yr（下降期）
  let declineStart;
  if (careerLength >= 20) {
    declineStart = Math.round(careerLength * 0.8);
  } else if (careerLength >= 16) {
    declineStart = 16;
  } else {
    declineStart = 12;
  }
  const careerYearOf = yr => Number(yr) - debutYear; // 0-indexed
  // refVelo: yr年での観測球速 → ピーク球速を逆算
  const toPeak = (yr, refVelo) => {
    const cy = careerYearOf(yr);
    if (cy <= declineStart) return refVelo;
    return refVelo + (cy - declineStart) * DECLINE_RATE;
  };
  // peakVelo: ピーク球速 → yr年の推定球速
  const fromPeak = (yr, peakVelo) => {
    const cy = careerYearOf(yr);
    let v;
    if (cy <= declineStart) v = peakVelo;
    else v = peakVelo - (cy - declineStart) * DECLINE_RATE;
    return Math.max(v, peakVelo - 6); // フロア: ピーク -6mph
  };
  return {
    toPeak, fromPeak,
    info: `debut=${debutYear} careerLen=${careerLength} peakAt=+${PEAK_CAREER_YEAR}yr declineFrom=+${declineStart}yr`,
  };
}

// ── Baseball Savant browser scraping + MLB The Show API (pre-2017) ──────────
// 設計:
//  Step1: Baseball Savant JSON API を in-page fetch で直接叩く（2017年以降・HTMLパース不要）
//         → pct列は「同一年の合計≈100」ヒューリスティックで特定
//  Step2: 2016年以前の未取得年 → MLB The Show API で球種・球速・球威を取得
//         （Brooksbaseball.net は廃止のため削除）
async function fetchBrowserData(slug, id, years, onProgress, playerName = '', apiKey = '', englishName = '', basic = {}) {
  const chromePath = findChrome();
  if (!chromePath) throw new Error('Chromeが見つかりません。Google ChromeまたはEdgeをインストールしてください。');

  const tmpDir = path.join(os.tmpdir(), 'mlb_pitcher_' + Date.now());
  onProgress('ブラウザを起動中...');
  const browser = await puppeteer.launch({
    executablePath: chromePath,
    headless: false,
    userDataDir: tmpDir,
    args: ['--disable-blink-features=AutomationControlled', '--no-first-run', '--no-default-browser-check'],
    ignoreDefaultArgs: ['--enable-automation'],
    defaultViewport: null,
  });

  try {
    const page = await browser.newPage();
    await page.evaluateOnNewDocument(() => {
      Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
      window.chrome = { runtime: {} };
    });

    // 画像・フォント・メディアをブロックして読み込みを高速化
    // ※ stylesheet はブロックしない: Baseball Savant はCSS依存でデータテーブルをレンダリングするため
    //   ブロックするとXHR取得後のテーブル描画が失敗し被打率/SLGが欠落する
    await page.setRequestInterception(true);
    const BLOCK_TYPES = new Set(['image', 'media', 'font']);
    page.on('request', req => {
      if (BLOCK_TYPES.has(req.resourceType())) req.abort();
      else req.continue();
    });

    const rawPitch = {};
    for (const yr of years) rawPitch[yr] = emptyPitchP();

    // ── 球速ブースト計算ヘルパー ─────────────────────────────────────────────
    // FanGraphs BIS / aging curve 推定値に対して補正を加える。
    // 速球系(ff/si): ベース+3mph ± 成績補正3mph
    // 変化球系: ベース+1mph ± 成績補正1.5mph
    // 成績補正: ERA・被打率・被本塁打率を各-1〜+1スコアに正規化して平均 → ±range をかける
    const FASTBALL_KEYS_SET = new Set(['ff', 'si']);
    const calcVeloBoostForYear = (key, basicYr) => {
      const isFB  = FASTBALL_KEYS_SET.has(key);
      const base  = isFB ? 3 : 1;
      const range = isFB ? 3 : 1.5;
      if (!basicYr) return base;
      const era = (typeof basicYr.era === 'number') ? basicYr.era : parseFloat(basicYr.era || '');
      const baa = basicYr.avg ? parseFloat(basicYr.avg) : NaN;
      const ip  = parseFloat(basicYr.ip || '0');
      const hr  = (typeof basicYr.hr === 'number') ? basicYr.hr : parseFloat(basicYr.hr || '');
      const hr9 = (ip > 0 && !isNaN(hr)) ? hr * 9 / ip : NaN;
      const scores = [];
      if (!isNaN(era))  scores.push(Math.min(1, Math.max(-1, (4.0 - era)   / 1.5)));
      if (!isNaN(baa))  scores.push(Math.min(1, Math.max(-1, (0.260 - baa) / 0.040)));
      if (!isNaN(hr9))  scores.push(Math.min(1, Math.max(-1, (1.0 - hr9)   / 0.5)));
      const perfAdj = scores.length > 0
        ? (scores.reduce((s, v) => s + v, 0) / scores.length) * range
        : 0;
      return base + perfAdj;
    };

    // ── サブタイプ追跡（表示名決定用）──────────────────────────────────────
    // { [key]: { [engName]: 累積pct } } 各データソースから球種名と割合を記録し
    // 最も比率の高いサブタイプを BG1以降の表示名として使用する
    const subtypeTracker = {};
    const trackSubtype = (key, engName, pctVal) => {
      const pct = parseFloat(String(pctVal).replace('%', ''));
      if (!key || !engName || isNaN(pct) || pct <= 0) return;
      if (!subtypeTracker[key]) subtypeTracker[key] = {};
      subtypeTracker[key][engName] = (subtypeTracker[key][engName] || 0) + pct;
    };

    // ── ヘルパー ─────────────────────────────────────────────────────────────
    // yearHasPct: 合計投球割合が85%以上あれば「Savant pct 取得済み」とみなす
    // （少数球種しか取れていない年は FanGraphs で補完させる）
    const yearHasPct = (yr) => {
      const total = PITCH_KEYS.reduce((s, k) => {
        const p = parseFloat(String(rawPitch[yr]?.[k]?.pct ?? ''));
        return s + (isNaN(p) ? 0 : p);
      }, 0);
      return total >= 85;
    };

    // rawPitch[yr] に Baseball Savant パース結果をマージ。
    // ── 設計方針 ────────────────────────────────────────────────────────────────
    // Sweeper + Slider → 'sl'、Knuckle Curve + Curveball → 'cu' のように
    // 同一キーにマッピングされる複数サブタイプを正しく合算する。
    // ・pct : 合算（15% + 22% = 37%）
    // ・velo / ba / slg : pct 加重平均
    // ※ 旧実装は「先着優先」だったため Sweeper/Knuckle Curve の pct が無視されて
    //   他球種が正規化で水増しされていた（ダルビッシュ FF 過多バグ）。
    const mergeHtmlData = (htmlData, yr) => {
      // 1. キー別にサブタイプをグループ化
      const byKey = {};
      for (const [ptName, vals] of Object.entries(htmlData)) {
        const key = PITCH_MAP_P[ptName];
        if (!key) continue;
        trackSubtype(key, ptName, vals.pct);
        if (!byKey[key]) byKey[key] = [];
        byKey[key].push(vals);
      }

      // pct フォーマット検出（安全策）: evaluate() 側で変換されるべきだが
      // 万一 0-1 の小数形式で来た場合は ×100 してパーセントに揃える
      // ※ 有効な pct の最大値が 1.0 未満かつ 2件以上 → 小数形式と判断
      const _allRawPcts = Object.values(htmlData)
        .map(v => parseFloat(String(v.pct || '')))
        .filter(n => !isNaN(n) && n > 0);
      const _pctMax   = _allRawPcts.length > 0 ? Math.max(..._allRawPcts) : 0;
      const pctScale  = (_allRawPcts.length >= 2 && _pctMax < 1.0) ? 100 : 1;

      // ── 重複排除 ────────────────────────────────────────────────────────────
      // Baseball Savant は複数テーブルで同じデータを "Split Finger"/"Split-Finger"、
      // "Four Seamer"/"4-Seam Fastball" のように表記違いで保持する。
      // 同一キーに同じ pct 値（小数第1位まで一致）のエントリが複数ある場合は重複とみなし、
      // データが最も充実しているエントリ1つに統合する（pct 二重加算を防ぐ）。
      for (const [key, list] of Object.entries(byKey)) {
        if (list.length <= 1) continue;
        const deduped = [];
        const seenPct = new Map(); // pctKey → deduped 配列のインデックス
        for (const v of list) {
          const p = parseFloat(String(v.pct || '')) * pctScale;
          if (!isNaN(p) && p > 1.0) {
            const pKey = p.toFixed(1);
            if (seenPct.has(pKey)) {
              // 既存エントリに不足データをマージ（velo/ba/slg/pa を補完）
              const ex = deduped[seenPct.get(pKey)];
              if (ex.velo === '--' && v.velo !== '--') ex.velo = v.velo;
              if (ex.ba   === '--' && v.ba   !== '--') ex.ba   = v.ba;
              if (ex.slg  === '--' && v.slg  !== '--') ex.slg  = v.slg;
              if (ex.pa   === '--' && v.pa   !== '--') ex.pa   = v.pa;
            } else {
              seenPct.set(pKey, deduped.length);
              deduped.push({ ...v });
            }
          } else {
            deduped.push(v); // pct なし / 1%以下はそのまま保持
          }
        }
        byKey[key] = deduped;
      }

      // 2. キー別に pct 合算 / velo は pct 加重平均 / ba・slg は PA（打席）加重平均
      for (const [key, list] of Object.entries(byKey)) {
        let pctTotal = 0;
        let veloN = 0, veloD = 0;
        let baN   = 0, baD   = 0;
        let slgN  = 0, slgD  = 0;

        for (const v of list) {
          const pRaw = parseFloat(String(v.pct || ''));
          const p    = !isNaN(pRaw) && pRaw > 0 ? pRaw * pctScale : NaN;  // 小数→% 変換適用
          const pa   = parseFloat(String(v.pa  || ''));
          // pct が有効なエントリのみ pctTotal に加算
          if (!isNaN(p) && p > 0) pctTotal += p;

          // 球速: pct 加重平均（Sweeper の様に pct='--' でも velo がある場合は別途 fallback）
          const vl = parseFloat(String(v.velo || ''));
          if (!isNaN(p) && p > 0 && !isNaN(vl) && vl > 0) { veloN += p * vl; veloD += p; }

          // BA/SLG: PA 加重平均（PA が取れない場合は pct を代用）
          const w = (!isNaN(pa) && pa > 0) ? pa : (!isNaN(p) && p > 0 ? p : 0);
          const ba  = parseFloat(String(v.ba  || ''));
          const slg = parseFloat(String(v.slg || ''));
          if (w > 0 && !isNaN(ba))  { baN  += w * ba;  baD  += w; }
          if (w > 0 && !isNaN(slg)) { slgN += w * slg; slgD += w; }
        }

        // velo fallback: pct='--' のサブタイプ（Sweeper 等）の velo を単純平均で補完
        // pct 加重平均で velo が取れなかった場合のみ使用
        if (veloD <= 0) {
          let fbVN = 0, fbVC = 0;
          for (const v of list) {
            const vl = parseFloat(String(v.velo || ''));
            if (!isNaN(vl) && vl > 0) { fbVN += vl; fbVC++; }
          }
          if (fbVC > 0) { veloN = fbVN / fbVC; veloD = 1; }  // 平均値を veloN/1 として設定
        }

        if (pctTotal <= 0) {
          // pct は取得できなかったが velo/ba/slg がある場合は保存
          // （FanGraphs が後で pct を補完する; applyFgRow が ba/slg を保持する）
          const fbVelo = veloD > 0 ? String(+(veloN / veloD).toFixed(1)) : '--';
          const fbBa   = baD  > 0 ? String(+(baN  / baD ).toFixed(3)) : '--';
          const fbSlg  = slgD > 0 ? String(+(slgN / slgD).toFixed(3)) : '--';
          if (fbVelo !== '--' || fbBa !== '--' || fbSlg !== '--') {
            rawPitch[yr][key] = { velo: fbVelo, ba: fbBa, slg: fbSlg, pct: '--' };
          }
          continue;
        }

        rawPitch[yr][key] = {
          velo: veloD > 0 ? String(+(veloN / veloD).toFixed(1)) : '--',
          ba:   baD   > 0 ? String(+(baN   / baD  ).toFixed(3)) : '--',
          slg:  slgD  > 0 ? String(+(slgN  / slgD ).toFixed(3)) : '--',
          pct:  String(Math.round(pctTotal)),
        };
      }
    };

    // ── Step 1: Baseball Savant キャリアページ 1 回ナビゲート ────────────────
    // Baseball Savant のキャリアページには全年分の Pitch Tracking テーブルが含まれる。
    // 「Year」列（/^\d{4}$/ にマッチするセル群）を検出し、年別にデータを抽出する。
    // ?season= パラメータは効かないため年別ナビゲーションは行わない。
    try {
      onProgress('Baseball Savant を読み込み中...');
      const savantUrl = `https://baseballsavant.mlb.com/savant-player/${slug}-${id}?stats=statcast-r-pitching-mlb`;
      // 'load' を使用: networkidle2 は Analytics 等の永続コネクションで60s タイムアウトするため
      // XHR データの完了は waitForFunction で BA/SLG の出現を監視して保証する
      await page.goto(savantUrl, { waitUntil: 'load', timeout: 60000 });

      // ── テーブル描画待機（2フェーズ）──────────────────────────────────────
      // フェーズ①: BA/SLG を含む Run Values テーブルが出るまで待つ（最大45秒）
      //   → XHR で後から読み込まれるデータ（被打率・SLG）の出現を確認するため
      //      /.³/（小数点3桁）が現れた時点で XHR 完了と判断する
      // フェーズ②: タイムアウト時は球種名テーブルのみで続行（velo/pct だけでも取得）
      const PITCH_KW = ['4-Seam','Fastball','Sinker','Slider','Riding','Knuckleball','Knuckler'];
      try {
        await page.waitForFunction(
          (kw) => {
            for (const t of document.querySelectorAll('table')) {
              const txt = t.innerText || '';
              if (/\d{4}/.test(txt) && kw.some(k => txt.includes(k)) && /\.\d{3}/.test(txt))
                return true;
            }
            return false;
          },
          { timeout: 45000 },
          PITCH_KW
        );
      } catch {
        // BA/SLG なしでも球種テーブルがあれば続行（velo/pct のみ取得）
        try {
          await page.waitForFunction(
            (kw) => [...document.querySelectorAll('table')].some(t =>
              /\d{4}/.test(t.innerText || '') && kw.some(k => (t.innerText||'').includes(k))
            ),
            { timeout: 10000 },
            PITCH_KW
          );
        } catch { /* テーブルが見つからなくても続行 */ }
      }

      // キャリアページの多年度 Pitch Tracking テーブルを一括パース
      // ── 設計方針 ────────────────────────────────────────────────────────────
      // 投球率(%) の取得戦略:
      //   ① `%` 列ヘッダーが存在すれば直接読む
      //   ② `#` / `pitches` 列（投球数）があれば 年間合計で除して計算（最も信頼性が高い）
      //   ヒューリスティック（合計≈100）は廃止。K%・xwOBA等が誤検知されるため。
      const careerData = await page.evaluate((yrs) => {
        const KWDS = ['4-Seam','Fastball','Four Seam','Seamer','Riding','Rising',
                      'Sinker','Slider','Changeup','Change-up','Curveball','Cutter','Split',
                      'Sweeper','Knuckle','Two-Seam','2-Seam','Hard Slider','Slow Curve'];
        function hasPK(t) { return t && KWDS.some(k => t.includes(k)); }

        const result = {};  // { "2021": { "4-Seam Fastball": {velo,ba,slg,pct} } }

        for (const tbl of document.querySelectorAll('table')) {
          const tblText = tbl.innerText || '';
          // 球種名キーワードと4桁年を両方含むテーブルのみ対象
          if (!hasPK(tblText) || !/\d{4}/.test(tblText)) continue;

          const allRows = [...tbl.querySelectorAll('tbody tr')]
            .map(r => [...r.querySelectorAll('td')])
            .filter(c => c.length >= 3);
          if (allRows.length < 2) continue;

          const nCols = Math.max(0, ...allRows.map(c => c.length));

          // ── 年列検出: /^\d{4}$/ にマッチするセルが最も多い列 ──
          let yearCol = -1, bestYearCnt = 0;
          for (let col = 0; col < nCols; col++) {
            const cnt = allRows.filter(c =>
              /^\d{4}$/.test((c[col]?.innerText || '').trim())
            ).length;
            if (cnt > bestYearCnt) { bestYearCnt = cnt; yearCol = col; }
          }
          if (yearCol < 0 || bestYearCnt < 1) continue;

          // ── 球種名列検出 ──
          let pitchCol = -1;
          for (let col = 0; col < nCols; col++) {
            if (col === yearCol) continue;
            if (allRows.some(c => hasPK((c[col]?.innerText || '')))) {
              pitchCol = col; break;
            }
          }
          if (pitchCol < 0) continue;

          // ── ヘッダー取得（ソートアイコン等の特殊文字を除去）──
          // Baseball Savant のテーブルヘッダーには ↕ などのソートアイコンが含まれるため
          // 英数字・%・#・スペース 以外の文字を除去してから比較する
          const hdrRow = tbl.querySelector('thead tr') || tbl.querySelector('tr');
          const hdr = hdrRow
            ? [...hdrRow.querySelectorAll('th,td')].map(h => {
                return h.innerText.trim().toLowerCase()
                  .replace(/[^\w%#\s]/g, '')   // 矢印・特殊記号を除去
                  .replace(/\s+/g, ' ')
                  .trim();
              })
            : [];

          // 各列のインデックス（ソートアイコン除去後のヘッダーで一致）
          const veloIdx = hdr.findIndex(h =>
            h === 'mph' || h.includes('velo') || h.includes('velocity') || h.includes('speed'));
          const baIdx   = hdr.findIndex(h =>
            h === 'ba' || h === 'avg' || h === 'batting avg' || h === 'batting average');
          const slgIdx  = hdr.findIndex(h =>
            h === 'slg' || h === 'slg%' || h === 'slugging' || h.startsWith('slg'));
          // PA（打席数）: BA/SLG の加重平均で使用
          const paIdx   = hdr.findIndex(h =>
            h === 'pa' || h === 'abs' || h === 'ab' || h === 'plate appearances' || h === 'at bats');

          // 投球割合: "%" 列を複数パターンで検索
          // （Baseball Savant は "%" / "Pitch%" / "Pitches%" 等いくつかの表記を使う）
          // whiff%・k%・put away% 等の "%" 含む別列と区別するため
          // 厳格な短パターン優先 → より長いパターンに広げる
          const pctIdx = (() => {
            const exact = hdr.findIndex(h => h === '%');
            if (exact >= 0) return exact;
            // 短い "%" 系ヘッダーを探す（whiff%, k%, put away% などは除外）
            return hdr.findIndex(h =>
              (h === 'pitch%' || h === 'pitches%' || h === 'usage%' ||
               h === 'pitch %' || h === 'pitches %' || h === '% pitches' ||
               h === 'pitch pct' || h === 'pct') &&
              !h.includes('whiff') && !h.includes('k%') && !h.includes('put')
            );
          })();

          // 投球数列: "%" が取れない場合に合計から割合を計算するフォールバック用
          const countIdx = hdr.findIndex(h =>
            h === '#' || h === 'pitch count' || h === 'total pitches' ||
            (h === 'pitches' && pctIdx < 0));  // "pitches" は "%" 列があれば重複するため除外

          // ── 行データを年別に抽出（複数テーブルのフィールドをマージ）──
          for (const cells of allRows) {
            const yr = (cells[yearCol]?.innerText || '').trim();
            if (!yrs.includes(yr)) continue;
            const pt = (cells[pitchCol]?.innerText || '').trim();
            if (!pt || !hasPK(pt)) continue;

            const g = (idx) => (idx >= 0 && idx < cells.length)
              ? (cells[idx]?.innerText.trim() || '--') : '--';

            // velo / ba / slg / pa / count はヘッダーマッチした列から（既存値があれば上書きしない）
            if (!result[yr]) result[yr] = {};
            const cur = result[yr][pt] || { velo: '--', ba: '--', slg: '--', pct: '--', pa: '--', count: '--' };
            result[yr][pt] = {
              velo:  (cur.velo  && cur.velo  !== '--') ? cur.velo  : g(veloIdx),
              ba:    (cur.ba    && cur.ba    !== '--') ? cur.ba    : g(baIdx),
              slg:   (cur.slg   && cur.slg   !== '--') ? cur.slg   : g(slgIdx),
              pa:    (cur.pa    && cur.pa    !== '--') ? cur.pa    : (paIdx    >= 0 ? g(paIdx)    : '--'),
              count: (cur.count && cur.count !== '--') ? cur.count : (countIdx >= 0 ? g(countIdx) : '--'),
              // pct: 上記の複数パターンで検索したインデックスから読む
              pct:   (cur.pct   && cur.pct   !== '--') ? cur.pct   : (pctIdx   >= 0 ? g(pctIdx)   : '--'),
            };
          }
          // break しない: 複数テーブルから全フィールドを収集する（Pitch Movement で velo、Run Values で ba/slg/pct）
        }

        // ── 後処理: pct が取れなかった球種は投球数（count）から割合を計算 ──
        // Run Values テーブルに "%" 列が見つからない場合のフォールバック
        for (const pitches of Object.values(result)) {
          const totalCount = Object.values(pitches).reduce((s, v) => {
            const n = parseFloat(v.count || '');
            return s + (isNaN(n) ? 0 : n);
          }, 0);
          if (totalCount <= 0) continue;
          for (const v of Object.values(pitches)) {
            if (v.pct && v.pct !== '--') continue;  // pct 取得済みはスキップ
            const n = parseFloat(v.count || '');
            if (!isNaN(n) && n > 0) v.pct = String(+(n / totalCount * 100).toFixed(1));
          }
        }

        return result;
      }, years);

      // careerData を rawPitch にマージ
      for (const [yr, pitchMap] of Object.entries(careerData)) {
        mergeHtmlData(pitchMap, yr);
      }
      const gotYears = Object.keys(careerData).filter(y => Object.keys(careerData[y]).length > 0);
      onProgress(`Baseball Savant: ${gotYears.length} 年分のピッチデータを取得`);

    } catch (e) {
      onProgress('⚠ Baseball Savant 取得失敗: ' + e.message);
    }

    // ── Step 2: pre-2017 年 → FanGraphs → MLB The Show → Claude の順でフォールバック ──
    const showKyuiMap = {};
    // pct が取得できなかった年は年代を問わず FanGraphs で補完する
    // （2017+ で Savant の % 列が欠落した場合も FanGraphs が pct を提供する）
    const preShowYears = years.filter(yr => !yearHasPct(yr));
    // デビュー年・キャリア長（Step 2c/2d 共通で使用）
    const debutYearNum  = Math.min(...years.map(Number));
    const lastYearNum   = Math.max(...years.map(Number));
    const careerLenNum  = lastYearNum - debutYearNum + 1;
    // Claude キャリアプロファイル（2c で取得 → 2d で活用）
    let claudeProfile = null;
    // The Show が充填した年（2d で velo を aging curve 補正するために追跡）
    const showFilledYears = new Set();
    // The Show カード参照（2d でピーク球速の下限推定に使用）
    let showCardRef = null;

    if (preShowYears.length > 0) {
      // ── 2a: FanGraphs (API key不要、2002年以降) ────────────────────────────
      const fgTargetYears = preShowYears.filter(yr => +yr >= 2002);
      // knuckleFsYears: FanGraphs で KN%（ナックルボール列）が実際に取得された年を記録。
      // ★ 球速でのナックル判定は廃止。SF%（スプリット）が偶然低速でも誤判定しない。
      const knuckleFsYears = new Set();
      if (fgTargetYears.length > 0) {
        onProgress(`FanGraphs 選手検索中... (${fgTargetYears.length}年分を取得予定)`);
        try {
          const fgId = await fetchFanGraphsId(englishName || playerName);
          if (fgId) {
            onProgress(`FanGraphs ID: ${fgId} → 球種データ取得中...`);
            const fgCount = await fetchFanGraphsPitchData(fgId, fgTargetYears, rawPitch, showKyuiMap, knuckleFsYears);
            if (fgCount > 0) {
              onProgress(`FanGraphs: ${fgCount}/${fgTargetYears.length}年分 取得完了`);
            } else {
              onProgress('FanGraphs: 球種データなし');
            }
          } else {
            onProgress('FanGraphs: 選手が見つかりませんでした');
          }
        } catch (e) {
          onProgress('⚠ FanGraphs 取得失敗: ' + e.message);
        }

        // ── FanGraphs 実測値にも球速ブースト適用（速球+3/変化球+1 ± 成績補正）──
        // FanGraphs BIS の実測値は実際の球速より低めに記録されているため、
        // aging curve と同じ補正を直接FanGraphsデータにも適用する。
        // ※ Statcast 導入（2015+）以降の年は Savant の実測値を使うためブーストしない
        // （Randy Johnson 2002 で急激にスピードが落ちないようにする）
        for (const yr of fgTargetYears.filter(y => yearHasPct(y) && +y < 2015)) {
          for (const key of PITCH_KEYS) {
            const d = rawPitch[yr]?.[key];
            if (!d || d.velo === '--') continue;
            const origVelo = parseFloat(d.velo);
            if (isNaN(origVelo) || origVelo <= 0) continue;
            const boost = calcVeloBoostForYear(key, basic[yr]);
            const newVelo = Math.round(origVelo + boost);
            rawPitch[yr][key].velo = String(newVelo);
            const ki = PITCH_KEYS.indexOf(key);
            const pctNum = parseFloat(String(d.pct).replace('%', ''));
            const byr = basic[yr];
            const bEraFg  = byr ? parseFloat(String(byr.era)) : NaN;
            const bBaaFg  = byr ? (byr.avg ? Number(byr.avg) * 1000 : NaN) : NaN;
            const bIpFg   = byr ? parseFloat(String(byr.ip).replace(/\.(\d)$/, '.$10')) : 0;
            const bHr9Fg  = (byr && bIpFg > 0) ? (byr.hr * 9 / bIpFg) : NaN;
            const kyui = calcKyuiPreStatcast(newVelo, ki, isNaN(pctNum) ? 20 : pctNum, bEraFg, bBaaFg, bHr9Fg);
            if (kyui !== '') {
              if (!showKyuiMap[yr]) showKyuiMap[yr] = {};
              showKyuiMap[yr][ki] = kyui;
            }
          }
          onProgress(`[FanGraphs boost] ${yr}: 速球+3/変化球+1 ± 成績補正適用`);
        }

        // ── 2a.2: Wikipedia 球種プロファイル取得（無料・APIキー不要）──────────────
        // Wikipedia 公開 API から投手の球種・球速情報を取得して FanGraphs データを検証する。
        // Anthropic API 不要のため常に実行する。結果は 2a.3 の Claude 補正がない場合の
        // フォールバックとして、または 2a.3 の補強として活用する。
        let wikiProfile = null;
        {
          // 日本語名は除去して英語名で検索（Wikipedia は英語版を優先）
          const wikiSearchName = (englishName || playerName).replace(/[　-鿿゠-ヿ぀-ゟ]/g, '').trim();
          if (wikiSearchName) {
            onProgress(`Wikipedia 球種プロファイル検索中: "${wikiSearchName}"...`);
            try {
              wikiProfile = await fetchWikipediaPitchProfile(wikiSearchName);
              if (wikiProfile) {
                const typeStr  = wikiProfile.pitchKeys.length
                  ? wikiProfile.pitchKeys.map(k => {
                      const idx = PITCH_KEYS.indexOf(k);
                      return idx >= 0 ? PITCH_NAMES_JA[idx] : k;
                    }).join(' / ')
                  : '（球種情報なし）';
                const veloStr = wikiProfile.veloMentions.length
                  ? ` 球速言及: ${Math.min(...wikiProfile.veloMentions)}-${Math.max(...wikiProfile.veloMentions)}mph`
                  : '';
                onProgress(`[2a.2 Wikipedia] "${wikiProfile.pageTitle}": ${typeStr}${veloStr}`);

                // ── FG データとの球種整合性チェック ──
                // FanGraphs が実際にデータを取得した年の球種セットを確認
                const fgActiveKeys = PITCH_KEYS.filter(k =>
                  fgTargetYears.some(yr => {
                    const d = rawPitch[yr]?.[k];
                    return d && d.pct !== '--' && parseFloat(String(d.pct)) > 5;
                  })
                );
                // Wikipedia にあって FG にない球種（FG の取りこぼし可能性）
                for (const key of wikiProfile.pitchKeys) {
                  if (!fgActiveKeys.includes(key)) {
                    const jn = PITCH_NAMES_JA[PITCH_KEYS.indexOf(key)] || key;
                    onProgress(`[2a.2 ⚠] ${jn}(${key}) が Wikipedia で言及されているが FanGraphs データに見当たりません`);
                  }
                }
                // FG にあって Wikipedia で言及のない球種（FF↔SI の混同確認）
                for (const key of fgActiveKeys) {
                  if (wikiProfile.pitchKeys.length > 0 && !wikiProfile.pitchKeys.includes(key)) {
                    // ff と si は混同しやすいため相互は警告しない（FG の FB→SI 再分類で同族扱い）
                    const sibling = key === 'ff' ? 'si' : key === 'si' ? 'ff' : null;
                    if (sibling && wikiProfile.pitchKeys.includes(sibling)) continue;
                    const jn = PITCH_NAMES_JA[PITCH_KEYS.indexOf(key)] || key;
                    onProgress(`[2a.2 ℹ] FanGraphsの ${jn}(${key}) は Wikipedia で言及なし（分類誤り・少用球種の可能性）`);
                  }
                }

                // ── Wikipedia 最高球速 → 主速球の速度キャップに活用 ──
                // Wikipedia で言及される最大球速は「キャリア最高」に相当することが多い。
                // FanGraphs+ブースト後に Wikipedia peak を超えている年は過剰ブーストの可能性あり。
                // Claude 2a.3 が careerPeakSpeeds を返した場合はそちらが優先される。
                if (wikiProfile.veloMentions.length > 0) {
                  const wikiPeak = Math.max(...wikiProfile.veloMentions);
                  const primaryFfKey = wikiProfile.primaryKey === 'si' ? 'si' : 'ff';
                  // フォールバック用として wikiProfile に格納（2a.3 非実行時に使用）
                  wikiProfile._capKey  = primaryFfKey;
                  wikiProfile._capPeak = wikiPeak; // 瞬間最大球速 (mph)
                  onProgress(`[2a.2 Wikipedia] 最高球速参考値: ${primaryFfKey}=${wikiPeak}mph（2a.3 未実行時のキャップに使用）`);
                }
              } else {
                onProgress('[2a.2 Wikipedia] 球種情報が見つかりませんでした（記事なし、または球種の記述なし）');
              }
            } catch (e) {
              onProgress('⚠ Wikipedia 取得エラー: ' + e.message);
            }
          }
        }

        // ── 2a.3: 複数ソース補正 (THT/MLB.com/BR/Wikipedia/BA) ─────────────────
        // FanGraphs BIS の球種分類・球速誤りを一次資料で検証して補正する。
        // ①Baseball Savant(Statcast) 実測年(2015+)はスキップ。
        // ②ODT プロファイルで overrideFgPitch=true の年は ODT が既に正しいためスキップ。
        // ③確実な資料根拠がある場合のみ上書き（推測補正はしない）。
        // ④ 2a.3 非実行時は wikiProfile._capPeak を主速球の速度キャップとして代用。
        if (apiKey) {
          // 補正対象: FanGraphs が実際に球種データを埋めた年 (yearHasPct=true) かつ pre-2015
          const docxP = DOCX_PLAYER_PROFILES[englishName] || DOCX_PLAYER_PROFILES[playerName] || null;
          const fgFilledYears = fgTargetYears.filter(yr =>
            yearHasPct(yr) && +yr < 2015 &&
            // overrideFgPitch=true の年は ODT がすでに上書きするためスキップ
            !(docxP?.overrideFgPitch && docxP.yearPcts?.[yr])
          );
          if (fgFilledYears.length > 0) {
            onProgress(`複数ソース補正: THT/MLB.com/BR/Wikipedia/BA で ${fgFilledYears.length}年分を検証中...`);
            try {
              // 現在の FanGraphs データ概要を構築（Claude への参考情報）
              const fgSummary = {};
              for (const yr of fgFilledYears) {
                fgSummary[yr] = {};
                for (const key of PITCH_KEYS) {
                  const d = rawPitch[yr]?.[key];
                  if (!d || d.pct === '--') continue;
                  const pct = parseFloat(String(d.pct).replace('%', ''));
                  if (isNaN(pct) || pct <= 0) continue;
                  const velo = parseFloat(d.velo);
                  fgSummary[yr][key] = {
                    speedMph: isNaN(velo) ? null : Math.round(velo),
                    pct: Math.round(pct),
                  };
                }
              }
              const fgCorr = await callClaudeForFgCorrection(
                apiKey, englishName || playerName, fgFilledYears, fgSummary
              );
              if (fgCorr) {
                if (fgCorr.pitcherCharacteristics) {
                  onProgress(`[2a.3] 投手特徴: ${fgCorr.pitcherCharacteristics}`);
                }
                const noteStr = fgCorr.note ? `（${fgCorr.note}）` : '';

                // ── ① yearCorrections: 特定年の球種・球速・割合を上書き ──
                if (Array.isArray(fgCorr.yearCorrections)) {
                  for (const corr of fgCorr.yearCorrections) {
                    const corrYears = (corr.years || []).filter(y => fgFilledYears.includes(y));
                    if (!corrYears.length || !Array.isArray(corr.pitches)) continue;
                    for (const yr of corrYears) {
                      // 補正対象球種キーセットを取得（補正後にpctを正規化するため）
                      const corrKeys = new Set((corr.pitches || []).map(p => p.key).filter(k => PITCH_KEYS.includes(k)));
                      const untouchedKeys = PITCH_KEYS.filter(k => !corrKeys.has(k));
                      // 補正対象外の球種が大きな割合を持っている場合は pct を再調整
                      for (const p of corr.pitches) {
                        const key = p.key;
                        if (!PITCH_KEYS.includes(key)) continue;
                        const corrSpeedMph = typeof p.avgSpeedMph === 'number' && p.avgSpeedMph > 0
                          ? p.avgSpeedMph : null;
                        const corrPct     = typeof p.pct === 'number' && p.pct > 0 ? p.pct : null;
                        if (!corrSpeedMph && !corrPct) continue;
                        if (!rawPitch[yr]) rawPitch[yr] = {};
                        // 既存エントリの ba/slg は保持（Savant 実測値を守るため）
                        const existing = rawPitch[yr][key] || { velo: '--', ba: '--', slg: '--', pct: '--' };
                        const newVelo = corrSpeedMph ? String(corrSpeedMph) : existing.velo;
                        const newPct  = corrPct      ? String(corrPct)      : existing.pct;
                        rawPitch[yr][key] = { velo: newVelo, ba: existing.ba, slg: existing.slg, pct: newPct };
                        // showKyuiMap を再計算
                        const ki = PITCH_KEYS.indexOf(key);
                        const byr = basic[yr];
                        const bEraC  = byr ? parseFloat(String(byr.era)) : NaN;
                        const bBaaC  = byr ? (byr.avg ? Number(byr.avg) * 1000 : NaN) : NaN;
                        const bIpC   = byr ? parseFloat(String(byr.ip).replace(/\.(\d)$/, '.$10')) : 0;
                        const bHr9C  = (byr && bIpC > 0) ? (byr.hr * 9 / bIpC) : NaN;
                        const pctNum = corrPct ?? parseFloat(String(existing.pct).replace('%', ''));
                        const kyui = calcKyuiPreStatcast(corrSpeedMph || parseFloat(existing.velo), ki, isNaN(pctNum) ? 20 : pctNum, bEraC, bBaaC, bHr9C);
                        if (kyui !== '') {
                          if (!showKyuiMap[yr]) showKyuiMap[yr] = {};
                          showKyuiMap[yr][ki] = kyui;
                        }
                        const noteKind = p.note ? ` (${p.note})` : '';
                        onProgress(`[2a.3 補正] ${yr} ${key}: speed=${newVelo}mph pct=${newPct}%${noteKind}`);
                      }
                      // 補正で追加された球種がなければ不要な球種を削除しない（既存を保持）
                    }
                  }
                }

                // ── ② careerPeakSpeeds: 全年への最高球速キャップ ──
                // 瞬間最大球速 → シーズン平均キャップ(-3.7mph) で rawPitch を上限制御
                if (fgCorr.careerPeakSpeeds && typeof fgCorr.careerPeakSpeeds === 'object') {
                  let capCount = 0;
                  for (const [key, peakMph] of Object.entries(fgCorr.careerPeakSpeeds)) {
                    if (!PITCH_KEYS.includes(key)) continue;
                    if (typeof peakMph !== 'number' || peakMph <= 0) continue;
                    // 瞬間最大 → シーズン平均上限（callClaudeForPeakProfile と同じ -3.7mph 変換）
                    const seasonAvgCap = peakMph - 3.7;
                    const ki = PITCH_KEYS.indexOf(key);
                    for (const yr of fgFilledYears) {
                      const d = rawPitch[yr]?.[key];
                      if (!d || d.velo === '--') continue;
                      const v = parseFloat(d.velo);
                      if (isNaN(v) || v <= seasonAvgCap) continue;
                      rawPitch[yr][key].velo = String(+seasonAvgCap.toFixed(1));
                      // showKyuiMap 再計算
                      const byr = basic[yr];
                      const bEraC  = byr ? parseFloat(String(byr.era)) : NaN;
                      const bBaaC  = byr ? (byr.avg ? Number(byr.avg) * 1000 : NaN) : NaN;
                      const bIpC   = byr ? parseFloat(String(byr.ip).replace(/\.(\d)$/, '.$10')) : 0;
                      const bHr9C  = (byr && bIpC > 0) ? (byr.hr * 9 / bIpC) : NaN;
                      const pctNum = parseFloat(String(d.pct).replace('%', ''));
                      const kyui = calcKyuiPreStatcast(seasonAvgCap, ki, isNaN(pctNum) ? 20 : pctNum, bEraC, bBaaC, bHr9C);
                      if (kyui !== '') {
                        if (!showKyuiMap[yr]) showKyuiMap[yr] = {};
                        showKyuiMap[yr][ki] = kyui;
                      }
                      capCount++;
                    }
                  }
                  if (capCount > 0) {
                    const peakSummary = Object.entries(fgCorr.careerPeakSpeeds)
                      .map(([k, v]) => `${k}:${v}mph`).join(' / ');
                    onProgress(`[2a.3 球速キャップ] 最高球速(${peakSummary}) からシーズン平均上限を適用 (${capCount}件)${noteStr}`);
                  }
                }

                if (!Array.isArray(fgCorr.yearCorrections) || fgCorr.yearCorrections.length === 0) {
                  onProgress(`[2a.3] 補正不要 — FanGraphsデータは資料と一致${noteStr}`);
                }
              } else {
                onProgress('[2a.3] 複数ソース: データ取得できませんでした');
              }
            } catch (e) {
              onProgress('⚠ 複数ソース補正 取得失敗: ' + e.message);
            }
          }
        }

        // ── 2a.3 非実行時のフォールバック: Wikipedia 最高球速キャップを適用 ──────────
        // Claude API なし（apiKey 未設定）かつ wikiProfile に球速情報がある場合に限り適用。
        // 2a.3 が実行された場合は careerPeakSpeeds が同様のキャップを担うためスキップ。
        if (!apiKey && wikiProfile?._capPeak && wikiProfile?._capKey) {
          const capKey  = wikiProfile._capKey;
          const capPeak = wikiProfile._capPeak;        // Wikipedia 最高球速 (mph)
          const seasonAvgCap = capPeak - 3.7;          // 瞬間最大 → シーズン平均上限
          const ki = PITCH_KEYS.indexOf(capKey);
          let wikiCapCount = 0;
          const fgFilledYearsForWiki = fgTargetYears.filter(yr => yearHasPct(yr) && +yr < 2015);
          for (const yr of fgFilledYearsForWiki) {
            const d = rawPitch[yr]?.[capKey];
            if (!d || d.velo === '--') continue;
            const v = parseFloat(d.velo);
            if (isNaN(v) || v <= seasonAvgCap) continue;
            rawPitch[yr][capKey].velo = String(+seasonAvgCap.toFixed(1));
            const byr = basic[yr];
            const bEra = byr ? parseFloat(String(byr.era)) : NaN;
            const bBaa = byr ? (byr.avg ? Number(byr.avg) * 1000 : NaN) : NaN;
            const bIp  = byr ? parseFloat(String(byr.ip).replace(/\.(\d)$/, '.$10')) : 0;
            const bHr9 = (byr && bIp > 0) ? (byr.hr * 9 / bIp) : NaN;
            const pctNum = parseFloat(String(d.pct).replace('%', ''));
            const kyui = calcKyuiPreStatcast(seasonAvgCap, ki, isNaN(pctNum) ? 20 : pctNum, bEra, bBaa, bHr9);
            if (kyui !== '') {
              if (!showKyuiMap[yr]) showKyuiMap[yr] = {};
              showKyuiMap[yr][ki] = kyui;
            }
            wikiCapCount++;
          }
          if (wikiCapCount > 0) {
            onProgress(`[2a.2 Wikipedia キャップ] ${capKey} 最高球速 ${capPeak}mph → シーズン平均上限 ${seasonAvgCap.toFixed(1)}mph 適用 (${wikiCapCount}年)`);
          }
        }

        // ── サブタイプ追跡: FanGraphs KN% で 'fs' が埋まった年を Knuckleball として登録 ──
        // FanGraphs KN% 列が実際に取得された年のみ Knuckleball としてサブタイプ登録する。
        // ★ 旧ロジック（球速 ≤83mph → Knuckleball 判定）は廃止:
        //    スプリット/フォークが高齢化等で 83mph 以下になることがあり誤判定を招く。
        //    真のナックルボーラー（ディッキー 72-77mph / ウェイクフィールド 65-70mph）は
        //    そもそも KN% 列にデータが入るため、列名での判定が唯一正確。
        for (const yr of fgTargetYears.filter(y => knuckleFsYears.has(y))) {
          const fsD = rawPitch[yr]?.['fs'];
          if (!fsD || fsD.pct === '--') continue;
          trackSubtype('fs', 'Knuckleball', parseFloat(fsD.pct) || 10);
        }
      }

      // ── 2a.5: ODTプロファイル(③) — FanGraphs未取得年を補完 ────────────────
      // FanGraphs(④)で pitch mix が取れなかった年(例: Shields 2001-2006)に
      // ODT分析ドキュメントの年度別球種・球速・割合を設定する。
      // yearHasPct=true(Savant/FanGraphs実測済み)の年はスキップ。
      // ただし overrideFgPitch=true の場合は FanGraphs実測年も ODT yearPcts で上書きする。
      // (例: 高津臣吾 — FanGraphs BIS の FB→SI 誤再分類を回避するため)
      {
        const docxP = DOCX_PLAYER_PROFILES[englishName] || DOCX_PLAYER_PROFILES[playerName] || null;
        if (docxP?.yearPcts) {
          for (const yr of years.filter(y => y !== '通算' && (!yearHasPct(y) || docxP.overrideFgPitch === true))) {
            const yPct = docxP.yearPcts[yr];
            // overrideFgPitch=true でも yearPcts に定義のない年は FanGraphs実測を保持
            if (!yPct) continue;
            const phase = docxP.phases?.find(p => +yr >= p.from && +yr <= p.to);
            let filledCount = 0;
            PITCH_KEYS.forEach(key => {
              const pct = yPct[key] ?? 0;
              if (pct <= 0) {
                rawPitch[yr][key] = { velo: '--', ba: '--', slg: '--', pct: '--' };
                return;
              }
              const avgKmh = phase?.[key] ?? null;
              if (!avgKmh) return;
              const mph = +((avgKmh - 4) / 1.6).toFixed(1);
              rawPitch[yr][key] = { velo: String(mph), ba: '--', slg: '--', pct: String(pct) };
              // showKyuiMap にも事前計算
              const ki = PITCH_KEYS.indexOf(key);
              const byr = basic[yr];
              const bEra = byr ? parseFloat(String(byr.era)) : NaN;
              const bBaa = byr ? (byr.avg ? Number(byr.avg) * 1000 : NaN) : NaN;
              const bIp  = byr ? parseFloat(String(byr.ip).replace(/\.(\d)$/, '.$10')) : 0;
              const bHr9 = (byr && bIp > 0) ? (byr.hr * 9 / bIp) : NaN;
              const kyui = calcKyuiPreStatcast(mph, ki, pct, bEra, bBaa, bHr9);
              if (kyui !== '') {
                if (!showKyuiMap[yr]) showKyuiMap[yr] = {};
                showKyuiMap[yr][ki] = kyui;
              }
              filledCount++;
            });
            if (filledCount > 0) onProgress(`[③ ODT] ${yr}: ${filledCount}球種を設定 (FanGraphs未取得年)`);
          }
        }
      }

      // ── 2b: MLB The Show (FanGraphs で未取得の年 + pre-2002) ──────────────
      const afterFgMissing = preShowYears.filter(yr => !yearHasPct(yr));
      if (afterFgMissing.length > 0) {
        onProgress(`MLB The Show API 検索中 (未取得 ${afterFgMissing.length}年分)...`);
        try {
          const showCard = await fetchMLBTheShowCard(englishName || playerName);
          if (showCard) {
            showCardRef = showCard; // 2d でピーク球速推定に使用
            onProgress(`MLB The Show: "${showCard.name}" (${showCard.rarity}) カード発見`);
            const pcts = estimateShowUsagePct(showCard.pitches.length);
            for (const yr of afterFgMissing) {
              showCard.pitches.forEach((p, i) => {
                const idx = PITCH_MAP_SHOW[p.name];
                if (idx === undefined) return;
                const key = PITCH_KEYS[idx];
                trackSubtype(key, p.name, pcts[i] || 5); // サブタイプ追跡
                const kyui = calcKyuiFromShow(p.speed, p.control, p.movement);
                // The Show の球速属性値（1-99スケール）は mph ではないため velo は設定しない
                // → aging curve が FanGraphs 実測値から逆算した mph を書き込む
                rawPitch[yr][key] = { velo: '--', ba: '--', slg: '--', pct: String(pcts[i] || 5) };
                if (kyui !== '') {
                  if (!showKyuiMap[yr]) showKyuiMap[yr] = {};
                  showKyuiMap[yr][idx] = kyui;
                }
              });
              showFilledYears.add(yr); // The Show で充填した年を記録
            }
            onProgress(`MLB The Show: ${afterFgMissing.length}年 × ${showCard.pitches.length}球種 設定完了`);
          } else {
            onProgress('MLB The Show: 該当カード未収録');
          }
        } catch (e) {
          onProgress('⚠ MLB The Show API 取得失敗: ' + e.message);
        }
      }
    }

    // ── 2b.5: ODTプロファイル(③) — 全年への球速キャップ + 禁止球種クリア ────
    // FanGraphsブーストで上限を超えた球速を ODT記載の最大値に抑える。
    // Savant実測(①)を含む全年に適用（Savant実測が上限を超えることはないが念のため）。
    // pitchMaxKmh[key]=0 の球種は rawPitch・showKyuiMap から完全削除。
    {
      const docxP = DOCX_PLAYER_PROFILES[englishName] || DOCX_PLAYER_PROFILES[playerName] || null;
      if (docxP?.pitchMaxKmh) {
        const maxKmh = docxP.pitchMaxKmh;
        let capCount = 0;
        for (const yr of years.filter(y => y !== '通算')) {
          for (const key of PITCH_KEYS) {
            const d = rawPitch[yr]?.[key];
            if (!d) continue;
            const limit = maxKmh[key];
            if (limit === undefined) continue;
            const ki = PITCH_KEYS.indexOf(key);
            // limit=0: この球種は投球しない → 全フィールドを '--' に
            if (limit === 0) {
              if (d.velo !== '--' || d.pct !== '--') {
                rawPitch[yr][key] = { velo: '--', ba: '--', slg: '--', pct: '--' };
                if (showKyuiMap[yr]) delete showKyuiMap[yr][ki];
                onProgress(`[③ ODT] ${yr} ${key}: 投球なし → 削除`);
                capCount++;
              }
              continue;
            }
            // limit>0: 球速が上限を超えていればキャップ + showKyuiMap 再計算
            if (d.velo !== '--') {
              const veloNum = parseFloat(d.velo);
              const maxMph = (limit - 4) / 1.6;
              if (!isNaN(veloNum) && veloNum > maxMph) {
                const cappedMph = +maxMph.toFixed(1);
                rawPitch[yr][key].velo = String(cappedMph);
                const byr = basic[yr];
                const bEra = byr ? parseFloat(String(byr.era)) : NaN;
                const bBaa = byr ? (byr.avg ? Number(byr.avg) * 1000 : NaN) : NaN;
                const bIp  = byr ? parseFloat(String(byr.ip).replace(/\.(\d)$/, '.$10')) : 0;
                const bHr9 = (byr && bIp > 0) ? (byr.hr * 9 / bIp) : NaN;
                const pctNum = parseFloat(String(d.pct).replace('%', ''));
                const newKyui = calcKyuiPreStatcast(cappedMph, ki, isNaN(pctNum) ? 20 : pctNum, bEra, bBaa, bHr9);
                if (newKyui !== '') {
                  if (!showKyuiMap[yr]) showKyuiMap[yr] = {};
                  showKyuiMap[yr][ki] = newKyui;
                }
                onProgress(`[③ ODT] ${yr} ${key}: ${Math.round(veloNum * 1.6 + 4)}km/h → キャップ ${limit}km/h`);
                capCount++;
              }
            }
          }
        }
        if (capCount === 0) onProgress('[③ ODT] 球速キャップ: 適用なし（全球種が上限以内）');
      }
    }

    // ── 2c: Claude ウェブ検索でキャリアピーク球速プロファイルを取得 ─────────
    // APIキーがある場合は常に実行（FanGraphs/The Show有無に関係なく）。
    // 年別データは埋めず、ピーク球速・球種プロファイルのみ取得 → Step 2d で aging curve に使用。
    if (apiKey) {
      onProgress(`Claude ウェブ検索: ${englishName || playerName} のキャリアピーク球速を調査中...`);
      try {
        const profile = await callClaudeForPeakProfile(apiKey, englishName || playerName, debutYearNum);
        if (profile && Array.isArray(profile.pitches) && profile.pitches.length > 0) {
          claudeProfile = profile;
          const noteStr = profile.note ? `（${profile.note}）` : '';
          onProgress(`Claude プロファイル取得: ${profile.pitches.length}球種${noteStr}`);
          profile.pitches.forEach(p => onProgress(`  ${p.name}: peakSpeed=${p.peakSpeed}mph, avgPct=${p.avgPct}%`));
        } else {
          onProgress('Claude: プロファイルを取得できませんでした（aging curve でフォールバック）');
        }
      } catch (e) {
        onProgress('⚠ Claude プロファイル取得失敗: ' + e.message);
      }
    }

    // ── 2d: キャリア軌跡推定 (全データソースで未取得の年を aging curve で推定) ──
    // claudeProfile がある場合: Claude のピーク球速を基準に aging curve を適用
    // ない場合: 取得済み参照年の観測値から逆算してピークを推定
    // ※ The Show はカード一枚の一律速度を全年に適用するため、claudeProfile がある場合は
    //   showFilledYears の velocity も aging curve で上書きする（pct/球種は保持）
    const stillMissing = preShowYears.filter(yr => !yearHasPct(yr));
    // The Show 充填年は claudeProfile の有無に関係なく常に aging curve で velo 補正する
    // （The Show はカード1枚の一律速度を全年に貼り付けるため）
    const showVeloOverrideYears = Array.from(showFilledYears);
    if (stillMissing.length > 0 || showVeloOverrideYears.length > 0) {
      const curve = buildVeloCurve(debutYearNum, careerLenNum);
      onProgress(`[aging curve] ${curve.info}`);

      const refYears = years.filter(yr => yearHasPct(yr));
      const profileByKey = {};

      for (const key of PITCH_KEYS) {
        const ki = PITCH_KEYS.indexOf(key);

        // ── 優先①: Claude プロファイルにピーク球速がある場合 ──
        // Claude は最高球速（瞬間最速）を返すため、-3.7mph（≈-6km/h）してシーズン平均相当に変換
        // claudeBased:true → FanGraphsバイアス補正のベースブーストは不要（成績補正のみ適用）
        if (claudeProfile) {
          const cp = claudeProfile.pitches.find(p => PITCH_MAP_SHOW[p.name] === ki);
          if (cp && cp.peakSpeed > 0) {
            trackSubtype(key, cp.name, cp.avgPct || 20); // サブタイプ追跡
            const peakVeloAdj = Math.round((cp.peakSpeed - 3.7) * 10) / 10; // 最高球速 → シーズン平均相当
            profileByKey[key] = { peakVelo: peakVeloAdj, avgPct: cp.avgPct || 20, claudeBased: true };
            onProgress(`[aging curve] ${key}: Claude maxSpeed=${cp.peakSpeed}mph → peakVelo=${peakVeloAdj}mph (-3.7mph) avgPct=${cp.avgPct}%`);
            continue;
          }
        }

        // ── 優先②: 参照年の観測値からピーク球速を逆算 ──
        // ※ refYears の velo は FanGraphs boost 適用済み（+3/+1 ± 成績補正）のため、
        //   toPeak する前にブーストを除去して生の FanGraphs 値に戻す。
        //   profileByKey には生の（プレブースト）ピーク球速を格納し、
        //   各推定年に apply する際に改めて calcVeloBoostForYear を足す。
        if (refYears.length === 0) continue;
        const valid = refYears
          .map(yr => ({ yr, d: rawPitch[yr]?.[key] }))
          .filter(({ d }) => d && d.velo !== '--' && parseFloat(d.velo) > 0 &&
                             d.pct !== '--' && parseFloat(String(d.pct).replace('%','')) >= 5);
        if (!valid.length) continue;
        const peakEstimates = valid.map(({ yr, d }) => {
          const boostedVelo = parseFloat(d.velo);
          const boost = calcVeloBoostForYear(key, basic[yr]); // FanGraphs boost を除去
          return curve.toPeak(yr, boostedVelo - boost);       // 生の FanGraphs 値でピーク逆算
        });
        const peakVelo = peakEstimates.reduce((s, v) => s + v, 0) / peakEstimates.length;
        const avgPct   = valid.reduce((s, { d }) =>
          s + parseFloat(String(d.pct).replace('%','')), 0) / valid.length;
        profileByKey[key] = { peakVelo, avgPct, claudeBased: false };
        onProgress(`[aging curve] ${key}: 逆算 peakVelo=${peakVelo.toFixed(1)}mph (raw) avgPct=${avgPct.toFixed(1)}%`);
      }

      // ── 優先②.5: The Show 属性値をプレブーストのピーク下限として使用 ─────────
      // FanGraphs のデータが選手の晩年しかない場合（例: Randy Johnson 2002-2004 = 38-40歳）、
      // 逆算ピーク球速は過小評価になる。The Show 属性値(1-99)は FanGraphs と同スケールの
      // ロー値として扱い、ブーストは推定時に加算する。
      // 例: Show属性93 → raw peak 93mph → 推定時 93+6=99mph → 99*1.6+4=162km/h
      if (!claudeProfile && showCardRef && showFilledYears.size > 0) {
        for (const key of PITCH_KEYS) {
          const ki = PITCH_KEYS.indexOf(key);
          const sp = showCardRef.pitches.find(p => PITCH_MAP_SHOW[p.name] === ki);
          if (!sp || sp.speed <= 0) continue;
          const showPeakEstimate = sp.speed; // Show属性(1-99) = FanGraphs同スケールのロー値
          const currentPeak = profileByKey[key]?.peakVelo ?? 0;
          if (showPeakEstimate > currentPeak) {
            profileByKey[key] = { peakVelo: showPeakEstimate, avgPct: profileByKey[key]?.avgPct ?? 20, claudeBased: false };
            onProgress(`[aging curve] ${key}: Show属性${sp.speed}=rawPeak採用 (FanGraphs逆算raw${currentPeak.toFixed(1)}mphより高いため上書き)`);
          }
        }
      }

      if (Object.keys(profileByKey).length > 0) {
        // ─ stillMissing 年: 全フィールドを新規設定 ─
        for (const yr of stillMissing) {
          for (const [key, { peakVelo, avgPct, claudeBased }] of Object.entries(profileByKey)) {
            // claudeBased: Claude が真の球速を返す → FanGraphs バイアス補正のベースブーストは不要
            // !claudeBased: FanGraphs/Show 由来のロー値 → ブースト (base+perf) を加算
            const boost = claudeBased ? 0 : calcVeloBoostForYear(key, basic[yr]);
            const estVelo = Math.round(curve.fromPeak(yr, peakVelo) + boost);
            rawPitch[yr][key] = { velo: String(estVelo), ba: '--', slg: '--', pct: String(Math.round(avgPct)) };
            // ★ 推定球速でのナックルボール判定は廃止（スプリット/フォークが 83mph 以下になりうるため誤判定の原因）。
            // 推定年のナックルボール表示は FanGraphs KN% 実績・Savant・The Show・Claude 検索で担保する。
          }
          // 割合を100に正規化してから球威を計算（成績データも渡す）
          const rawPcts = PITCH_KEYS.map(k => rawPitch[yr][k]?.pct ?? '--');
          const normalized = normalizePctToSum100(rawPcts);
          const byr = basic[yr];
          const bEra  = byr ? parseFloat(String(byr.era))  : NaN;
          const bBaa  = byr ? (byr.avg ? Number(byr.avg) * 1000 : NaN) : NaN;
          const bIp   = byr ? parseFloat(String(byr.ip).replace(/\.(\d)$/, '.$10'))  : 0;
          const bHr9  = (byr && bIp > 0) ? (byr.hr * 9 / bIp) : NaN;
          PITCH_KEYS.forEach((key, ki) => {
            if (!profileByKey[key]) return;
            const normPct = normalized[ki];
            if (normPct === '--' || Number(normPct) < 5) {
              rawPitch[yr][key] = { velo: '--', ba: '--', slg: '--', pct: '--' };
              return;
            }
            rawPitch[yr][key].pct = normPct;
            const veloNum = parseFloat(rawPitch[yr][key].velo);
            const kyui = calcKyuiPreStatcast(veloNum, ki, Number(normPct), bEra, bBaa, bHr9);
            if (kyui !== '') {
              if (!showKyuiMap[yr]) showKyuiMap[yr] = {};
              showKyuiMap[yr][ki] = kyui;
            }
          });
        }
        if (stillMissing.length > 0) {
          onProgress(`キャリア軌跡推定完了: ${stillMissing.join(', ')} (${Object.keys(profileByKey).length}球種, 速球+3/変化球+1 ± 成績補正済み)`);
        }

        // ─ showVeloOverrideYears: velocity のみ aging curve で上書き（pct/球種は The Show 値を保持）─
        if (showVeloOverrideYears.length > 0) {
          for (const yr of showVeloOverrideYears) {
            const syrB   = basic[yr];
            const sEra   = syrB ? parseFloat(String(syrB.era))  : NaN;
            const sBaa   = syrB ? (syrB.avg ? Number(syrB.avg) * 1000 : NaN) : NaN;
            const sIp    = syrB ? parseFloat(String(syrB.ip).replace(/\.(\d)$/, '.$10')) : 0;
            const sHr9   = (syrB && sIp > 0) ? (syrB.hr * 9 / sIp) : NaN;
            for (const [key, { peakVelo, claudeBased }] of Object.entries(profileByKey)) {
              if (!rawPitch[yr]?.[key]) continue; // velo='--' でも aging curve で上書きする
              const boost = claudeBased ? 0 : calcVeloBoostForYear(key, basic[yr]);
              const estVelo = Math.round(curve.fromPeak(yr, peakVelo) + boost);
              rawPitch[yr][key].velo = String(estVelo);
              // 球威も aging curve 球速で再計算（成績データ付き）
              const ki = PITCH_KEYS.indexOf(key);
              const pctNum = parseFloat(String(rawPitch[yr][key].pct).replace('%', ''));
              const kyui = calcKyuiPreStatcast(estVelo, ki, isNaN(pctNum) ? 20 : pctNum, sEra, sBaa, sHr9);
              if (kyui !== '') {
                if (!showKyuiMap[yr]) showKyuiMap[yr] = {};
                showKyuiMap[yr][ki] = kyui;
              }
            }
          }
          onProgress(`The Show 年 velo 補正完了: ${showVeloOverrideYears.join(', ')} → aging curve 速球+3/変化球+1 ± 成績補正適用`);
        }
      }
    }

    // ── 2e: 緩急差・変化量ボーナスを showKyuiMap に適用 (Statcast非実測年のみ) ──
    // Baseball Savant 実測年 (BA/SLG に数値がある年) は Savant が実際の
    // 変化量・変化方向を計測しているためスキップ。
    // FanGraphs(④)・ODT(③)・Claude推定・Aging curve 由来の年に以下を加算する:
    //   ① PITCH_MOVEMENT_BONUS: 球種固有の変化量特性による推定補正
    //   ② calcKakkyoSaBonus:    主速球との球速差が大きいほど変化球が有効になる緩急差補正
    // 適用後の値は addAbilityToFile で perfBoost と合算されて最終球威になる。
    for (const yr of years.filter(y => y !== '通算')) {
      const yrPitch = rawPitch[yr];
      if (!yrPitch || !showKyuiMap[yr]) continue;
      // BA/SLG 実測値 (Statcast) がある年はスキップ
      const hasStatcast = PITCH_KEYS.some(k => {
        const d = yrPitch[k];
        return d && d.ba !== '--' && d.ba !== '' && d.ba != null && !isNaN(Number(d.ba));
      });
      if (hasStatcast) continue;
      const ffMph = getPrimaryFfMph(yrPitch);
      let bonusLog = [];
      for (const key of PITCH_KEYS) {
        const ki = PITCH_KEYS.indexOf(key);
        const cur = showKyuiMap[yr][ki];
        if (cur === undefined) continue;
        const d = yrPitch[key];
        if (!d || d.velo === '--') continue;
        const speed = parseFloat(d.velo);
        if (isNaN(speed) || speed <= 0) continue;
        const movBonus     = PITCH_MOVEMENT_BONUS[ki] ?? 0;
        const kakkyoBonus  = calcKakkyoSaBonus(speed, ki, ffMph);
        const totalBonus   = movBonus + kakkyoBonus;
        if (totalBonus > 0) {
          const newVal = Math.max(30, Math.min(110, Number(cur) + totalBonus));
          showKyuiMap[yr][ki] = newVal;
          bonusLog.push(`${key}:+${totalBonus}(mov+${movBonus}/緩急差+${kakkyoBonus})`);
        }
      }
      if (bonusLog.length > 0) {
        onProgress(`[2e 変化量・緩急差] ${yr}: ${bonusLog.join(' ')}`);
      }
    }

    // ── pitchNameOverrides を算出（サブタイプ追跡結果から）──────────────────
    // SUBTYPE_DISPLAY_JA に登録されているサブタイプのうち、
    // 累積 pct が最大のものを各 key の表示名として採用する
    const pitchNameOverrides = {};
    const KEY_TO_IDX = Object.fromEntries(PITCH_KEYS.map((k, i) => [k, i]));
    for (const [key, nameMap] of Object.entries(subtypeTracker)) {
      const dominant = Object.entries(nameMap).sort((a, b) => b[1] - a[1])[0]?.[0];
      const idx = KEY_TO_IDX[key];
      if (dominant !== undefined && SUBTYPE_DISPLAY_JA[dominant] !== undefined && idx !== undefined) {
        pitchNameOverrides[idx] = SUBTYPE_DISPLAY_JA[dominant];
      }
    }

    return { rawPitch, showKyuiMap, pitchNameOverrides };
  } finally {
    await browser.close();
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch {}
  }
}

// ── Excel build ───────────────────────────────────────────────────────────────
async function buildExcel(playerName, years, basic, vsLeftByYear, rawPitch) {
  const N_MAIN = 22;
  const N_SUB  = 4;

  // Career pitch data: アウト数加重平均
  const careerPitch = emptyPitchP();
  for (const key of PITCH_KEYS) {
    const entries = years
      .map(yr => ({ outs: ipToOuts(basic[yr]?.ip || '0'), d: rawPitch[yr]?.[key] }))
      .filter(e => e.outs > 0 && e.d?.velo && e.d.velo !== '--');
    if (!entries.length) continue;

    const wAvg = (field, toDecimal) => {
      const valid = entries.filter(e => e.d[field] && e.d[field] !== '--');
      if (!valid.length) return '--';
      const totOut = valid.reduce((s, e) => s + e.outs, 0);
      if (!totOut) return '--';
      const sum = valid.reduce((s, e) => {
        let v = parseFloat(String(e.d[field]).replace('%', ''));
        if (isNaN(v)) return s;
        if (toDecimal && v > 1) v = v / 100;
        return s + v * e.outs;
      }, 0);
      const avg = sum / totOut;
      return toDecimal ? String(avg.toFixed(3)) : String(avg.toFixed(1));
    };
    careerPitch[key] = {
      velo: wAvg('velo', false),
      ba:   wAvg('ba',   true),
      slg:  wAvg('slg',  true),
      pct:  wAvg('pct',  false),  // % 形式で保存（decimal変換しない）
    };
  }

  // 通算平均で穴埋め: pct が有効なのに velo / ba / slg が欠けている年を補完
  for (const yr of years) {
    for (const key of PITCH_KEYS) {
      const d = rawPitch[yr]?.[key];
      if (!d) continue;
      const pn = parseFloat(String(d.pct ?? '').replace('%', ''));
      if (isNaN(pn) || pn <= 0) continue;          // 使用率0の球種は補完不要
      const career = careerPitch[key];
      if (!career) continue;
      if ((!d.velo || d.velo === '--') && career.velo && career.velo !== '--')
        rawPitch[yr][key].velo = career.velo;
      if ((!d.ba   || d.ba   === '--') && career.ba   && career.ba   !== '--')
        rawPitch[yr][key].ba   = career.ba;
      if ((!d.slg  || d.slg  === '--') && career.slg  && career.slg  !== '--')
        rawPitch[yr][key].slg  = career.slg;
    }
  }

  // 投球率が 5% 未満・0・'--'・不明 の球種をすべてマスク（velo/ba/slg/pct → '--'）
  // ※ pct は % 形式（0〜100）。"0.9"=0.9%、"45.4"=45.4% — decimal 変換しない。
  // ※ pct が '--' の場合でも velo/ba/slg に値が残ることがある（例: 2023年シンカーの.000等）。
  //    この場合も全フィールドをマスクして表上に不要なデータが残らないようにする。
  function maskLowUsage(src) {
    for (const key of PITCH_KEYS) {
      const d = src[key];
      if (!d) continue;
      const pctStr = String(d.pct ?? '').replace('%', '').trim();
      const pn     = parseFloat(pctStr);
      // マスク条件: pct が '--' / '' / NaN / 0以下 / 5%未満 のいずれか
      if (pctStr === '--' || pctStr === '' || isNaN(pn) || pn < 5.0) {
        src[key] = { velo: '--', ba: '--', slg: '--', pct: '--' };
      }
    }
  }
  for (const yr of years) maskLowUsage(rawPitch[yr]);
  maskLowUsage(careerPitch);

  // Pre-compute normalized usage% for all years + career (total = 100)
  const normalizedPct = {};
  for (const yk of [...years, '通算']) {
    const src = yk === '通算' ? careerPitch : rawPitch[yk];
    const rawPcts = PITCH_KEYS.map(k => src?.[k]?.pct ?? '--');
    normalizedPct[yk] = normalizePctToSum100(rawPcts);
  }

  const mainCols = [
    '選手名','年度','チーム','勝利','敗北','防御率','試合数','GS','HLD','セーブ',
    'イニング','被安打','自責点','被本塁打','四死球','奪三振','被打率','WHIP',
    '対左被打率','SB','PK','CS',
  ];
  const subCols = ['球速','被打率','SLG','投球率'];

  const hRow0 = [...mainCols, ...PITCH_NAMES_JA.flatMap(n => [n, '', '', ''])];
  const hRow1 = [...Array(N_MAIN).fill(''), ...PITCH_NAMES_JA.flatMap(() => subCols)];

  function pitchRowVals(yk) {
    const src = yk === '通算' ? careerPitch : rawPitch[yk];
    return PITCH_KEYS.flatMap((key, ki) => {
      const d = src?.[key];
      return [
        d?.velo && d.velo !== '--' ? d.velo : '--',
        fmtAvg(d?.ba),
        fmtAvg(d?.slg),
        normalizedPct[yk]?.[ki] ?? '--',
      ];
    });
  }

  function buildRow(yk) {
    const b = basic[yk];
    if (!b) return Array(N_MAIN + PITCH_KEYS.length * N_SUB).fill('');
    return [
      playerName, yk, b.team,
      b.w, b.l, b.era,
      b.g, b.gs, b.hld, b.sv,
      b.ip,
      b.h, b.er, b.hr, b.bb, b.so,
      fmtAvg(b.avg),
      b.whip,
      fmtAvg(vsLeftByYear[yk]),
      b.sb, b.pk, b.cs,
      ...pitchRowVals(yk),
    ];
  }

  const allRows = [hRow0, hRow1, ...years.map(buildRow), buildRow('通算')];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(allRows);

  ws['!merges'] = [
    ...Array.from({ length: N_MAIN }, (_, c) => ({ s:{r:0,c}, e:{r:1,c} })),
    ...PITCH_NAMES_JA.map((_, i) => {
      const c = N_MAIN + i * N_SUB;
      return { s:{r:0,c}, e:{r:0,c:c+N_SUB-1} };
    }),
  ];

  ws['!cols'] = [
    {wch:12},{wch:6},{wch:6},
    {wch:5},{wch:5},{wch:7},
    {wch:5},{wch:5},{wch:5},{wch:5},
    {wch:8},
    {wch:5},{wch:5},{wch:5},{wch:6},{wch:6},
    {wch:7},{wch:7},{wch:9},
    {wch:5},{wch:5},{wch:5},
    ...Array(28).fill({wch:7}),
  ];

  XLSX.utils.book_append_sheet(wb, ws, playerName + '成績');

  const outFile = path.join(OUT_DIR, playerName + '_成績.xlsx');
  XLSX.writeFile(wb, outFile);

  const ejWb = new ExcelJS.Workbook();
  await ejWb.xlsx.readFile(outFile);
  ejWb.worksheets[0].views = [{ state:'frozen', xSplit:2, ySplit:2, topLeftCell:'C3', activeCell:'C3' }];
  await ejWb.xlsx.writeFile(outFile);

  return outFile;
}

// ── Job management ────────────────────────────────────────────────────────────
const jobs = new Map();

async function runCreateJob(jobId, params) {
  const upd = msg => { const j = jobs.get(jobId); if (j) { j.progress = msg; console.log('[job]', msg); } };
  try {
    upd('MLB Stats API から投手成績を取得中...');
    const { years, basic, vsLeftByYear } = await fetchPitchingStats(params.id, params.y1, params.y2);

    upd('ブラウザを起動して Baseball Savant / MLB The Show から球種データを取得中...');
    const apiKey = params.apiKey || process.env.ANTHROPIC_API_KEY || '';
    const { rawPitch, showKyuiMap, pitchNameOverrides } = await fetchBrowserData(params.slug, params.id, years, upd, params.name, apiKey, params.fullName || '', basic);

    upd('Excel ファイルを生成中...');
    const outFile = await buildExcel(params.name, years, basic, vsLeftByYear, rawPitch);

    upd('スタミナ・制球を計算中...');
    let abilityRows = 0;
    try {
      abilityRows = await addAbilityToFile(outFile, showKyuiMap, pitchNameOverrides);
      upd(`スタミナ・制球追加完了: ${abilityRows} 行`);
    } catch (e) {
      upd('⚠ スタミナ・制球追加失敗: ' + e.message);
    }

    const j = jobs.get(jobId);
    if (j) {
      j.status      = 'done';
      j.result      = path.basename(outFile);
      j.abilityRows = abilityRows;
      j.progress    = '完了';
    }
  } catch (e) {
    const j = jobs.get(jobId);
    if (j) { j.status = 'error'; j.error = e.message; j.progress = 'エラー'; }
    console.error('[job error]', e.message);
  }
}

// ── HTML ──────────────────────────────────────────────────────────────────────
const HTML = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>MLB投手成績ツール</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Meiryo UI','Meiryo','Yu Gothic UI',sans-serif;background:#f0e8f5;
  min-height:100vh;display:flex;align-items:flex-start;justify-content:center;padding:30px 20px}
.card{background:white;border-radius:12px;box-shadow:0 4px 20px rgba(112,48,160,.15);
  padding:32px;width:100%;max-width:640px}
h1{color:#7030A0;font-size:20px;margin-bottom:16px}
h1::before{content:"⚾ "}
.tabs{display:flex;gap:0;margin-bottom:24px;border-bottom:2px solid #e0c8f0}
.tab{padding:10px 22px;cursor:pointer;font-size:14px;font-weight:bold;color:#999;
  border-bottom:3px solid transparent;margin-bottom:-2px;transition:all .15s}
.tab.active{color:#7030A0;border-bottom-color:#7030A0}
.tab:hover:not(.active){color:#555}
.panel{display:none}.panel.active{display:block}
.sec{margin-bottom:16px}
label{display:block;font-size:12px;font-weight:bold;color:#555;margin-bottom:5px}
input{width:100%;padding:9px 12px;border:1px solid #ddd;border-radius:6px;
  font-size:14px;font-family:inherit}
input:focus{outline:none;border-color:#7030A0}
.row{display:flex;gap:10px}.row>div{flex:1}
button{padding:10px 22px;border:none;border-radius:6px;cursor:pointer;
  font-size:14px;font-family:inherit;font-weight:bold;transition:all .15s}
.btn-s{background:#555;color:white;white-space:nowrap}
.btn-s:hover{background:#333}
.btn-p{background:#7030A0;color:white}
.btn-p:hover:not(:disabled){background:#5a1e85}
.btn-p:disabled{background:#bbb;cursor:not-allowed}
.ri{padding:8px 12px;background:#f8f8f8;border:1px solid #eee;border-radius:4px;
  cursor:pointer;margin-bottom:4px;font-size:13px;transition:background .1s}
.ri:hover{background:#f0e8f5;border-color:#ce93d8}
.ri .n{font-weight:bold;color:#333}.ri .m{color:#888;font-size:11px;margin-left:8px}
.results{margin-top:8px}
.ods-area{border:2px dashed #ce93d8;border-radius:8px;padding:12px 16px;
  display:flex;align-items:center;gap:10px;background:#faf5ff;margin-bottom:10px;transition:border-color .2s}
.ods-area.sel{border-color:#7030A0;background:#f3e5f5}
.ods-icon{font-size:22px;flex-shrink:0}
.ods-path{font-size:12px;color:#999;flex:1;word-break:break-all}
.ods-path.has{color:#4a1470;font-weight:bold;font-size:13px}
.pbox{margin-top:16px;padding:14px 16px;background:#f9f5ff;border:1px solid #ce93d8;
  border-radius:8px;display:none}
.ptxt{font-size:13px;color:#555}
.sp{display:inline-block;width:12px;height:12px;border:2px solid #ddd;
  border-top-color:#7030A0;border-radius:50%;animation:spin .7s linear infinite;
  vertical-align:middle;margin-right:6px}
@keyframes spin{to{transform:rotate(360deg)}}
.done{margin-top:14px;padding:14px 16px;background:#e8f5e9;border:1px solid #a5d6a7;
  border-radius:8px;display:none;font-size:14px;color:#2e7d32}
.err{margin-top:14px;padding:14px 16px;background:#ffebee;border:1px solid #ef9a9a;
  border-radius:8px;display:none;font-size:14px;color:#c62828;white-space:pre-wrap}
.note{font-size:11px;color:#aaa;margin-top:14px;line-height:1.7}
.badge-row{display:flex;flex-wrap:wrap;gap:4px;margin-bottom:16px}
.badge{background:#f3e5f5;color:#7030A0;border:1px solid #ce93d8;border-radius:4px;
  padding:3px 8px;font-size:11px;font-weight:bold}
.badge.red{background:#fce4ec;color:#c00060;border-color:#f48fb1}
.opt-label{font-size:11px;color:#888;margin-left:6px;font-weight:normal}
</style>
</head>
<body>
<div class="card">
  <h1>MLB投手成績ツール</h1>
  <div class="tabs">
    <div class="tab active" onclick="switchTab('create',this)">新規作成</div>
    <div class="tab" onclick="switchTab('add',this)">既存ファイルに追加</div>
  </div>

  <!-- ── Tab 1: 新規作成 ── -->
  <div id="panel-create" class="panel active">
    <div class="sec">
      <label>① 選手検索（英語名）</label>
      <div class="row">
        <div style="flex:3"><input id="q" type="text" placeholder="例: Yoshinobu Yamamoto"
          onkeydown="if(event.key==='Enter')doSearch()"></div>
        <div style="flex:1"><button class="btn-s" onclick="doSearch()">🔍 検索</button></div>
      </div>
      <div class="results" id="results"></div>
    </div>
    <div class="sec">
      <label>② 選手情報</label>
      <div class="row">
        <div><label>英語スラッグ</label><input id="slug" type="text" placeholder="yoshinobu-yamamoto"></div>
        <div><label>MLB ID</label><input id="pid" type="number" placeholder="808982"></div>
      </div>
      <div class="row" style="margin-top:10px">
        <div><label>日本語名（ファイル名）</label><input id="jaName" type="text" placeholder="山本由伸"></div>
        <div><label>英語フルネーム</label><input id="fullName" type="text" placeholder="Yoshinobu Yamamoto"></div>
      </div>
      <div class="row" style="margin-top:10px">
        <div><label>開始年</label><input id="y1" type="number" placeholder="2024"></div>
        <div><label>終了年</label><input id="y2" type="number" placeholder="2026"></div>
      </div>
    </div>

    <div class="sec" style="margin-top:4px">
      <label>Claude APIキー <span class="opt-label">（省略可）pre-2017年のMLB The Show未収録選手に使用</span></label>
      <input id="apiKey" type="password" placeholder="sk-ant-..." oninput="saveApiKey(this.value)">
    </div>

    <div class="sec" style="margin-top:4px">
      <div class="badge-row">
        <span class="badge">投手成績取得</span>
        <span style="font-size:14px;color:#aaa;align-self:center">→</span>
        <span class="badge">Excel生成</span>
        <span style="font-size:14px;color:#aaa;align-self:center">→</span>
        <span class="badge red">7球種 × 4項目</span>
        <span style="font-size:14px;color:#aaa;align-self:center">→</span>
        <span class="badge">スタミナ〜対盗塁（AY〜BE）自動追加</span>
      </div>
      <button class="btn-p" id="btnCreate" onclick="doCreate()">▶ 成績ファイルを作成</button>
    </div>
    <div class="pbox" id="cPbox"><div class="ptxt" id="cPtxt"><span class="sp"></span>処理中...</div></div>
    <div class="done" id="cDone"></div>
    <div class="err"  id="cErr"></div>
    <div class="note">
      ※ Chromeが自動起動します（Baseball Savant へのアクセス）<br>
      ※ pre-2017年は <strong>FanGraphs</strong>（2002+・API key不要）→ MLB The Show → Claude の順で球種を補完<br>
      ※ AY=スタミナ・AZ=制球・BA=精神・BB=奪三振・BC=重さ・BD=対左・BE=対盗塁 を自動追加<br>
      ※ 出力先: このツールと同じフォルダ
    </div>
  </div>

  <!-- ── Tab 2: 既存ファイルに追加（将来拡張用） ── -->
  <div id="panel-add" class="panel">
    <div class="sec">
      <p style="font-size:13px;color:#888;margin-top:8px">
        既存の投手成績ファイルへの球種データ再取得は今後対応予定です。<br>
        制球値の追加は <strong>stats_tool2</strong> ツールを使用してください。
      </p>
    </div>
  </div>
</div>

<script>
let cTimer = null;

// APIキー localStorage 読み書き
(function(){
  const k = localStorage.getItem('mlb_tool_apikey');
  if (k) { const el = document.getElementById('apiKey'); if (el) el.value = k; }
})();
function saveApiKey(v) {
  if (v && v.trim()) localStorage.setItem('mlb_tool_apikey', v.trim());
  else localStorage.removeItem('mlb_tool_apikey');
}

function switchTab(id, el) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  document.getElementById('panel-' + id).classList.add('active');
}

async function doSearch() {
  const q = document.getElementById('q').value.trim();
  if (!q) return;
  const el = document.getElementById('results');
  el.innerHTML = '<div style="font-size:12px;color:#888">検索中...</div>';
  try {
    const r = await fetch('/api/search?name=' + encodeURIComponent(q));
    const data = await r.json();
    if (!data.length) { el.innerHTML = '<div style="font-size:12px;color:#888">見つかりませんでした</div>'; return; }
    el.innerHTML = data.map(p =>
      '<div class="ri" onclick="pick(' + p.id + ',\\'' + p.name.replace(/'/g,"\\\\'") + '\\')">' +
      '<span class="n">' + p.name + '</span>' +
      '<span class="m">' + p.position + ' · debut ' + p.debut + ' · ID: ' + p.id + '</span></div>'
    ).join('');
  } catch(e) { el.innerHTML = '<div style="font-size:12px;color:#c00">エラー: '+e.message+'</div>'; }
}

function pick(id, name) {
  document.getElementById('pid').value      = id;
  document.getElementById('fullName').value = name;
  document.getElementById('slug').value     = name.toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'');
  document.getElementById('results').innerHTML =
    '<div style="font-size:12px;color:#7030A0">✓ ' + name + '（ID: ' + id + '）</div>';
}

async function doCreate() {
  const slug=document.getElementById('slug').value.trim(), id=parseInt(document.getElementById('pid').value);
  const name=document.getElementById('jaName').value.trim(), fullName=document.getElementById('fullName').value.trim();
  const y1=parseInt(document.getElementById('y1').value), y2=parseInt(document.getElementById('y2').value);
  const apiKey=(document.getElementById('apiKey').value||'').trim();
  if (!slug||!id||!name||!fullName||!y1||!y2){alert('すべての項目を入力してください');return;}
  document.getElementById('btnCreate').disabled=true;
  document.getElementById('cPbox').style.display='block';
  document.getElementById('cDone').style.display='none';
  document.getElementById('cErr').style.display='none';
  setCP('処理を開始しています...');
  const r=await fetch('/api/create',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({slug,id,name,fullName,y1,y2,apiKey})});
  const {jobId}=await r.json();
  cTimer=setInterval(()=>pollCreate(jobId),1500);
}

async function pollCreate(jobId) {
  const r=await fetch('/api/job/'+jobId); const j=await r.json();
  setCP(j.progress);
  if (j.status==='done') {
    clearInterval(cTimer);
    document.getElementById('cPbox').style.display='none';
    document.getElementById('btnCreate').disabled=false;
    const el=document.getElementById('cDone'); el.style.display='block';
    let msg = '✓ 完了: ' + j.result + ' を作成しました';
    if (j.abilityRows > 0) msg += '（スタミナ・制球: ' + j.abilityRows + ' 行追加）';
    el.textContent = msg;
  } else if (j.status==='error') {
    clearInterval(cTimer);
    document.getElementById('cPbox').style.display='none';
    document.getElementById('btnCreate').disabled=false;
    const el=document.getElementById('cErr'); el.style.display='block';
    el.textContent='✗ エラー: '+j.error;
  }
}

function setCP(msg){document.getElementById('cPtxt').innerHTML='<span class="sp"></span>'+msg;}
</script>
</body>
</html>`;

// ── HTTP Server ───────────────────────────────────────────────────────────────
const server = http.createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');

  if (req.method === 'GET' && url.pathname === '/') {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    return res.end(HTML);
  }
  if (req.method === 'GET' && url.pathname === '/api/search') {
    searchPlayers(url.searchParams.get('name') || '')
      .then(data => { res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' }); res.end(JSON.stringify(data)); })
      .catch(e   => { res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' }); res.end(JSON.stringify({ error: e.message })); });
    return;
  }
  if (req.method === 'GET' && url.pathname === '/api/browse') {
    const fp = browseFile();
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    return res.end(JSON.stringify({ path: fp }));
  }
if (req.method === 'GET' && url.pathname.startsWith('/api/job/')) {
    const job = jobs.get(url.pathname.slice('/api/job/'.length));
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    return res.end(JSON.stringify(job || { status: 'unknown' }));
  }
  if (req.method === 'POST' && url.pathname === '/api/create') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', () => {
      try {
        const params = JSON.parse(body);
        const jobId  = crypto.randomUUID();
        jobs.set(jobId, { status:'running', progress:'開始中...', result:null, abilityRows:0, error:null });
        runCreateJob(jobId, params);
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ jobId }));
      } catch (e) {
        res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ error: e.message }));
      }
    });
    return;
  }
  res.writeHead(404); res.end('Not found');
});

server.on('error', err => {
  if (err.code === 'EADDRINUSE') {
    const url = `http://localhost:${PORT}`;
    console.log('\n  ⚾  MLB投手成績ツール（既に起動済み）\n\n  URL: ' + url + '\n');
    try { spawn('cmd.exe', ['/c', 'start', '', url], { detached:true, shell:false, stdio:'ignore' }).unref(); } catch {}
    setTimeout(() => process.exit(0), 2000);
  } else {
    console.error('サーバーエラー:', err.message);
    process.exit(1);
  }
});

server.listen(PORT, '127.0.0.1', () => {
  const url = `http://localhost:${PORT}`;
  console.log('\n  ⚾  MLB投手成績ツール\n\n  URL: ' + url + '\n  Ctrl+C で停止\n');
  try { spawn('cmd.exe', ['/c', 'start', '', url], { detached:true, shell:false, stdio:'ignore' }).unref(); } catch {}
});
