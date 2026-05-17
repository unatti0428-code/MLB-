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

const PORT    = 3942;
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

async function addAbilityToFile(xlsxPath, showKyuiMap = {}, pitchNameOverrides = {}, extraOptions = {}) {
  // extraOptions:
  //   outPitchBoosts {Object}  idx → 追加球威ポイント  (例: {6: 5} でスプリット+5)
  //   wikiCapKmh     {number|null}  BG列の実km/h上限（Wikipedia 最高球速の実変換値）
  const { outPitchBoosts = {}, wikiCapKmh = null } = extraOptions;
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

  // 通算球威: 各年の球威をアウト数加重平均で算出するためのストア
  // { outs: number, kyuiByI: { [pitchListIdx]: number } }
  const yearKyuiStore = [];

  let count = 0;
  for (const { rn, yr, ipRaw, g, gs, bb, eraRaw, hr, so, avgRaw, vsLRaw, sb, pk, cs, pitchData } of dataRows) {
    const ip = parseIP(ipRaw);
    if (!ip) continue;

    const isCareerRow = String(yr).trim() === '通算';
    const _rowKyuiByI = {}; // この行の球種別 kyui（非通算行のみ蓄積）

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

      // BG列（球速）= 最高球速。
      // 先発投手は長いイニングをペーシングするため FanGraphs/推定のシーズン平均は
      // 実際のゲーム最高球速より 3mph 程度低い。非 Statcast 年（BA='--'）かつ
      // 先発比率 > 50% の場合に +3mph を加算して最高球速に近似させる。
      // Statcast 年（BA あり）は計測精度が高いためそのまま使用。
      // 中継ぎ・抑えはシーズン平均 ≈ 最高球速のためブーストなし。
      const _isStarter       = gs > 0 && g > 0 && (gs / g) > 0.5;
      const _hasStatcastVelo = baStr !== '--' && baStr !== '';
      const _starterBoostMph = (!_hasStatcastVelo && _isStarter) ? 3 : 0;
      // ── BIS-split 対策: フォーシーム(idx=0)の非Statcast先発年 ──────────────────
      // FanGraphs BIS は「速い部分→si」「遅い部分→ff」に分割して記録するケースがある。
      // その場合、si の速度が ff より大幅に高く（>4mph）なる。
      // BG列（最高球速）は速球の最高球速を表すため、si と ff の速い方を採用する。
      // ※ Statcast 年(BA あり)は正確な計測値が入るためそのままとする。
      let _bgVeloNum = veloNum;
      let _bisSplitDetected = false; // BIS-split検知フラグ（球速・球威両方の補正に使用）
      if (pg.idx === 0 && !_hasStatcastVelo && !isNaN(veloNum)) {
        // pitchData[5] = シンカー(si)
        const _siPd  = pitchData[5];
        const _siPct = _siPd ? parseFloat(String(_siPd.pct ?? '').replace('%', '')) : NaN;
        const _siV   = (_siPd && !isNaN(_siPct) && _siPct >= 5)
                       ? parseFloat(String(_siPd.velo || '')) : NaN;
        if (!isNaN(_siV) && _siV > veloNum + 4) {
          _bgVeloNum = _siV; // si が ff より 4mph 以上速い → si 速度を BG・球威表示に使用
          _bisSplitDetected = true;
        }
      }
      let kyusoku = calcKyuSoku(isNaN(_bgVeloNum) ? _bgVeloNum : _bgVeloNum + _starterBoostMph);
      // ── Wikipedia 最高球速 km/h キャップ（BG列表示上限）──────────────────────────
      // Wikipedia が最高球速（例: 100mph=161km/h）を明記している場合、
      // ゲームスケール換算(×1.6+4)の誤差でキャップを超えることがあるため実変換値で上限を設定。
      // ff(idx=0)・si(idx=5) の主速球グループにのみ適用する。
      // ★ 閾値: wikiCapKmh < 150（≒93mph未満）は適用しない。
      //   Wikipediaが低速度のみ言及（変化球・晩年球速等）している場合に
      //   本来の速球BGが誤って下げられる問題を防ぐ。
      if (wikiCapKmh && wikiCapKmh >= 150 && kyusoku !== '' && (pg.idx === 0 || pg.idx === 5) && Number(kyusoku) > wikiCapKmh) {
        kyusoku = wikiCapKmh;
      }
      if (kyusoku !== '') redPurpleCell(ws.getCell(rn, base + 0), kyusoku, fontSize);

      let kyui = '';
      if (isCareerRow) {
        // ── 通算行: 各年度の球威をアウト数加重平均で算出 ──────────────────────────
        // BA/SLGパスで再計算すると「通算pct（全キャリア割合）vs Statcast年のBAのみ」の
        // 不整合が生じ kyui が MAX 超えや転向選手で歪む。
        // 各年度の最終 kyui（補正済み）をそのままアウト数加重平均することで正確な通算値を得る。
        let sumKO = 0, sumO = 0;
        for (const { outs, kyuiByI } of yearKyuiStore) {
          if (kyuiByI[i] !== undefined) { sumKO += kyuiByI[i] * outs; sumO += outs; }
        }
        if (sumO > 0) kyui = Math.round(sumKO / sumO);
      } else {
        // ── 年度行: 通常の計算 ──────────────────────────────────────────────────
        const ah = calcAH_pitch(pg.idx, baNum);
        const ai = calcAI_pitch(pg.idx, slgNum);
        if (ah !== '' && ai !== '') {
          const aj = (Number(ah) + Number(ai)) / 2;
          const ak = calcAK_pitch(pg.idx, aj, pctNum);
          kyui = calcKyuI(aj, ak, pctNum);
        } else {
          // BA/SLG なし年: パフォーマンス補正値を算出
          const perfBoost = calcPerfBoost(perfEra, perfBaa1000, perfHr9);
          // ── BIS-split 対策: フォーシーム(idx=0)で si 速度を採用した年 ──────────────
          // FanGraphs BIS は「速い部分→si」「遅い部分→ff」に分割して記録するケースがある。
          // showKyuiMap[yr][0] は低速 ff 値で計算されているため球威が過小になる。
          // BIS-split 検知時は si の高速値（_bgVeloNum）で 球威 を再計算して上書きする。
          if (_bisSplitDetected) {
            const est = calcKyuiPreStatcast(_bgVeloNum, pg.idx, pctNum, perfEra, perfBaa1000, perfHr9);
            if (est !== '') kyui = Math.max(30, Math.min(110, Number(est) + perfBoost));
          } else {
          // フォールバック①: showKyuiMap（The Show / FanGraphs / 推定データ）+ 補正
          const showKyui = showKyuiMap[yr]?.[pg.idx];
          if (showKyui !== undefined) {
            kyui = Math.max(30, Math.min(110, Number(showKyui) + perfBoost));
          } else if (!isNaN(veloNum) && veloNum > 0) {
            // フォールバック②: showKyuiMap にキーがない場合（未取得年・データ欠損）→ 球速 + 補正で推定
            const est = calcKyuiPreStatcast(veloNum, pg.idx, pctNum, perfEra, perfBaa1000, perfHr9);
            if (est !== '') kyui = est;
          }
          }
          // ── 球威キャップ（③以降推定年: 投球回数・防御率による上限制約）─────────────────
          // ①②(Baseball Savant 実測) は BA/SLG あり → if(ah!==''&&ai!=='') ブランチを通るため制約なし。
          // ③(推定)/④(FanGraphs)/⑤(Claude+AgingCurve) は BA/SLG='--' → このブランチで制約適用。
          if (kyui !== '') {
            const eraNum = !isNaN(perfEra) ? perfEra : NaN;
            if (!isNaN(eraNum) && eraNum >= 0) {
              const A = Number(kyui);
              let base = 95, threshold = 95, mult = null;

              if (ip >= 200 && ip <= 270) {
                if      (eraNum >= 3.70)              { threshold = 95; mult = 0.4; }
                else if (eraNum >= 2.70)              { threshold = 95; mult = 0.5; }
                else if (eraNum >= 1.70)              { threshold = 95; mult = 0.6; }
              } else if (ip >= 151 && ip < 200) {
                if      (eraNum >= 3.30)              { threshold = 95; mult = 0.4; }
                else if (eraNum >= 2.30)              { threshold = 95; mult = 0.5; }
                else if (eraNum >= 1.30)              { threshold = 95; mult = 0.6; }
              } else if (ip >= 100 && ip <= 150) {
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
        // ── Wikipedia 決め球 球威補正 ─────────────────────────────────────────────
        // Wikipedia テキストから「weapon/signature/out-pitch」等の語句近傍に出現する
        // 変化球が決め球として検出された場合、その球種の球威に+5 を加算する。
        // 対象: Statcast年(BA/SLG実測)・非Statcast年ともに適用。
        // 例: クレメンスのスプリット(idx=6) → 高速スプリットの落差球威を正しく反映
        if (kyui !== '' && outPitchBoosts[pg.idx]) {
          kyui = Math.max(30, Math.min(110, Number(kyui) + outPitchBoosts[pg.idx]));
        }
        // 蓄積: 通算加重平均に使用
        if (kyui !== '') _rowKyuiByI[i] = Number(kyui);
      }
      if (kyui !== '') redPurpleCell(ws.getCell(rn, base + 1), kyui, fontSize);

      redPurpleCell(ws.getCell(rn, base + 2), pctNum, fontSize);
    });

    // 年度行のみ kyuiStore に蓄積（通算行はスキップ）
    if (!isCareerRow && Object.keys(_rowKyuiByI).length > 0) {
      yearKyuiStore.push({ outs: ipToOuts(ipRaw), kyuiByI: _rowKyuiByI });
    }

    count++;
  }

  await wb.xlsx.writeFile(xlsxPath);
  return count;
}

// ── チーム略称の正規化テーブル ────────────────────────────────────────────────
// MLB Stats API が返す略称はシーズン・移転により揺れがあるため統一する
const TEAM_ABBR_NORMALIZE = {
  // ── 略称の表記揺れ ────────────────────────────────────────────────────────
  'TBD': 'TB',  'TBR': 'TB',  'TBA': 'TB',          // Rays
  'KCR': 'KC',                                        // Royals
  'CHW': 'CWS',                                       // White Sox
  'ANA': 'LAA',  'CAL': 'LAA',                        // Angels
  'OAK': 'ATH',                                       // Athletics
  'FLA': 'MIA',  'FLO': 'MIA',                        // Marlins
  'MON': 'WSH',  'WSN': 'WSH',  'WAS': 'WSH',         // Nationals
  'SDP': 'SD',                                        // Padres
  'SFG': 'SF',  'SFN': 'SF',                          // Giants
  'SLN': 'STL',                                        // Cardinals
  // ── フルチーム名 → 略称（APIがabbreviationを返さない場合のフォールバック）──
  'New York Yankees':               'NYY',
  'New York Mets':                  'NYM',
  'Los Angeles Angels':             'LAA',
  'Los Angeles Angels of Anaheim':  'LAA',
  'Anaheim Angels':                 'LAA',
  'California Angels':              'LAA',
  'Los Angeles Dodgers':            'LAD',
  'Chicago Cubs':                   'CHC',
  'Chicago White Sox':              'CWS',
  'Boston Red Sox':                 'BOS',
  'Baltimore Orioles':              'BAL',
  'Tampa Bay Rays':                 'TB',
  'Tampa Bay Devil Rays':           'TB',
  'Toronto Blue Jays':              'TOR',
  'Cleveland Guardians':            'CLE',
  'Cleveland Indians':              'CLE',
  'Detroit Tigers':                 'DET',
  'Kansas City Royals':             'KC',
  'Minnesota Twins':                'MIN',
  'Houston Astros':                 'HOU',
  'Oakland Athletics':              'ATH',
  'Sacramento Athletics':           'ATH',
  'Seattle Mariners':               'SEA',
  'Texas Rangers':                  'TEX',
  'Atlanta Braves':                 'ATL',
  'Miami Marlins':                  'MIA',
  'Florida Marlins':                'MIA',
  'Philadelphia Phillies':          'PHI',
  'Washington Nationals':           'WSH',
  'Montreal Expos':                 'WSH',
  'Arizona Diamondbacks':           'ARI',
  'Colorado Rockies':               'COL',
  'San Diego Padres':               'SD',
  'San Francisco Giants':           'SF',
  'Cincinnati Reds':                'CIN',
  'Milwaukee Brewers':              'MIL',
  'Pittsburgh Pirates':             'PIT',
  'St. Louis Cardinals':            'STL',
};
function normalizeTeamAbbr(raw) {
  if (!raw) return raw;
  const mapped = TEAM_ABBR_NORMALIZE[raw];
  if (mapped) return mapped;
  if (raw.length <= 3) return raw;
  // 未登録のフル名は先頭3文字で代替（例: "Boston Braves" → "BOS"）
  return raw.slice(0, 3).toUpperCase();
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
      teamStr = normalizeTeamAbbr(primary.team?.abbreviation || primary.team?.name || '???') + named.length;
    } else {
      teamStr = normalizeTeamAbbr(row.team?.abbreviation || row.team?.name || '???');
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

  // ── 球速の抽出: fastball 関連の mph 値を優先して取得 ─────────────────────────
  // ★ km/h 変換は廃止: 丸め誤差で 93mph スプリッター→94mph と誤判定しキャップが
  //   誤発動する問題があったため削除。mph / miles per hour 表記のみ対応。
  //
  // 抽出戦略:
  //   ① fastball 関連キーワード近傍（前後 200 文字）の mph 値 → fastballVeloMentions
  //   ② それ以外の全 mph 値                                    → allVeloMentions
  // 優先度: ① > ②（① があれば ① の MAX を使用、なければ ② の MAX）
  // これにより "93mph splitter" は除外し "100mph fastball" だけをキャップに使える。
  const _FASTBALL_WORDS = [
    'fastball', 'four-seam', '4-seam', 'four seam', '4 seam', 'two-seam', '2-seam',
    'heater', 'heat', 'straight', 'sinker', 'sinking fastball', 'gas', 'fireball',
  ];
  const _VPH_RE = /(\d{2,3})\s*(?:[-–to]+\s*(\d{2,3}))?\s*(?:mph|miles per hour)/gi;
  const _rawVelos = []; // { mph, pos }
  for (const m of text.matchAll(_VPH_RE)) {
    const v1 = parseInt(m[1]), v2 = m[2] ? parseInt(m[2]) : null;
    if (v1 >= 80 && v1 <= 106) _rawVelos.push({ mph: v1, pos: m.index });
    if (v2 && v2 >= 80 && v2 <= 106) _rawVelos.push({ mph: v2, pos: m.index });
  }
  const allVeloMentions = _rawVelos.map(v => v.mph);
  // fastball 関連語句の近傍（前後 200 文字以内）に出現する速度のみ抽出
  const fastballVeloMentions = _rawVelos
    .filter(({ mph, pos }) => _FASTBALL_WORDS.some(w =>
      lower.substring(Math.max(0, pos - 200), pos + 200).includes(w)
    ))
    .map(v => v.mph);
  // 最終的な veloMentions: fastball 関連があればそれ優先、なければ全体
  const veloMentions = fastballVeloMentions.length > 0 ? fastballVeloMentions : allVeloMentions;

  // ── 決め球（out pitch / signature pitch）検出 ──────────────────────────────
  // "weapon", "signature", "best pitch", "out pitch" 等の語句から 200 文字以内に
  // 変化球（ff/si 以外）の球種ワードが出現した場合を「決め球」と判定する。
  // 速球（ff/si）は除外: 速球はほぼ全投手の「主球種」であり、決め球扱いには馴染まない。
  const _OUT_INDICATORS = [
    'out pitch', 'out-pitch', 'put-away', 'put away', 'signature', 'best pitch',
    'weapon', 'devastating', 'go-to pitch', 'go to pitch', 'trademark',
    'relied on', 'relied heavily', 'primary weapon', 'most effective',
    'known for', 'feared for', 'deadliest', 'crown jewel', 'bread and butter',
  ];
  const _OUT_WINDOW = 200; // 近傍判定ウィンドウ (文字数)
  let outPitchKey = null;

  // インジケーターの全出現位置を収集
  const _indPos = [];
  for (const word of _OUT_INDICATORS) {
    let pos = 0;
    while ((pos = lower.indexOf(word, pos)) !== -1) { _indPos.push(pos); pos += word.length; }
  }

  if (_indPos.length > 0) {
    let _minDist = _OUT_WINDOW + 1;
    for (const { key, words } of WIKI_PITCH_PATTERNS) {
      if (key === 'ff' || key === 'si') continue; // 速球は対象外
      if (!pitchCounts[key]) continue;            // テキスト中に出現しない球種はスキップ
      for (const word of words) {
        let pos = 0;
        while ((pos = lower.indexOf(word, pos)) !== -1) {
          for (const ip of _indPos) {
            const dist = Math.abs(pos - ip);
            if (dist < _minDist) { _minDist = dist; outPitchKey = key; }
          }
          pos += word.length;
        }
      }
    }
  }

  if (pitchKeys.length === 0 && veloMentions.length === 0) return null;
  return { pitchKeys, primaryKey, pitchCounts, veloMentions, outPitchKey };
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

    // ── Step 2: 記事テキストを取得（最大 50000 文字）──
    // クレメンス・ジョンソン等の長大記事では「Pitching style」セクションが
    // 25000 文字以降に来る場合がある。50000 文字まで拡張して球速言及を確実に捕捉する。
    const qt = encodeURIComponent(bestTitle);
    const extractRes = await wikiGet(
      `https://en.wikipedia.org/w/api.php?action=query&titles=${qt}&prop=extracts&explaintext=true&format=json&origin=*`
    ).catch(() => null);

    const pages = extractRes?.query?.pages;
    if (!pages) return null;
    const page = Object.values(pages)[0];
    if (!page || page.missing !== undefined) return null;

    const fullText = (page.extract || '').slice(0, 50000);
    if (!fullText) return null;

    const profile = parsePitchProfile(fullText);
    if (!profile) return null;
    return { ...profile, pageTitle: page.title };
  } catch {
    return null;
  }
}

/**
 * 日本語 Wikipedia の「選手としての特徴」等のセクションから球速・球種情報を抽出する。
 * 英語 Wikipedia より正確な最高球速（"最速100mph" 形式）を記述していることが多い。
 *
 * ★ 実装方針:
 *   長大な記事（経歴セクションが多い選手）では全文の30000文字以内に対象セクションが
 *   収まらない場合がある。そのため MediaWiki の action=parse&prop=sections API で
 *   セクション番号を特定し、action=parse&section=N&prop=wikitext で該当セクションの
 *   テキストだけを取得する「セクション直接取得」方式を採用する。
 *   これにより記事の全長に依存せず確実に対象セクションを抽出できる。
 *
 * @param {string} searchName  日本語選手名（カタカナ可）または英語名
 * @returns {{ veloMentions:number[], outPitchKey:string|null, pageTitle:string, sectionText:string } | null}
 */
async function fetchJaWikiCharSection(searchName) {
  // searchName はオブジェクト { _jaTitle: "..." } または文字列のどちらも受け付ける
  if (!searchName) return null;
  if (typeof searchName === 'object' && !searchName._jaTitle) return null;
  if (typeof searchName === 'string' && !searchName.trim()) return null;

  const jaWikiGet = url => new Promise((resolve, reject) => {
    const mod = url.startsWith('https') ? require('https') : require('http');
    mod.get(url, {
      headers: {
        'User-Agent': 'MLB-PitchTool/1.0 (Node.js; baseball stats research)',
        'Accept': 'application/json',
      },
    }, res => {
      if (res.statusCode === 301 || res.statusCode === 302) {
        const loc = res.headers.location;
        if (loc) return jaWikiGet(loc).then(resolve).catch(reject);
        return reject(new Error('redirect without location'));
      }
      let buf = '';
      res.on('data', c => buf += c);
      res.on('end', () => {
        try { resolve(JSON.parse(buf)); } catch (e) { reject(e); }
      });
    }).on('error', reject);
  });

  // 対象セクション名（日本語 Wikipedia の投手記事で使われるセクション）
  const TARGET_SECTIONS = [
    '選手としての特徴', '投球スタイル', '投手としての特徴', 'プレースタイル',
    '特徴', '球種', '投球',
  ];

  // 日本語テキストから mph 値を抽出するパターン
  const _JA_MPH_RE = /(\d{2,3})\s*mph/gi;

  // 日本語球種ワード → PITCH_KEYS マッピング
  const JA_PITCH_WORDS = [
    { key: 'ff', words: ['フォーシーム', '4シーム', '速球', '直球', 'ストレート', 'ファストボール'] },
    { key: 'si', words: ['シンカー', 'ツーシーム', '2シーム'] },
    { key: 'sl', words: ['スライダー', 'スイーパー'] },
    { key: 'ch', words: ['チェンジアップ', 'スクリューボール', 'パームボール'] },
    { key: 'cu', words: ['カーブ', 'ナックルカーブ'] },
    { key: 'fc', words: ['カットボール', 'カッター'] },
    { key: 'fs', words: ['スプリッター', 'スプリット', 'フォークボール', 'フォーク', 'ナックルボール', 'ナックル'] },
  ];

  // 日本語 決め球 インジケーター
  const JA_OUT_INDICATORS = [
    '最大の武器', '代名詞', '得意とした', '決め球', '主武器', '看板', '持ち味',
    '最も恐れられた', '最も得意', '強みとした', '主要な武器',
  ];

  // wikitext から mph を抽出してスプリット結果を返す
  const extractFromText = (text) => {
    const veloMentions = [];
    for (const m of text.matchAll(_JA_MPH_RE)) {
      const v = parseInt(m[1]);
      if (v >= 80 && v <= 106) veloMentions.push(v);
    }
    let outPitchKey = null;
    let _minDist = 300;
    for (const word of JA_OUT_INDICATORS) {
      let pos = 0;
      while ((pos = text.indexOf(word, pos)) !== -1) {
        for (const { key, words: jaPitchWords } of JA_PITCH_WORDS) {
          if (key === 'ff' || key === 'si') continue;
          for (const pw of jaPitchWords) {
            let ppos = 0;
            while ((ppos = text.indexOf(pw, ppos)) !== -1) {
              const dist = Math.abs(ppos - pos);
              if (dist < _minDist) { _minDist = dist; outPitchKey = key; }
              ppos += pw.length;
            }
          }
        }
        pos += word.length;
      }
    }
    return { veloMentions, outPitchKey };
  };

  try {
    // ── Step 1: 日本語 Wikipedia の記事タイトルを決定 ──
    // ★ jaTitle が直接渡された場合はそのまま使用（opensearch をスキップ）。
    //   英語名での opensearch は日本語 Wikipedia では機能しないため、
    //   呼び出し元で langlinks API 等を使って正確なタイトルを取得して渡すことを推奨。
    let bestTitle = (typeof searchName === 'object' && searchName._jaTitle) ? searchName._jaTitle : null;

    if (!bestTitle) {
      // フォールバック: opensearch でカタカナ名を探す（日本語名のみ有効）
      let titles = [];
      // カタカナのみのクエリを優先（英語名は機能しないためスキップ）
      const hasJapanese = /[ぁ-ヿ]/.test(typeof searchName === 'string' ? searchName : '');
      const queries = hasJapanese
        ? [searchName + ' 野球', searchName]
        : []; // 英語名のみの場合は opensearch をスキップ（必ず失敗するため）
      for (const q of queries) {
        const res = await jaWikiGet(
          `https://ja.wikipedia.org/w/api.php?action=opensearch&search=${encodeURIComponent(q)}&limit=5&namespace=0&format=json&origin=*`
        ).catch(() => null);
        titles = res?.[1] ?? [];
        if (titles.length) break;
      }
      if (!titles.length) return null;
      bestTitle = titles[0];
    }

    // ★ 念のため: タイトルが searchName と完全一致（入力名が既に正確なタイトル）の場合はそのまま
    const _plainSearch = typeof searchName === 'string' ? searchName : '';
    if (!bestTitle && _plainSearch) bestTitle = _plainSearch;
    const qt = encodeURIComponent(bestTitle);

    // ── Step 2: action=parse&prop=sections でセクション一覧を取得 ──
    // 記事全文を取得せず目次だけ取得するため高速・軽量。
    const sectionsRes = await jaWikiGet(
      `https://ja.wikipedia.org/w/api.php?action=parse&page=${qt}&prop=sections&format=json&origin=*`
    ).catch(() => null);

    const sections = sectionsRes?.parse?.sections ?? [];
    const pageTitle = sectionsRes?.parse?.title ?? bestTitle;

    // ── Step 3: 対象セクション番号を特定 ──
    let targetSectionIndices = [];
    for (const sec of sections) {
      if (TARGET_SECTIONS.includes(sec.line)) {
        targetSectionIndices.push(sec.index); // "1", "2", ... の文字列
      }
    }

    if (targetSectionIndices.length === 0) {
      // セクション見つからず → 全文フォールバック（exlimit を大きく設定）
      const extractRes = await jaWikiGet(
        `https://ja.wikipedia.org/w/api.php?action=query&titles=${qt}&prop=extracts&explaintext=true&exlimit=1&format=json&origin=*`
      ).catch(() => null);
      const pages = extractRes?.query?.pages;
      const page = pages ? Object.values(pages)[0] : null;
      const fullText = page?.extract ?? '';
      if (!fullText) return null;
      const { veloMentions, outPitchKey } = extractFromText(fullText);
      if (veloMentions.length === 0 && !outPitchKey) return null;
      return { veloMentions, outPitchKey, pageTitle, sectionText: fullText.slice(0, 3000) };
    }

    // ── Step 4: 対象セクションの wikitext を取得（セクション番号指定）──
    // action=parse&section=N&prop=wikitext は該当セクションのみ返すため高速。
    // wikitext には [[リンク]] や {{テンプレート}} が含まれるが mph 値は生テキストとして
    // 含まれるため、専用の抽出ロジックで問題なく取得できる。
    let combinedText = '';
    for (const idx of targetSectionIndices) {
      const secRes = await jaWikiGet(
        `https://ja.wikipedia.org/w/api.php?action=parse&page=${qt}&section=${idx}&prop=wikitext&format=json&origin=*`
      ).catch(() => null);
      const wikitext = secRes?.parse?.wikitext?.['*'] ?? '';
      if (wikitext) combinedText += '\n' + wikitext;
    }

    if (!combinedText.trim()) return null;

    const { veloMentions, outPitchKey } = extractFromText(combinedText);
    if (veloMentions.length === 0 && !outPitchKey) return null;
    return { veloMentions, outPitchKey, pageTitle, sectionText: combinedText };
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
async function pitFetchBrowserData(slug, id, years, onProgress, playerName = '', apiKey = '', englishName = '', basic = {}) {
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
    // Wikipedia 球種プロファイル（2a.2 で取得 → 2d.5 で活用）
    let wikiProfile = null;
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
        // ※ wikiProfile は外側スコープで宣言済み（Step 2d.5 でも参照するため）
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
                  // 実際の mph→km/h 変換でのキャップ値（BG列表示上限に使用）
                  // ゲームスケール(×1.6+4)ではなく実換算(×1.60934)で表示上限を設定
                  // 例: 100mph → 161km/h（ゲームスケールなら164だが実際は161）
                  wikiProfile._capKmh  = Math.round(wikiPeak * 1.60934);
                  onProgress(`[2a.2 Wikipedia] 最高球速参考値: ${primaryFfKey}=${wikiPeak}mph (MAX ${wikiProfile._capKmh}km/h)（BG上限・2a.3未実行時キャップに使用）`);
                }
                // ── 決め球（out pitch）検出 ──────────────────────────────────────────
                if (wikiProfile.outPitchKey) {
                  wikiProfile._outPitchKey = wikiProfile.outPitchKey;
                  const outJa = { sl:'スライダー', ch:'チェンジアップ', cu:'カーブ', fc:'カット', fs:'スプリット' }[wikiProfile.outPitchKey] ?? wikiProfile.outPitchKey;
                  onProgress(`[2a.2 Wikipedia] 決め球検出: ${outJa}(${wikiProfile.outPitchKey}) — 球威に+5補正を適用予定`);
                }
              } else {
                onProgress('[2a.2 Wikipedia] 球種情報が見つかりませんでした（記事なし、または球種の記述なし）');
              }
            } catch (e) {
              onProgress('⚠ Wikipedia 取得エラー: ' + e.message);
            }
          }
        }

        // ── 2a.2ja: 日本語 Wikipedia「選手としての特徴」セクション補正 ────────────
        // 英語 Wikipedia は「94mph fastball」のようにシーズン平均球速を記述することが多く、
        // キャリア最高球速（例: クレメンス 100mph）を取りこぼす場合がある。
        // 日本語 Wikipedia の「選手としての特徴」セクションは「最速100mph」等と
        // 明示的に最高球速を記述する傾向があるため、英語版の補正・上書きに活用する。
        // ■ 優先度: 日本語 Wikipedia ≥ 英語 Wikipedia (jaPeak >= enPeak の場合のみ上書き)
        // ■ 日本語タイトル取得方法:
        //   ① 英語Wikipedia の langlinks API → en記事と紐づく ja記事タイトルを直接取得
        //      (例: "Roger Clemens" → "ロジャー・クレメンス")
        //   ② playerName がカタカナを含む場合は直接タイトルとして使用
        //   opensearch は英語名では機能しないため使用しない
        {
          // ─ ステップ①: 英語Wikiの langlinks から日本語タイトルを取得 ─
          const _enTitle = wikiProfile?.pageTitle || englishName;
          let _jaTitle = null;

          if (_enTitle) {
            try {
              const _llUrl = `https://en.wikipedia.org/w/api.php?action=query&titles=${encodeURIComponent(_enTitle)}&prop=langlinks&lllang=ja&format=json&origin=*`;
              const _llRes = await new Promise((resolve, reject) => {
                require('https').get(_llUrl, { headers: { 'User-Agent': 'MLB-PitchTool/1.0', 'Accept': 'application/json' } }, res => {
                  let b = ''; res.on('data', c => b += c); res.on('end', () => { try { resolve(JSON.parse(b)); } catch(e) { reject(e); } });
                }).on('error', reject);
              }).catch(() => null);
              const _llPages = _llRes?.query?.pages;
              if (_llPages) {
                const _llPage = Object.values(_llPages)[0];
                _jaTitle = _llPage?.langlinks?.[0]?.['*'] ?? null;
              }
            } catch { /* langlinks 取得失敗 */ }
          }

          // ─ ステップ②: カタカナ名フォールバック ─
          // playerName がカタカナを含む場合（例: "ロジャー・クレメンス"）は直接使用
          if (!_jaTitle) {
            const _hasKata = /[゠-ヿ]/.test(playerName || '');
            if (_hasKata) _jaTitle = playerName.trim();
          }

          if (_jaTitle) {
            onProgress(`日本語 Wikipedia「選手としての特徴」検索中: "${_jaTitle}"...`);
            try {
              const jaProfile = await fetchJaWikiCharSection({ _jaTitle });
              if (jaProfile && jaProfile.veloMentions.length > 0) {
                const jaPeak = Math.max(...jaProfile.veloMentions);
                const enPeak = wikiProfile?._capPeak ?? 0;
                onProgress(`[2a.2ja 日本語Wikipedia] "${jaProfile.pageTitle}": 最高球速言及 ${jaPeak}mph (英語版: ${enPeak}mph)`);
                if (jaPeak >= enPeak && jaPeak >= 93) {
                  if (!wikiProfile) wikiProfile = { pitchKeys: [], primaryKey: null, pitchCounts: {}, veloMentions: [], outPitchKey: null };
                  wikiProfile._capKey  = wikiProfile._capKey || 'ff';
                  wikiProfile._capPeak = jaPeak;
                  wikiProfile._capKmh  = Math.round(jaPeak * 1.60934);
                  onProgress(`[2a.2ja ✓] 日本語Wikipedia 最高球速 ${jaPeak}mph (${wikiProfile._capKmh}km/h) を採用 — 英語版(${enPeak}mph)より正確`);
                } else if (jaPeak < enPeak) {
                  onProgress(`[2a.2ja ℹ] 日本語Wikipedia ${jaPeak}mph < 英語版 ${enPeak}mph のため英語版を維持`);
                }
                if (jaProfile.outPitchKey && !wikiProfile?._outPitchKey) {
                  if (!wikiProfile) wikiProfile = { pitchKeys: [], primaryKey: null, pitchCounts: {}, veloMentions: [], outPitchKey: null };
                  wikiProfile._outPitchKey = jaProfile.outPitchKey;
                  const outJa = { sl:'スライダー', ch:'チェンジアップ', cu:'カーブ', fc:'カット', fs:'スプリット' }[jaProfile.outPitchKey] ?? jaProfile.outPitchKey;
                  onProgress(`[2a.2ja Wikipedia] 決め球検出(日本語): ${outJa}(${jaProfile.outPitchKey}) — 球威に+5補正を適用予定`);
                }
              } else {
                onProgress('[2a.2ja 日本語Wikipedia] 対象セクションで球速情報が見つかりませんでした');
              }
            } catch (e) {
              onProgress('⚠ 日本語Wikipedia 取得エラー: ' + e.message);
            }
          } else {
            onProgress('[2a.2ja 日本語Wikipedia] 日本語タイトルを特定できませんでした（langlinks なし・カタカナ名なし）');
          }
        }

        // ── 2a.2b: ※ 無効化 ──────────────────────────────────────────────────────
        // このステップは aging curve (Step 2d) より前に実行されるため、
        // 2a.2b で si をクリアしても Step 2d が si を再推定して上書きしてしまう問題があった。
        // また 2008-2009 等の FG si 実測値（aging curve の参照年）を早期に削除すると
        // aging curve の参照年が消えて推定精度が低下する。
        // → si→ff 速度転写は aging curve 完了後の Step 2d.5 で実施する。

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
        // ★ capPeak < 93mph（≒150km/h未満）はキャップ不適用。
        //   Wikipedia が変化球・晩年球速のみ言及する場合に速球が誤って下げられるのを防ぐ。
        if (!apiKey && wikiProfile?._capPeak && wikiProfile?._capPeak >= 93 && wikiProfile?._capKey) {
          const capKey  = wikiProfile._capKey;
          const capPeak = wikiProfile._capPeak;        // Wikipedia 最高球速 (mph)
          const seasonAvgCap = capPeak - 3.7;          // 瞬間最大 → シーズン平均上限
          const ki = PITCH_KEYS.indexOf(capKey);
          let wikiCapCount = 0;
          const fgFilledYearsForWiki = fgTargetYears.filter(yr => yearHasPct(yr) && +yr < 2015);
          for (const yr of fgFilledYearsForWiki) {
            // ── capKey(ff) のキャップ適用 ──
            const d = rawPitch[yr]?.[capKey];
            if (d && d.velo !== '--') {
              const v = parseFloat(d.velo);
              if (!isNaN(v) && v > seasonAvgCap) {
                rawPitch[yr][capKey].velo = String(+seasonAvgCap.toFixed(1));
                const byr = basic[yr];
                const bEra = byr ? parseFloat(String(byr.era)) : NaN;
                const bBaa = byr ? (byr.avg ? Number(byr.avg) * 1000 : NaN) : NaN;
                const bIp  = byr ? parseFloat(String(byr.ip).replace(/\.(\d)$/, '.$10')) : 0;
                const bHr9 = (byr && bIp > 0) ? (byr.hr * 9 / bIp) : NaN;
                const pctNum = parseFloat(String(d.pct).replace('%', ''));
                const kyui = calcKyuiPreStatcast(seasonAvgCap, ki, isNaN(pctNum) ? 20 : pctNum, bEra, bBaa, bHr9);
                if (kyui !== '') { if (!showKyuiMap[yr]) showKyuiMap[yr] = {}; showKyuiMap[yr][ki] = kyui; }
                wikiCapCount++;
              }
            }
            // ── BIS-split 対策: si が ff より 4mph 以上速い年のみ si もキャップ ──────
            // FanGraphs BIS は「速い部分→si」「遅い部分→ff」に分割するため si が実際の
            // ファストボールより高く記録される。この場合のみ si にも同じキャップを適用する。
            // 通常のシンカー投手（si ≈ ff）は対象外にして誤キャップを防ぐ。
            if (capKey === 'ff') {
              const ffD = rawPitch[yr]?.['ff'];
              const siD = rawPitch[yr]?.['si'];
              if (siD && siD.velo !== '--' && ffD && ffD.velo !== '--') {
                const siV = parseFloat(siD.velo);
                const ffV = parseFloat(ffD.velo);
                // si > ff + 4mph → BIS-split と判断してsiもキャップ
                if (!isNaN(siV) && !isNaN(ffV) && siV > ffV + 4 && siV > seasonAvgCap) {
                  const siKi = PITCH_KEYS.indexOf('si');
                  rawPitch[yr]['si'].velo = String(+seasonAvgCap.toFixed(1));
                  const byr = basic[yr];
                  const bEra = byr ? parseFloat(String(byr.era)) : NaN;
                  const bBaa = byr ? (byr.avg ? Number(byr.avg) * 1000 : NaN) : NaN;
                  const bIp  = byr ? parseFloat(String(byr.ip).replace(/\.(\d)$/, '.$10')) : 0;
                  const bHr9 = (byr && bIp > 0) ? (byr.hr * 9 / bIp) : NaN;
                  const siPct = parseFloat(String(siD.pct).replace('%', ''));
                  const siKyui = calcKyuiPreStatcast(seasonAvgCap, siKi, isNaN(siPct) ? 20 : siPct, bEra, bBaa, bHr9);
                  if (siKyui !== '') { if (!showKyuiMap[yr]) showKyuiMap[yr] = {}; showKyuiMap[yr][siKi] = siKyui; }
                  wikiCapCount++;
                }
              }
            }
          }
          if (wikiCapCount > 0) {
            onProgress(`[2a.2 Wikipedia キャップ] ${capKey}(+BIS-split si) 最高球速 ${capPeak}mph → シーズン平均上限 ${seasonAvgCap.toFixed(1)}mph 適用 (${wikiCapCount}年)`);
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

    // ── 2d.5: Wikipedia primaryKey='ff' 選手の si→ff 速度転写（aging curve 後）────
    // aging curve 完了後に実施することで、FG si 実測年（2008-2009等）が aging curve の
    // 参照年として正しく機能した上で、最終的に si の速度を ff に転写できる。
    // 条件: Wikipedia primaryKey='ff' かつ si が ff より 3mph 以上速い年
    // 処理: rawPitch[yr]['ff'].velo ← si の速度（ff の pct は保持）、si を無効化
    if (wikiProfile?.primaryKey === 'ff') {
      let d5Count = 0;
      for (const yr of years.filter(y => y !== '通算')) {
        const ffD = rawPitch[yr]?.['ff'];
        const siD = rawPitch[yr]?.['si'];
        if (!siD || siD.velo === '--') continue;
        const siVelo = parseFloat(siD.velo);
        if (isNaN(siVelo) || siVelo <= 0) continue;
        const ffVelo = (ffD && ffD.velo !== '--') ? parseFloat(ffD.velo) : 0;
        if (siVelo <= ffVelo + 3) continue; // si が大幅に速くない年はスキップ

        // ff の速度を si の速度で上書き（投球割合は ff のものを保持）
        const ffPct = (ffD && ffD.pct !== '--') ? ffD.pct : siD.pct;
        rawPitch[yr]['ff'] = {
          velo: siD.velo,
          ba:   ffD?.ba  ?? '--',
          slg:  ffD?.slg ?? '--',
          pct:  ffPct,
        };
        rawPitch[yr]['si'] = { velo: '--', ba: '--', slg: '--', pct: '--' };

        // showKyuiMap: si(idx=5) のエントリを削除し ff(idx=0) を新速度で再計算
        if (showKyuiMap[yr]) {
          delete showKyuiMap[yr][5];
          const byr  = basic[yr];
          const bEra = byr ? parseFloat(String(byr.era)) : NaN;
          const bBaa = byr ? (byr.avg ? Number(byr.avg) * 1000 : NaN) : NaN;
          const bIp  = byr ? parseFloat(String(byr.ip).replace(/\.(\d)$/, '.$10')) : 0;
          const bHr9 = (byr && bIp > 0) ? (byr.hr * 9 / bIp) : NaN;
          const pctNum = parseFloat(String(ffPct).replace('%', ''));
          const newKyui = calcKyuiPreStatcast(siVelo, 0, isNaN(pctNum) ? 20 : pctNum, bEra, bBaa, bHr9);
          if (newKyui !== '') showKyuiMap[yr][0] = newKyui;
        }

        const siKmh = Math.round(siVelo * 1.6 + 4);
        const ffKmh = ffVelo > 0 ? Math.round(ffVelo * 1.6 + 4) : '--';
        onProgress(`[2d.5 si→ff転写] ${yr}: si=${siKmh}km/h → ff(旧${ffKmh}km/h) 速度転写、割合=${ffPct}%保持`);
        d5Count++;
      }
      if (d5Count > 0) {
        onProgress(`[2d.5] ${d5Count}年でsi→ff速度転写完了（aging curve後・Wikipedia primaryKey='ff'）`);
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

    return { rawPitch, showKyuiMap, pitchNameOverrides, wikiProfile };
  } finally {
    await browser.close();
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch {}
  }
}

// ── Excel build ───────────────────────────────────────────────────────────────
async function pitBuildExcel(playerName, years, basic, vsLeftByYear, rawPitch) {
  const N_MAIN = 22;
  const N_SUB  = 4;

  // Career pitch data: アウト数加重平均
  const careerPitch = emptyPitchP();

  // ── 通算 pct 専用: キャリア全体のアウト数を分母に使用 ────────────────────────
  // 【修正理由】球種ごとに「velo が有効な年だけ」を分母にすると、
  // 途中から球種を変えた投手（例: RA Dickey のナックルボール転向）で
  // 分母が球種によって異なり、通算割合が大きく歪む。
  // 正しい通算 pct = Σ(全年の pct × アウト数) / Σ(キャリア全体のアウト数)
  // 未使用年（pct = '--'）は 0% として扱い分母には含める。
  const _totalCareerOuts = years.reduce((s, yr) => s + ipToOuts(basic[yr]?.ip || '0'), 0);
  const _careerPctRaw = {};
  if (_totalCareerOuts > 0) {
    for (const key of PITCH_KEYS) {
      const sumPctOuts = years.reduce((s, yr) => {
        const outs = ipToOuts(basic[yr]?.ip || '0');
        if (!outs) return s;
        const d = rawPitch[yr]?.[key];
        if (!d || d.pct == null || d.pct === '--') return s; // 未使用年は 0%
        const pct = parseFloat(String(d.pct).replace('%', ''));
        if (isNaN(pct) || pct <= 0) return s;
        return s + pct * outs;
      }, 0);
      const avg = sumPctOuts / _totalCareerOuts;
      _careerPctRaw[key] = avg >= 0.05 ? String(avg.toFixed(1)) : '--';
    }
  }

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
      // pct は全キャリアのアウト数を分母にした値を使用（球種ごとの分母バグを修正）
      pct:  _careerPctRaw[key] ?? '--',
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

async function pitRunCreateJob(jobId, params) {
  const upd = msg => { const j = jobs.get(jobId); if (j) { j.progress = msg; console.log('[job]', msg); } };
  try {
    upd('MLB Stats API から投手成績を取得中...');
    const { years, basic, vsLeftByYear } = await fetchPitchingStats(params.id, params.y1, params.y2);

    upd('ブラウザを起動して Baseball Savant / MLB The Show から球種データを取得中...');
    const apiKey = params.apiKey || process.env.ANTHROPIC_API_KEY || '';
    const { rawPitch, showKyuiMap, pitchNameOverrides, wikiProfile } = await pitFetchBrowserData(params.slug, params.id, years, upd, params.name, apiKey, params.fullName || '', basic);

    upd('Excel ファイルを生成中...');
    const outFile = await pitBuildExcel(params.name, years, basic, vsLeftByYear, rawPitch);

    upd('スタミナ・制球を計算中...');
    let abilityRows = 0;
    try {
      // ── Wikipedia 決め球ブースト・BG km/h キャップを extraOptions に組み立て ──
      const _outPitchBoosts = {};
      if (wikiProfile?._outPitchKey) {
        const _outIdx = PITCH_KEYS.indexOf(wikiProfile._outPitchKey);
        if (_outIdx >= 0) {
          _outPitchBoosts[_outIdx] = 5; // 決め球 +5
          const _outJa = { sl:'スライダー', ch:'チェンジアップ', cu:'カーブ', fc:'カット', fs:'スプリット' }[wikiProfile._outPitchKey] ?? wikiProfile._outPitchKey;
          upd(`[決め球補正] ${_outJa}(idx=${_outIdx}) 球威 +5 を全年度に適用`);
        }
      }
      const _wikiCapKmh = wikiProfile?._capKmh ?? null;
      if (_wikiCapKmh) upd(`[BG上限] Wikipedia最高球速 ${wikiProfile._capPeak}mph → BG列上限 ${_wikiCapKmh}km/h`);

      abilityRows = await addAbilityToFile(outFile, showKyuiMap, pitchNameOverrides, {
        outPitchBoosts: _outPitchBoosts,
        wikiCapKmh: _wikiCapKmh,
      });
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
async function batFetchMLBStats(id, y1, y2) {
  const yby = await mlbGet(
    `https://statsapi.mlb.com/api/v1/people/${id}/stats?stats=yearByYear&group=hitting&sportId=1`
  );
  const allSplits = (yby.stats[0]?.splits || []).filter(s => s.sport?.id === 1);
  const byYear = {};
  for (const s of allSplits) {
    const yr = s.season;
    if (!byYear[yr]) byYear[yr] = [];
    byYear[yr].push(s);
  }
  const years = Object.keys(byYear).filter(y => +y >= y1 && +y <= y2).sort();
  if (!years.length) throw new Error(`ID ${id} に ${y1}〜${y2} の成績データがありません`);

  const basic = {};
  for (const yr of years) {
    const rows = byYear[yr];
    const row  = rows.find(r => !r.team) || rows[0];
    let teamStr;
    if (rows.length > 1) {
      const named   = rows.filter(r => r.team);
      const primary = named.reduce((a, b) => (a.stat.gamesPlayed >= b.stat.gamesPlayed ? a : b));
      teamStr = normalizeTeamAbbr(primary.team?.abbreviation || primary.team?.name || '???') + named.length;
    } else {
      teamStr = normalizeTeamAbbr(row.team?.abbreviation || row.team?.name || '???');
    }
    const st = row.stat;
    basic[yr] = {
      team: teamStr, g: st.gamesPlayed, pa: st.atBats,
      r: st.runs, h: st.hits, d: st.doubles, t: st.triples, hr: st.homeRuns,
      rbi: st.rbi, bb: st.baseOnBalls, so: st.strikeOuts,
      sb: st.stolenBases, cs: st.caughtStealing,
      avg: st.avg, obp: st.obp, ops: st.ops,
    };
  }

  const career = await mlbGet(
    `https://statsapi.mlb.com/api/v1/people/${id}/stats?stats=career&group=hitting&sportId=1`
  );
  const cs = career.stats[0]?.splits[0]?.stat || {};
  basic['通算'] = {
    team: basic[years[years.length - 1]]?.team?.replace(/\d+$/, '') || '---',
    g: cs.gamesPlayed, pa: cs.atBats,
    r: cs.runs, h: cs.hits, d: cs.doubles, t: cs.triples, hr: cs.homeRuns,
    rbi: cs.rbi, bb: cs.baseOnBalls, so: cs.strikeOuts,
    sb: cs.stolenBases, cs: cs.caughtStealing,
    avg: cs.avg, obp: cs.obp, ops: cs.ops,
  };

  const splitsRaw = {};
  await Promise.all(years.map(async yr => {
    const [vl, rp] = await Promise.all([
      mlbGet(`https://statsapi.mlb.com/api/v1/people/${id}/stats?stats=statSplits&group=hitting&sportId=1&sitCodes=vl&season=${yr}`),
      mlbGet(`https://statsapi.mlb.com/api/v1/people/${id}/stats?stats=statSplits&group=hitting&sportId=1&sitCodes=risp&season=${yr}`),
    ]);
    splitsRaw[yr] = {
      vsLAB:  vl.stats[0]?.splits[0]?.stat?.atBats || 0,
      vsLH:   vl.stats[0]?.splits[0]?.stat?.hits   || 0,
      rispAB: rp.stats[0]?.splits[0]?.stat?.atBats || 0,
      rispH:  rp.stats[0]?.splits[0]?.stat?.hits   || 0,
    };
  }));

  // キャッチャー守備成績（CS/SB率算出用） & 全ポジション守備（2002年以前フォールバック用）
  const catcherFielding  = { byYear: {}, career: null };
  const mlbApiFielding   = { byYear: {} };  // 全ポジション: FanGraphs/BB-Ref 欠損時の補完用
  try {
    const [fldYby, fldCar] = await Promise.all([
      mlbGet(`https://statsapi.mlb.com/api/v1/people/${id}/stats?stats=yearByYear&group=fielding&sportId=1`),
      mlbGet(`https://statsapi.mlb.com/api/v1/people/${id}/stats?stats=career&group=fielding&sportId=1`),
    ]);
    for (const s of (fldYby.stats?.[0]?.splits || [])) {
      const pos = s.position?.abbreviation;
      if (!pos || !s.season || s.sport?.id !== 1) continue;
      if (pos === 'C') {
        catcherFielding.byYear[s.season] = {
          sb: s.stat?.stolenBases    ?? 0,
          cs: s.stat?.caughtStealing ?? 0,
          pb: s.stat?.passedBall     ?? 0,
          g:  s.stat?.gamesPlayed    ?? 0,
        };
      }
      // 全ポジション守備データ: innings / games / CS(捕手専用) を保存
      if (!mlbApiFielding.byYear[s.season]) mlbApiFielding.byYear[s.season] = {};
      mlbApiFielding.byYear[s.season][pos] = {
        g:   s.stat?.gamesPlayed        ?? 0,
        inn: s.stat?.innings            ?? null, // "XXX.Y" 形式 or null
        rf9: s.stat?.rangeFactorPer9Inn ?? null, // API直値（なければ po+a から算出）
        po:  s.stat?.putOuts            ?? null, // 刺殺
        a:   s.stat?.assists            ?? null, // 補殺
        ch:  s.stat?.chances            ?? null, // 守備機会
        fld: s.stat?.fielding           ?? null, // 守備率 string ".987"
        cs:  s.stat?.caughtStealing     ?? 0,
        sb:  s.stat?.stolenBases        ?? 0,
        pb:  s.stat?.passedBall         ?? 0,
      };
    }
    const carCat = (fldCar.stats?.[0]?.splits || []).find(s => s.position?.abbreviation === 'C');
    if (carCat) catcherFielding.career = {
      sb: carCat.stat?.stolenBases    ?? 0,
      cs: carCat.stat?.caughtStealing ?? 0,
      pb: carCat.stat?.passedBall     ?? 0,
      g:  carCat.stat?.gamesPlayed    ?? 0,
    };
  } catch {}

  return { years, basic, splitsRaw, catcherFielding, mlbApiFielding };
}

// ── Browser data: Baseball Savant + FanGraphs ─────────────────────────────────
const PITCH_MAP = {
  '4-Seam Fastball': 'ff', 'Sinker': 'si', 'Two-Seam Fastball': 'si',
  'Changeup': 'ch', 'Slider': 'sl', 'Sweeper': 'st',
  'Curveball': 'cu', 'Knuckle Curve': 'cu', 'Cutter': 'fc',
  'Split-Finger': 'fs', 'Splitter': 'fs',
};
const emptyPitch = () => ({
  ff:{ba:'--',pa:0}, si:{ba:'--',pa:0}, ch:{ba:'--',pa:0},
  sl:{ba:'--',pa:0}, st:{ba:'--',pa:0}, cu:{ba:'--',pa:0},
  fc:{ba:'--',pa:0}, fs:{ba:'--',pa:0},
});

async function batFetchBrowserData(slug, id, playerFullName, years, onProgress, splitsRaw = {}) {
  const chromePath = findChrome();
  if (!chromePath) throw new Error('Chromeが見つかりません。Google ChromeまたはMicrosoft Edgeをインストールしてください。');

  const tmpDir = path.join(os.tmpdir(), 'mlb_stats_' + Date.now());
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

    const y1 = years[0], y2 = years[years.length - 1];

    // ── Baseball Savant ──────────────────────────────────────────────────────
    const sprintSpeed = {};
    const rawPitch = {};
    for (const yr of years) rawPitch[yr] = emptyPitch();

    try {
      onProgress('Baseball Savant を読み込み中...');
      const savantUrl = `https://baseballsavant.mlb.com/savant-player/${slug}-${id}` +
        `?stats=statcast&player_type=batter&startSeason=${y1}&endSeason=${y2}`;
      await page.goto(savantUrl, { waitUntil: 'networkidle2', timeout: 60000 });

      const savantRaw = await page.evaluate(() => {
        try {
          const statcast = window.serverVals?.statcast;
          const sprintData = (Array.isArray(statcast) ? statcast : [])
            .filter(e => e.year)
            .map(e => ({ year: String(e.year), pct: e.percent_speed_order }));
          const tables = document.querySelectorAll('table');
          let pitchTable = null;
          for (const t of tables) {
            if (t.innerText.includes('4-Seam') || t.innerText.includes('Sinker')) { pitchTable = t; break; }
          }
          if (!pitchTable) return { sprintData, pitchData: {} };
          const headers = [...pitchTable.querySelectorAll('thead th,thead td')].map(h => h.innerText.trim());
          const paIdx = headers.indexOf('PA'), baIdx = headers.indexOf('BA');
          if (paIdx < 0 || baIdx < 0) return { sprintData, pitchData: {} };
          const pitchData = {};
          for (const row of pitchTable.querySelectorAll('tbody tr')) {
            const cells = [...row.querySelectorAll('td')];
            if (cells.length < 2) continue;
            const yr = cells[0].innerText.trim(), pt = cells[1].innerText.trim();
            const ba = cells[baIdx]?.innerText.trim() || '--';
            const pa = parseInt(cells[paIdx]?.innerText.trim()) || 0;
            if (!pitchData[yr]) pitchData[yr] = {};
            pitchData[yr][pt] = { ba, pa };
          }
          return { sprintData, pitchData };
        } catch (e) {
          return { sprintData: [], pitchData: {}, error: e.message };
        }
      });

      for (const { year, pct } of (savantRaw.sprintData || [])) sprintSpeed[year] = pct;
      for (const yr of years) {
        for (const [ptName, { ba, pa }] of Object.entries(savantRaw.pitchData?.[yr] || {})) {
          const key = PITCH_MAP[ptName];
          if (!key) continue;
          const baNum = parseFloat(ba);
          if (!isNaN(baNum) && baNum >= 1.0) continue;
          rawPitch[yr][key] = { ba: (ba === '--' || ba === '') ? '--' : ba.replace('.', '__D__'), pa };
        }
      }
      if (savantRaw.error) onProgress('⚠ Baseball Savant 一部取得失敗: ' + savantRaw.error);
    } catch (e) {
      onProgress('⚠ Baseball Savant 取得失敗（空データで続行）: ' + e.message);
    }

    // ── Baseball Savant キャッチャーフレーミング（Puppeteerナビゲーション方式）──
    // フレーミングページはJSレンダリングのため fetch()では空テーブルになる。
    // Puppeteerでナビゲートし waitForSelector でJS描画完了を待って取得する。
    // パフォーマンス優先のためキャリアページ(year=0)のみ取得し、
    // 年別行には career フォールバック（processFile内）で対応する。
    const catcherFraming = { byYear: {}, career: null };
    try {
      onProgress('キャッチャーフレーミング データを取得中...');

      // テーブルから対象選手の行を抽出する共通関数
      // pid(数値ID)とlastName(姓)の両方でマッチングを試みる
      const extractFramingRow = async (pid, lastName) => page.evaluate((pid, lastName) => {
        // ── CSV フォールバック（同一オリジンfetch）────────────────────────
        async function tryCSV(yearParam) {
          try {
            const r = await fetch(
              `/catcher_framing?year=${yearParam}&team=&min=0&type=catcher&sort=4,1&csv=true`
            );
            if (!r.ok) return null;
            const text = await r.text();
            const lines = text.trim().split('\n');
            if (lines.length < 2) return null;
            const cols = lines[0].replace(/\r/g, '').split(',').map(c => c.replace(/"/g, '').trim().toLowerCase());
            const idIdx     = cols.findIndex(c => c === 'player_id' || c === 'pitcher_id' || c === 'id');
            const pitchIdx  = cols.findIndex(c => c === 'pitches' || c === 'n_called_pitches' || c.includes('pitch'));
            const runsIdx   = cols.findIndex(c => c === 'runs_extra' || c.includes('framing') || c.includes('run'));
            const nameIdx   = cols.findIndex(c => c === 'last_name' || c === 'name' || c === 'player_name' || c === 'last_name, first_name');
            if (pitchIdx < 0 || runsIdx < 0) return { csvErr: 'no_cols', cols: cols.slice(0, 10) };
            for (const raw of lines.slice(1)) {
              const cells = raw.replace(/\r/g, '').split(',').map(c => c.replace(/"/g, '').trim());
              const idMatch   = idIdx   >= 0 && String(cells[idIdx]) === String(pid);
              const nameMatch = nameIdx >= 0 && cells[nameIdx]?.toLowerCase().includes(lastName);
              if (!idMatch && !nameMatch) continue;
              const pitches = parseInt(cells[pitchIdx].replace(/,/g, '')) || 0;
              const runs    = parseFloat(cells[runsIdx]) || 0;
              if (!pitches) continue;
              return { pitches, runs };
            }
            return { csvErr: 'player_not_found' };
          } catch(e) { return { csvErr: e.message }; }
        }

        // ── HTML テーブルパース ──────────────────────────────────────────
        const table = document.querySelector('table');
        if (!table) return { err: 'no_table' };
        const rawHeaders = [...table.querySelectorAll('thead th,thead td')]
          .map(h => h.textContent.trim());
        const headers = rawHeaders.map(h => h.toLowerCase());
        const pitchIdx = headers.findIndex(h => h === 'pitches' || h === 'pitch' || h.includes('pitch'));
        const runsIdx  = headers.findIndex(h =>
          h.includes('framing run') || h.includes('run value') ||
          h.includes('runs_extra')  || h === 'framing');
        if (pitchIdx < 0 || runsIdx < 0)
          return { err: 'header_not_found', headers: rawHeaders.slice(0, 10) };

        for (const row of table.querySelectorAll('tbody tr')) {
          const cells = [...row.querySelectorAll('td')];
          // ID と名前の両方でマッチング
          const links = [...row.querySelectorAll('a')];
          const idMatch   = links.some(a => {
            const href = a.href || a.getAttribute('href') || '';
            return href.includes('-' + pid) || href.includes('/' + pid) || href.includes('=' + pid);
          });
          const nameMatch = row.textContent.toLowerCase().includes(lastName);
          if (!idMatch && !nameMatch) continue;
          const pitches = parseInt((cells[pitchIdx]?.textContent.trim() || '0').replace(/,/g, '')) || 0;
          const runs    = parseFloat(cells[runsIdx]?.textContent.trim()) || 0;
          if (!pitches) return null;
          return { pitches, runs };
        }
        // 行が見つからなかった場合の診断情報
        const rowCount = table.querySelectorAll('tbody tr').length;
        return { err: 'player_not_found', rowCount,
          sampleHrefs: [...(table.querySelector('tbody tr')?.querySelectorAll('a') || [])]
            .map(a => a.href || a.getAttribute('href')).slice(0, 3),
        };
      }, pid, lastName);

      const lastName = playerFullName.toLowerCase().split(' ').pop(); // 姓でマッチング

      // テーブルが実データ入りで描画されるまで待つ共通ヘルパー
      const waitForTableData = async (timeout = 20000) => {
        await page.waitForFunction(
          () => {
            const rows = document.querySelectorAll('table tbody tr');
            if (rows.length === 0) return false;
            // 数字を含むセルが存在すれば描画完了と判断
            return [...rows[0].querySelectorAll('td')].some(td => /\d/.test(td.textContent));
          },
          { timeout }
        ).catch(() => {});
      };

      // ── キャリア通算ページ（year=0）取得 ──────────────────────────────────
      await page.goto(
        'https://baseballsavant.mlb.com/catcher_framing?year=0&team=&min=0&type=catcher&sort=4,1',
        { waitUntil: 'domcontentloaded', timeout: 30000 }
      );
      await waitForTableData(20000);

      let careerResult = await extractFramingRow(id, lastName);

      // テーブルで見つからない場合は CSV フォールバック
      if (!careerResult?.pitches) {
        const csvRes = await page.evaluate(async (pid, ln) => {
          try {
            const r = await fetch('/catcher_framing?year=0&team=&min=0&type=catcher&sort=4,1&csv=true');
            if (!r.ok) return { csvErr: r.status };
            const text = await r.text();
            const lines = text.trim().split('\n');
            if (lines.length < 2) return { csvErr: 'empty' };
            const cols = lines[0].replace(/\r/g,'').split(',').map(c=>c.replace(/"/g,'').trim().toLowerCase());
            const pidIdx   = cols.findIndex(c => c === 'player_id' || c === 'pitcher_id');
            const pchIdx   = cols.findIndex(c => c === 'pitches' || c === 'n_called_pitches' || c.includes('pitch'));
            const runIdx   = cols.findIndex(c => c === 'runs_extra' || c.includes('framing') || c.includes('run'));
            if (pchIdx < 0 || runIdx < 0) return { csvErr: 'no_cols', cols: cols.slice(0,10) };
            for (const raw of lines.slice(1)) {
              const c = raw.replace(/\r/g,'').split(',').map(x=>x.replace(/"/g,'').trim());
              if (pidIdx >= 0 && String(c[pidIdx]) !== String(pid)) {
                if (!c.join(' ').toLowerCase().includes(ln)) continue;
              }
              const pitches = parseInt(c[pchIdx].replace(/,/g,'')) || 0;
              const runs    = parseFloat(c[runIdx]) || 0;
              if (!pitches) continue;
              return { pitches, runs };
            }
            return { csvErr: 'not_found' };
          } catch(e) { return { csvErr: e.message }; }
        }, id, lastName);
        if (csvRes?.pitches) careerResult = csvRes;
        else onProgress(`⚠ フレーミングCSV: ${JSON.stringify(csvRes).slice(0, 120)}`);
      }

      if (careerResult?.pitches) {
        catcherFraming.career = { pitches: careerResult.pitches, runs: careerResult.runs };
        const lead = Math.round(careerResult.runs * 1500 / careerResult.pitches);
        onProgress(`キャッチャーフレーミング(通算) 取得: pitches=${careerResult.pitches} runs=${careerResult.runs} → リード≈${lead}`);
      } else {
        // 詳細診断情報を出力
        const info = careerResult?.err || careerResult?.err;
        onProgress(`キャッチャーフレーミング: データなし [${info || '未検出'}] rowCount=${careerResult?.rowCount}`);
      }

      // ── 年別ページ取得（捕手と確認できた場合のみ、失敗は career で代替）──
      if (catcherFraming.career) {
        for (const yr of years) {
          try {
            await page.goto(
              `https://baseballsavant.mlb.com/catcher_framing?year=${yr}&team=&min=0&type=catcher&sort=4,1`,
              { waitUntil: 'domcontentloaded', timeout: 20000 }
            );
            await waitForTableData(12000);
            const yrResult = await extractFramingRow(id, lastName);
            if (yrResult?.pitches) catcherFraming.byYear[yr] = yrResult;
          } catch {}
        }
        const yrCount = Object.values(catcherFraming.byYear).filter(Boolean).length;
        if (yrCount > 0) onProgress(`フレーミング年別取得: ${yrCount}年分`);
      }
    } catch (e) {
      onProgress('⚠ キャッチャーフレーミング 取得失敗: ' + e.message);
    }

    // ── FanGraphs ────────────────────────────────────────────────────────────
    const fieldingByYear = {};
    for (const yr of years) fieldingByYear[yr] = {};

    try {
      onProgress('FanGraphs を読み込み中...');
      try {
        // domcontentloaded で十分（fetch APIはDOMロード後に使用可能）
        await page.goto('https://www.fangraphs.com/', { waitUntil: 'domcontentloaded', timeout: 60000 });
      } catch (e) {
        const title = await page.title().catch(() => '');
        if (!title.toLowerCase().includes('fangraphs')) throw new Error('FanGraphs 読み込み失敗: ' + e.message);
      }

      // FanGraphs DRS は 2003年以降のみ有効（2002年以前は BB-Ref 仮想DRSにフォールバック）
      const fgYears = years.filter(yr => parseInt(yr) >= 2003);
      onProgress(`FanGraphs から守備データを取得中... (対象: ${fgYears.length}年, 2002年以前はBB-Ref仮想DRS)`);

      if (fgYears.length > 0) {
        const fieldingRaw = await page.evaluate(async (yearsArr, pName) => {
          // 全年並列取得（高速化）
          const entries = await Promise.all(yearsArr.map(async yr => {
            try {
              const r = await fetch(
                `/api/leaders/major-league/data?age=0&pos=all&stats=fld&lg=all&qual=0` +
                `&season=${yr}&season1=${yr}&startdate=&enddate=&month=0&hand=&team=0` +
                `&pageitems=2000&pagenum=1&ind=0&rost=0&players=0&type=1`
              );
              const d = await r.json();
              const rows = Array.isArray(d.data) ? d.data : [];
              return { yr, data: rows.filter(row => row.PlayerName === pName)
                .map(row => ({ pos: row.Pos, inn: row.Inn, drs: row.DRS })) };
            } catch { return { yr, data: [] }; }
          }));
          const result = {};
          for (const { yr, data } of entries) result[yr] = data;
          return result;
        }, fgYears, playerFullName);

        for (const yr of fgYears) {
          const entries = Array.isArray(fieldingRaw[yr]) ? fieldingRaw[yr] : [];
          for (const { pos, inn, drs } of entries) {
            if (!pos || pos === 'undefined') continue;  // 無効ポジションを除外
            const drsNum = Number(drs);
            fieldingByYear[yr][pos] = { inn: String(inn || 0), drs: isNaN(drsNum) ? 0 : drsNum };
          }
        }
      }
    } catch (e) {
      onProgress('⚠ FanGraphs 取得失敗（空データで続行）: ' + e.message);
    }

    // ── Baseball Reference フォールバック（歴代選手・DRS欠損/スプリット欠損時）────
    // 発動条件: FanGraphs で守備データが取れない年あり、または MLB API スプリット全欠損
    const bbRefSplits = {};
    let bbRefOfGames = {};  // BB-Ref から取得した外野G数（OF→RF/LF/CF 按分用）
    let battingHand = null;  // BB-Ref から取得: 'L'=左打 / 'R'=右打 / 'S'=両打
    const missingFieldingYears = years.filter(yr => Object.keys(fieldingByYear[yr]).length === 0);
    const allSplitsEmpty = years.every(yr => !(splitsRaw[yr]?.vsLAB) && !(splitsRaw[yr]?.rispAB));
    if (missingFieldingYears.length > 0 || allSplitsEmpty) {
      try {
        onProgress('Baseball Reference からデータを取得中...');

        // ── Step 1: MLB Stats API xrefIds から BB-Ref ID 取得（最も確実）─────────
        let bbSlug = null;
        try {
          const xrefData = await mlbGet(
            `https://statsapi.mlb.com/api/v1/people/${id}?hydrate=xrefIds`
          );
          const xrefs = xrefData?.people?.[0]?.xrefIds ?? [];
          const brefEntry = xrefs.find(x => {
            const t = String(x.xrefIdType ?? '').toLowerCase();
            return t.includes('bref') || t.includes('bbref') || t === 'br';
          });
          if (brefEntry?.xrefId) {
            bbSlug = String(brefEntry.xrefId).trim();
            onProgress(`BB-Ref ID (MLB Stats API): ${bbSlug}`);
          }
        } catch {}

        // ── Step 2: BB-Ref 検索ページで名前検索 ─────────────────────────────────
        // ※ページ全体のリンクを正規表現で拾うと関係ない選手(サイドバー等)を
        //   誤取得するため、アンカーテキストで選手名と一致するリンクのみ採用する
        if (!bbSlug) {
          try {
            const searchUrl = `https://www.baseball-reference.com/search/search.fcgi?search=${encodeURIComponent(playerFullName)}`;
            await page.goto(searchUrl, { waitUntil: 'domcontentloaded', timeout: 20000 });

            // 単一結果の場合 BB-Ref は選手ページへ直接リダイレクトする
            const finalUrl = page.url();
            const urlMatch = finalUrl.match(/\/players\/[a-z]\/([a-z0-9]+)\.shtml/);
            if (urlMatch) {
              // リダイレクト先が正しい選手か h1 で確認
              const isMatch = await page.evaluate((name) => {
                const h1 = document.querySelector('#info h1') || document.querySelector('h1');
                if (!h1) return false;
                const t = h1.textContent.trim().toLowerCase();
                return name.toLowerCase().split(' ').filter(p => p.length > 1).every(p => t.includes(p));
              }, playerFullName);
              if (isMatch) {
                bbSlug = urlMatch[1];
                onProgress(`BB-Ref ID (検索リダイレクト): ${bbSlug}`);
              }
            }

            // 複数候補リストの場合: アンカーテキストが選手名を含むリンクのみ採用
            if (!bbSlug) {
              bbSlug = await page.evaluate((name) => {
                const parts = name.toLowerCase().split(' ').filter(p => p.length > 1);
                for (const a of document.querySelectorAll('a[href*="/players/"]')) {
                  const href = a.getAttribute('href') || '';
                  const m    = href.match(/\/players\/[a-z]\/([a-z0-9]+)\.shtml/);
                  if (!m) continue;
                  const txt = a.textContent.trim().toLowerCase();
                  if (parts.every(p => txt.includes(p))) return m[1];
                }
                return null;
              }, playerFullName);
              if (bbSlug) onProgress(`BB-Ref ID (検索リスト): ${bbSlug}`);
            }
          } catch {}
        }

        // ── Step 3: 姓5+名2+連番 の命名規則でスラッグ候補を検証（01〜05）────────
        if (!bbSlug) {
          const SUFFIXES = new Set(['jr.', 'sr.', 'ii', 'iii', 'iv', 'v', 'jr', 'sr']);
          const cleanName = (playerFullName || '').normalize('NFD')
            .replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z ]/g, '').trim();
          const nameParts = cleanName.split(/\s+/).filter(p => !SUFFIXES.has(p));
          const firstName = nameParts[0] || '';
          const lastName  = nameParts[nameParts.length - 1] || '';
          const prefix    = lastName.slice(0, 5) + firstName.slice(0, 2);
          for (let n = 1; n <= 5 && !bbSlug; n++) {
            const cand = prefix + String(n).padStart(2, '0');
            try {
              await page.goto(
                `https://www.baseball-reference.com/players/${cand[0]}/${cand}.shtml`,
                { waitUntil: 'domcontentloaded', timeout: 20000 }
              );
              const isMatch = await page.evaluate((name) => {
                const h1 = document.querySelector('#info h1') || document.querySelector('h1');
                if (!h1) return false;
                const t = h1.textContent.trim().toLowerCase();
                return name.toLowerCase().split(' ').filter(p => p.length > 1).every(p => t.includes(p));
              }, playerFullName);
              if (isMatch) { bbSlug = cand; onProgress(`BB-Ref ID (命名規則): ${bbSlug}`); }
            } catch {}
          }
        }

        if (bbSlug) {
          // ── 選手ページ取得（Step3でナビゲート済みでなければ再取得）────────────
          const currentUrl = page.url();
          if (!currentUrl.includes(bbSlug)) {
            // BB-Ref は analytics 等で通信が続くため networkidle2 は使わず
            // domcontentloaded で DOM 完成後に waitForSelector でテーブルを待つ
            await page.goto(
              `https://www.baseball-reference.com/players/${bbSlug[0]}/${bbSlug}.shtml`,
              { waitUntil: 'domcontentloaded', timeout: 30000 }
            );
          }

          // ── 打席情報（LHB/RHB/Switch）を取得 → 対左BA推計に使用 ────────────
          battingHand = await page.evaluate(() => {
            const info = document.querySelector('#info') || document.querySelector('#necro-bio') || document.body;
            const txt = info?.textContent || '';
            if (/bats:\s*left/i.test(txt))  return 'L';
            if (/bats:\s*right/i.test(txt)) return 'R';
            if (/bats:\s*(both|switch)/i.test(txt)) return 'S';
            return null;
          }).catch(() => null);
          onProgress(`[診断] 打席=${battingHand || '不明'}`);

          // ── 診断ログ（スキップ原因を特定）────────────────────────────────────
          onProgress(`[診断] bbSlug=${bbSlug}, missingFieldingYears=${missingFieldingYears.length}/${years.length}, allSplitsEmpty=${allSplitsEmpty}`);
          // FanGraphs で取得できた年の守備データ状況を表示（0の年が真に欠損）
          const fgStatus = years.map(yr => `${yr}:${Object.keys(fieldingByYear[yr]||{}).join('|')||'空'}`).join(', ');
          onProgress(`[診断] FanGraphs守備: ${fgStatus}`);

          // ── 守備データ（standard_fielding テーブル）→ DRS推定 ─────────────────
          // ▶ page.content() で Node.js 側に生 HTML を取得し、コメント除去後に
          //   DOMParser 経由でテーブルを解析する
          // ▶ BB-Ref の守備テーブル ID:
          //     新形式: #players_standard_fielding
          //     旧形式: #standard_fielding / #fielding_standard
          //     コメント内に隠れている場合も存在する
          if (missingFieldingYears.length > 0 || allSplitsEmpty) {
            try {
              // 守備テーブルが JS 遅延ロードされる場合があるため最大5秒待つ
              await page.waitForSelector(
                '#players_standard_fielding, #standard_fielding, #fielding_standard',
                { timeout: 5000 }
              ).catch(() => {}); // タイムアウトしても続行

              const rawHtml = await page.content();
              const fieldingIdPresent = rawHtml.includes('id="players_standard_fielding"') ||
                                        rawHtml.includes('id="standard_fielding"') ||
                                        rawHtml.includes('id="fielding_standard"');
              onProgress(`BB-Ref HTML取得: ${rawHtml.length.toLocaleString()} chars, 守備テーブル存在=${fieldingIdPresent}`);

              // コメント内テーブルを露出させる（<!-- --> を除去）
              const cleanHtml = rawHtml.replace(/<!--([\s\S]*?)-->/g, '$1');
              const fieldingIdClean = cleanHtml.includes('id="players_standard_fielding"') ||
                                      cleanHtml.includes('id="standard_fielding"') ||
                                      cleanHtml.includes('id="fielding_standard"');
              onProgress(`コメント除去後: ${cleanHtml.length.toLocaleString()} chars, 守備テーブル存在=${fieldingIdClean}`);

              // 解析済み HTML を Puppeteer の evaluate へ渡して DOMParser で処理
              const bbResult = await page.evaluate((html) => {
                try {
                  const doc2 = new DOMParser().parseFromString(html, 'text/html');
                  const tableIds = [...doc2.querySelectorAll('table[id]')]
                    .map(t => t.id).filter(Boolean);
                  // BB-Ref 守備テーブルの ID: 新形式・旧形式・コメント除去後のいずれかに対応
                  const table = doc2.querySelector('#players_standard_fielding') ||
                                doc2.querySelector('#standard_fielding') ||
                                doc2.querySelector('#fielding_standard') ||
                                // フォールバック: RF/9 列を持つテーブルを探す
                                [...doc2.querySelectorAll('table')].find(t => {
                                  const h = t.querySelector('[data-stat="range_factor_9inn"], [data-stat="rf9"]');
                                  return !!h;
                                });
                  if (!table) return { err: 'no_table', tableIds };

                  // ヘッダーの data-stat（全件・デバッグ用）
                  const headerStats = [...new Set(
                    [...table.querySelectorAll('[data-stat]')]
                      .map(el => el.getAttribute('data-stat')).filter(Boolean)
                  )];

                  // サンプル行のデータを確認（最初にデータが入る行を1件）
                  let sampleRow = null;
                  const result = {};
                  // ofGames[yr] = { RF: N, LF: N, CF: N }  ← G比率按分用
                  const ofGames = {};
                  // ── 年キャリーフォワード ──────────────────────────────────────────
                  // BB-Ref の続き行（同年の2ポジション目以降）には year_id が空白になる
                  // currentYr を引き継ぐことで RF/LF/CF 個別行も正しく取得する
                  let currentYr = '';
                  for (const row of table.querySelectorAll('tbody tr, tr')) {
                    if (row.classList.contains('thead') || row.classList.contains('minors_table')) continue;
                    // tfoot 内（通算行）はスキップ
                    if (row.closest && row.closest('tfoot')) continue;
                    // 年取得: 旧形式(year_ID) と新形式(year_id) 両対応
                    const yearEl = row.querySelector('[data-stat="year_ID"]') ||
                                   row.querySelector('[data-stat="year_id"]');
                    const rawYrText  = yearEl ? yearEl.textContent.trim() : '';
                    const yrDigits   = rawYrText.replace(/\D/g, '');
                    // ── 通算行の検知: year セルに "Career" 等の英字 → currentYr リセット後スキップ
                    if (rawYrText && /[a-zA-Z]/.test(rawYrText)) { currentYr = ''; continue; }
                    if (yrDigits.length === 4) currentYr = yrDigits;
                    // 有効な年がなければスキップ（テーブル冒頭の無効行等）
                    if (!currentYr) continue;
                    const yr = currentYr;
                    // 最初の有効行のデータを全 stat と値で記録
                    if (!sampleRow) {
                      sampleRow = {};
                      for (const el of row.querySelectorAll('[data-stat]')) {
                        const s = el.getAttribute('data-stat');
                        const v = el.textContent.trim();
                        if (s && v) sampleRow[s] = v;
                      }
                    }
                    // ── 複数名フォールバック付きゲッター ────────────────────────────
                    // BB-Ref 新形式は f_ プレフィックス、旧形式はプレフィックスなし
                    const get = (...names) => {
                      for (const n of names) {
                        const el = row.querySelector(`[data-stat="${n}"]`);
                        const v = el ? el.textContent.trim() : '';
                        if (v && v !== '--' && v !== '.') return v;
                      }
                      return '';
                    };
                    const pos = get('f_position', 'pos', 'position');
                    if (!pos || pos === 'Pos' || pos === 'Position') continue;
                    const gVal  = parseInt(get('f_games', 'g', 'G') || '0') || 0;
                    const innRaw = get('f_innings', 'f_inn', 'f_inn_outs', 'Inn', 'inn_outs', 'inn');
                    // ── 通算行ヒューリスティック: イニング > 2000 は1シーズン最大値を超えるためスキップ
                    const innNum = parseFloat(innRaw || '0') || 0;
                    if (innNum > 2000) continue;
                    result[yr] = result[yr] || {};
                    result[yr][pos] = {
                      inn:   innRaw,
                      g:     gVal,
                      ch:    get('f_chances', 'f_tc', 'chances', 'tc')                        || '0',
                      e:     get('f_errors',  'f_e',  'e')                                    || '0',
                      fld:   get('f_fielding_perc', 'f_pct', 'fielding_perc')                 || '0',
                      lgFld: get('f_fielding_perc_lg', 'f_lg_fielding_perc', 'lg_fielding_perc') || '0',
                      rf9:   get('f_range_factor_per_nine', 'f_rf9', 'range_factor_9inn')     || '0',
                      lgRf9: get('f_range_factor_per_nine_lg', 'f_lg_rf9', 'lg_range_factor_9inn') || '0',
                      // 捕手専用（他ポジションでは空欄になる）
                      cs:    get('f_caught_stealing', 'caught_stealing', 'cs_catcher', 'rcs')  || '0',
                      sb:    get('f_stolen_bases',    'stolen_bases',    'sb_against')         || '0',
                      pb:    get('f_passed_ball',     'passed_ball',     'pb')                 || '0',
                    };
                    // G 数を外野ポジション別に記録（OF→RF/LF/CF 按分用）
                    if (['RF','LF','CF'].includes(pos) && gVal > 0) {
                      ofGames[yr] = ofGames[yr] || {};
                      ofGames[yr][pos] = gVal;
                    }
                  }
                  return { result, ofGames, headerStats, tableIds, sampleRow };
                } catch (e) { return { err: e.message }; }
              }, cleanHtml);

              if (bbResult?.err) {
                onProgress(`⚠ BB-Ref 守備テーブル未検出 (${bbResult.err})`);
                onProgress(`  検出テーブルIDs: [${(bbResult.tableIds||[]).slice(0, 12).join(', ')}]`);
              } else if (bbResult?.result) {
                onProgress(`BB-Ref 守備ヘッダー(全): [${(bbResult.headerStats||[]).join(', ')}]`);
                if (bbResult.sampleRow) {
                  onProgress(`BB-Ref 守備サンプル行: ${JSON.stringify(bbResult.sampleRow)}`);
                }
                // FanGraphs データがない年を全て対象にする（missingFieldingYears が空でも同じ結果）
                const targetYears = years.filter(yr => Object.keys(fieldingByYear[yr]).length === 0);
                // targetYears が空で BB-Ref に結果があれば、全年を対象に（years が空の場合も対応）
                const effectiveTargetYears = targetYears.length > 0
                  ? targetYears
                  : Object.keys(bbResult.result).filter(yr => yr.length === 4);
                onProgress(`BB-Ref 守備対象年: [${effectiveTargetYears.join(', ')}] / 取得年: [${Object.keys(bbResult.result).join(', ')}]`);
                for (const yr of effectiveTargetYears) {
                  const posMap = bbResult.result[yr];
                  if (!posMap) continue;
                  if (!fieldingByYear[yr]) fieldingByYear[yr] = {};  // years が空の場合も対応
                  for (const [pos, d] of Object.entries(posMap)) {
                    const innStr  = String(d.inn || '0').replace(/[,\s]/g, '');
                    const innMatch = innStr.match(/^(\d+)(?:\.(\d))?$/);
                    const innFull  = innMatch ? parseInt(innMatch[1]) : 0;
                    const innFrac  = innMatch ? parseInt(innMatch[2] || '0') : 0;
                    const innDec   = innFull + innFrac / 3;
                    const innFmt   = innFull + '.' + innFrac;
                    // innDec < 1 のみスキップ（lgRf9 = 0 でも書き込む）
                    if (innDec < 1) continue;
                    const ch    = parseInt(d.ch)    || 0;
                    const fld   = parseFloat(d.fld)   || 0;
                    const lgFld = parseFloat(d.lgFld) || 0;
                    const rf9   = parseFloat(d.rf9)   || 0;
                    const lgRf9 = parseFloat(d.lgRf9) || 0;
                    // ── 仮想 DRS 計算 ────────────────────────────────────────────
                    let drsVal;
                    if (pos === 'C') {
                      // ── 捕手専用: CS%ベース + RF/9補助 ────────────────────────
                      // 根拠: 捕手のRF/9はチーム三振率に依存するため個人守備力を測れない
                      //       → CS%（盗塁阻止率）が主要指標、RF/9は補助（係数0.05）
                      //
                      // [CS%ベース盗塁阻止DRS] FanGraphs実測との較正結果:
                      //   CS%差10%・120試み → stealDRS = +4.8
                      //   現役優良捕手(CS%40%,lg30%) 120試み → +4.8 ≈ FanGraphs典型値 +4〜+6 ✓
                      //   現役最優秀(CS%50%,lg30%) 130試み → +10.4 ≈ FanGraphs上位値 +8〜+12 ✓
                      const cs    = parseInt(d.cs)  || 0;
                      const sbA   = parseInt(d.sb)  || 0;
                      const total = cs + sbA;
                      // 時代別リーグ平均CS%
                      const yrNum    = parseInt(yr) || 1950;
                      const lgCsPct  = yrNum < 1970 ? 0.38
                                     : yrNum < 1985 ? 0.36
                                     : yrNum < 1995 ? 0.33
                                     : yrNum < 2005 ? 0.30
                                     : yrNum < 2015 ? 0.28
                                     :                0.26;
                      // 盗塁阻止DRS: CS%差 × 試み数 × 0.40
                      const stealDRS = total >= 10
                        ? (cs / total - lgCsPct) * total * 0.40
                        : 0;
                      // ブロッキングDRS補助: RF/9差（係数0.05でチーム依存を抑制）
                      const blockDRS = lgRf9 > 0
                        ? (rf9 - lgRf9) * innDec / 9 * 0.05
                        : 0;
                      const errorDRS = (ch > 0 && lgFld > 0) ? ch * (fld - lgFld) * 0.5 : 0;
                      drsVal = Math.round(stealDRS + blockDRS + errorDRS);
                    } else {
                      // ── 捕手以外: RF/9ベース ─────────────────────────────────
                      // 係数 0.25: RF/9差 0.3 × 1296inn/9 × 0.25 ≈ +11 (実測+10〜+15と一致)
                      const rangeDRS = lgRf9 > 0 ? (rf9 - lgRf9) * innDec / 9 * 0.25 : 0;
                      const errorDRS = (ch > 0 && lgFld > 0) ? ch * (fld - lgFld) * 0.5 : 0;
                      drsVal = Math.round(rangeDRS + errorDRS);
                    }
                    fieldingByYear[yr][pos] = { inn: innFmt, drs: drsVal, g: parseInt(d.g) || 0 };
                  }
                  const written = Object.entries(fieldingByYear[yr])
                    .filter(([,v]) => v && !isNaN(v.drs))
                    .map(([p,v]) => `${p}(inn=${v.inn},drs=${v.drs})`);
                  if (written.length > 0)
                    onProgress(`BB-Ref 守備推定 ${yr}: ${written.join(', ')}`);
                }
                // OF→RF/LF/CF 按分用のG数を上位スコープに保存
                if (bbResult.ofGames && Object.keys(bbResult.ofGames).length > 0) {
                  bbRefOfGames = bbResult.ofGames;
                  onProgress(`BB-Ref 外野G数: ${JSON.stringify(bbRefOfGames)}`);
                }
              }
            } catch (e) {
              onProgress(`⚠ BB-Ref 守備HTML取得失敗: ${e.message}`);
            }
          }

          // ── スプリット取得（通算 → 年別欠損の補完値として使用）─────────────
          try {
            const splitUrl = `https://www.baseball-reference.com/players/split.fcgi?id=${bbSlug}&year=all&t=b`;
            await page.goto(splitUrl, { waitUntil: 'domcontentloaded', timeout: 25000 });

            // スプリットページも HTML コメント内にテーブルが入っている場合があるため
            // Node.js 側で page.content() を取得してコメント除去後に DOMParser で解析
            const splitRawHtml = await page.content();
            const splitCleanHtml = splitRawHtml.replace(/<!--([\s\S]*?)-->/g, '$1');
            onProgress(`BB-Ref スプリットHTML: ${splitRawHtml.length.toLocaleString()} chars`);

            const bbSplit = await page.evaluate((html) => {
              try {
                const doc2 = new DOMParser().parseFromString(html, 'text/html');
                // 診断: スプリットページに存在するテーブルの split ラベルを収集
                const diagLabels = [];
                const diagTableIds = [];
                for (const table of doc2.querySelectorAll('table')) {
                  if (table.id) diagTableIds.push(table.id);
                  for (const row of table.querySelectorAll('tbody tr, tr')) {
                    const splitCell = row.querySelector('[data-stat="split"]') ||
                                      row.querySelector('th') ||
                                      row.querySelector('td');
                    const txt = (splitCell?.textContent || '').trim();
                    if (txt && diagLabels.length < 40) diagLabels.push(txt);
                  }
                }
                const parseSplit = (doc) => {
                  for (const table of doc.querySelectorAll('table')) {
                    let vsLRow = null, rispRow = null;
                    for (const row of table.querySelectorAll('tbody tr, tr')) {
                      // data-stat="split" の th か、最初の th/td のテキストで判定
                      const splitCell = row.querySelector('[data-stat="split"]') ||
                                        row.querySelector('th') ||
                                        row.querySelector('td');
                      const txt = (splitCell?.textContent || '').trim().toLowerCase();
                      if (!vsLRow  && (txt === 'vs. lhp' || txt === 'lhp' || txt === 'left' ||
                                       txt === 'vs lhp'  || txt.includes('lhp') || txt.includes('left-handed')))
                        vsLRow = row;
                      if (!rispRow && (txt === 'risp' || txt.includes('scoring position') ||
                                       txt === 'bases loaded' || txt.includes('risp')))
                        rispRow = row;
                    }
                    if (vsLRow || rispRow) {
                      const getN = (row, stat) => {
                        if (!row) return 0;
                        const c = row.querySelector(`[data-stat="${stat}"]`);
                        return parseInt(c?.textContent.trim() || '0') || 0;
                      };
                      const result = {
                        vsLAB:  getN(vsLRow,  'AB'),
                        vsLH:   getN(vsLRow,  'H'),
                        rispAB: getN(rispRow, 'AB'),
                        rispH:  getN(rispRow, 'H'),
                        _diagLabels:   diagLabels,
                        _diagTableIds: diagTableIds,
                      };
                      if (result.vsLAB > 0 || result.rispAB > 0) return result;
                    }
                  }
                  return { _diagLabels: diagLabels, _diagTableIds: diagTableIds };
                };
                // まずコメント除去済みHTMLを試し、なければ元のページを試す
                return parseSplit(doc2) || parseSplit(document);
              } catch (e) { return { _err: e.message }; }
            }, splitCleanHtml);

            // 診断ログ
            if (bbSplit?._diagTableIds?.length > 0)
              onProgress(`スプリットページ テーブルID: [${bbSplit._diagTableIds.slice(0,15).join(', ')}]`);
            if (bbSplit?._diagLabels?.length > 0)
              onProgress(`スプリットページ ラベルサンプル: [${bbSplit._diagLabels.slice(0,20).join(' | ')}]`);

            if (bbSplit && (bbSplit.vsLAB > 0 || bbSplit.rispAB > 0)) {
              for (const yr of years) bbRefSplits[yr] = bbSplit;
              onProgress(`BB-Ref スプリット(通算): 対左 ${bbSplit.vsLAB}AB / RISP ${bbSplit.rispAB}AB`);
            } else {
              onProgress('BB-Ref スプリット: データなし（診断ログを確認してください）');
            }
          } catch (e) {
            onProgress('⚠ BB-Ref スプリット取得失敗: ' + e.message);
          }

        } else {
          onProgress(`⚠ Baseball Reference: 選手ページ未発見 (${playerFullName})`);
        }
      } catch (e) {
        onProgress('⚠ Baseball Reference 取得失敗: ' + e.message);
      }
    }

    // ── MLB The Show Speed (Method 2: Baseball Savantデータがない年度のみ取得) ──────
    const mlbTheShowSpeed = {};
    const yearsWithoutSS = years.filter(yr => sprintSpeed[yr] == null);
    if (yearsWithoutSS.length > 0) {
      try {
        onProgress('MLB The Show データを取得中...');
        await page.goto('https://mlbtheshow.com/', { waitUntil: 'domcontentloaded', timeout: 30000 });
        for (const yr of yearsWithoutSS) {
          try {
            // 年度に対応するゲームエディション: 2023→/23/, 2024→/24/, 2025以降→パスなし
            const gyNum = parseInt(yr) - 2000;
            const apiBase = gyNum >= 25 ? '' : `/${gyNum}`;
            const apiUrl = `https://mlbtheshow.com${apiBase}/apis/items.json?type=mlb_card` +
              `&page=1&per_page=100&name=${encodeURIComponent(playerFullName)}&series=Live`;
            const data = await page.evaluate(async url => {
              try {
                const r = await fetch(url);
                if (!r.ok) return null;
                return await r.json();
              } catch { return null; }
            }, apiUrl);
            if (data?.items?.length > 0) {
              const ln = playerFullName.toLowerCase().split(' ').pop();
              const card = data.items.find(c => c.name?.toLowerCase().includes(ln)) || data.items[0];
              // SPDフィールド名は spd / speed / SPD のいずれか
              const spd = card?.spd ?? card?.speed ?? card?.SPD ?? null;
              if (spd != null) mlbTheShowSpeed[yr] = Number(spd);
            }
          } catch {}
        }
        const got = Object.keys(mlbTheShowSpeed);
        if (got.length > 0)
          onProgress('MLB The Show SPD 取得: ' + got.map(y => `${y}→${mlbTheShowSpeed[y]}`).join(', '));
        else if (yearsWithoutSS.length > 0)
          onProgress('MLB The Show データなし（盗塁ベース推計に切替）');
      } catch (e) {
        onProgress('⚠ MLB The Show 取得失敗（盗塁ベース推計で代替）: ' + e.message);
      }
    }

    // ── OF → 個別外野ポジションへのマッピング ─────────────────────────────────
    // 年キャリーフォワード修正後も個別行(RF/LF/CF)が存在する場合はそのまま使用。
    // OF 合算のみの年:
    //   ① bbRefOfGames[yr] に RF/LF/CF の G数がある → G 数比率で Inn・DRS を按分
    //   ② G数不明                                   → RF のみに全量割り当て（フォールバック）
    for (const yr of Object.keys(fieldingByYear)) {
      const fy = fieldingByYear[yr];
      if (!fy || !fy['OF']) continue;
      const hasIndividual = fy['LF'] || fy['CF'] || fy['RF'];
      if (hasIndividual) continue;  // 個別行がある年は何もしない

      const ofEntry = fy['OF'];
      const ofInn   = String(ofEntry.inn || '0');
      const ofDRS   = typeof ofEntry.drs === 'number' ? ofEntry.drs : 0;
      const [ofFull, ofFrac] = ofInn.split('.').map(Number);
      const ofOuts  = (ofFull || 0) * 3 + (ofFrac || 0);

      // G数比率で按分（bbRefOfGames は BB-Ref から収集した RF/LF/CF 別 G数）
      const gMap  = bbRefOfGames[yr] || {};
      const rfG   = gMap['RF'] || 0;
      const lfG   = gMap['LF'] || 0;
      const cfG   = gMap['CF'] || 0;
      const totalG = rfG + lfG + cfG;

      if (totalG > 0) {
        // G数比率でイニング・DRS を按分して各ポジションに書き込む
        const distribute = (pos, g) => {
          if (g <= 0) return;
          const ratio = g / totalG;
          const outs  = Math.round(ofOuts * ratio);
          fy[pos] = {
            inn: Math.floor(outs / 3) + '.' + (outs % 3),
            drs: Math.round(ofDRS * ratio),
            g:   g,
          };
        };
        if (rfG > 0) distribute('RF', rfG);
        if (lfG > 0) distribute('LF', lfG);
        if (cfG > 0) distribute('CF', cfG);
        onProgress(`OF按分 ${yr}: RF(${rfG}G) LF(${lfG}G) CF(${cfG}G) → total=${totalG}G, ofInn=${ofInn}, ofDRS=${ofDRS}`);
      } else {
        // G数不明 → RF フォールバック
        fy['RF'] = { inn: ofEntry.inn, drs: ofDRS, g: ofEntry.g || 0 };
      }
    }

    return { sprintSpeed, rawPitch, fieldingByYear, mlbTheShowSpeed, catcherFraming, bbRefSplits, battingHand };
  } finally {
    await browser.close();
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch {}
  }
}

// ── Excel build helpers ───────────────────────────────────────────────────────
function weightedBA(entries) {
  let sumH = 0, sumPA = 0;
  for (const e of entries) {
    if (!e || e.ba === '--' || !e.pa) continue;
    sumH  += parseFloat(e.ba.replace('__D__', '0.')) * e.pa;
    sumPA += e.pa;
  }
  return sumPA === 0 ? '--' : (sumH / sumPA).toFixed(3).split('.')[1];
}
function addInningsList(list) {
  const total = list.filter(Boolean).reduce((acc, s) => {
    const [f, r] = String(s).split('.');
    return acc + parseInt(f) * 3 + parseInt(r || 0);
  }, 0);
  return Math.floor(total / 3) + '.' + (total % 3);
}
function innToOuts(s) {
  const [f, r] = String(s).split('.');
  return parseInt(f) * 3 + parseInt(r || 0);
}

async function batBuildExcel(playerName, years, basic, splitsRaw, sprintSpeed, mlbTheShowSpeed, rawPitch, fieldingByYear) {
  const splits = {};
  for (const yr of years) {
    const d = splitsRaw[yr];
    splits[yr] = {
      vsLeft: d.vsLAB  === 0 ? '--' : (d.vsLH  / d.vsLAB ).toFixed(3).split('.')[1],
      risp:   d.rispAB === 0 ? '--' : (d.rispH / d.rispAB).toFixed(3).split('.')[1],
    };
  }
  const totVsLAB  = Object.values(splitsRaw).reduce((s, d) => s + d.vsLAB,  0);
  const totVsLH   = Object.values(splitsRaw).reduce((s, d) => s + d.vsLH,   0);
  const totRispAB = Object.values(splitsRaw).reduce((s, d) => s + d.rispAB, 0);
  const totRispH  = Object.values(splitsRaw).reduce((s, d) => s + d.rispH,  0);
  splits['通算'] = {
    vsLeft: totVsLAB  === 0 ? '--' : (totVsLH  / totVsLAB ).toFixed(3).split('.')[1],
    risp:   totRispAB === 0 ? '--' : (totRispH / totRispAB).toFixed(3).split('.')[1],
  };

  // 走力 3段階方式
  // ① Baseball Savant ≥50 → そのまま採用
  // ② Savant <50 or なし → MLB The Show SPD と ③SB計算値の平均
  // ③ MLB The Showもなし → SB計算値のみ
  //
  // ③ SB計算式
  //   base値は盗塁試行率(500PA換算)で決定:
  //     試行≥8 → base60 (net10→70, 15→75, 20→80 … 40→100, 80→140 スケール維持)
  //     試行3〜7 → base50 (時々走る)
  //     試行1〜2 → base40 (ほぼ走らない)
  //     試行0   → base30 (走らない＝遅い選手と判定)
  //   speed = base + netSBper500  (範囲: 20〜140)
  function calcSBSpeed(b) {
    if (!b) return 30;
    const pa500 = (b.pa || 0) + (b.bb || 0);
    if (!pa500) return 30;
    const sb = b.sb || 0, cs = b.cs || 0;
    const netSB = sb - cs;
    const netSBper500 = netSB * 500 / pa500;
    const attPer500   = (sb + cs) * 500 / pa500;
    const base = attPer500 >= 8 ? 60
               : attPer500 >= 3 ? 50
               : attPer500 >= 1 ? 40
               : 30;
    return Math.max(20, Math.min(140, Math.round(base + netSBper500)));
  }
  function getRawSpeedInput(yr) {
    const b   = basic[yr];
    const ss  = sprintSpeed[yr];
    if (ss != null && !isNaN(Number(ss)) && Number(ss) >= 50) return Number(ss); // ①
    const sbSpd = calcSBSpeed(b);                                                 // ③ SBベース
    const ms  = mlbTheShowSpeed[yr];
    if (ms != null && !isNaN(Number(ms))) return Math.round((Number(ms) + sbSpd) / 2); // ②+③
    return sbSpd;                                                                  // ③のみ
  }
  const totalG = years.reduce((s, yr) => s + (basic[yr]?.g || 0), 0);
  const careerRawSpeed = totalG === 0 ? 50
    : Math.round(years.reduce((s, yr) => s + getRawSpeedInput(yr) * (basic[yr]?.g || 0), 0) / totalG);

  const pitchBA = {};
  for (const yr of years) {
    const d = rawPitch[yr];
    const fmt = v => (!v || v.ba === '--') ? '--' : v.ba.replace('__D__', '');
    pitchBA[yr] = {
      ff: fmt(d.ff), si: fmt(d.si), ch: fmt(d.ch),
      sl: weightedBA([d.sl, d.st]),
      cu: fmt(d.cu), fc: fmt(d.fc), fs: fmt(d.fs),
    };
  }
  pitchBA['通算'] = {
    ff: weightedBA(years.map(yr => rawPitch[yr].ff)),
    si: weightedBA(years.map(yr => rawPitch[yr].si)),
    ch: weightedBA(years.map(yr => rawPitch[yr].ch)),
    sl: weightedBA(years.flatMap(yr => [rawPitch[yr].sl, rawPitch[yr].st])),
    cu: weightedBA(years.map(yr => rawPitch[yr].cu)),
    fc: weightedBA(years.map(yr => rawPitch[yr].fc)),
    fs: weightedBA(years.map(yr => rawPitch[yr].fs)),
  };

  const positions = ['C','1B','2B','3B','SS','LF','CF','RF'];
  const fieldingCareer = {};
  for (const pos of positions) {
    const entries = years.map(yr => fieldingByYear[yr]?.[pos]).filter(Boolean);
    if (!entries.length) continue;
    const totalOuts = entries.reduce((s, e) => s + innToOuts(e.inn), 0);
    const wDRS = totalOuts === 0 ? 0
      : entries.reduce((s, e) => s + e.drs * innToOuts(e.inn), 0) / totalOuts;
    fieldingCareer[pos] = { inn: addInningsList(entries.map(e => e.inn)), drs: Math.round(wDRS) };
  }
  function getF(yk, pos, field) {
    const v = (yk === '通算' ? fieldingCareer : (fieldingByYear[yk] || {}))[pos]?.[field];
    if (v == null || (typeof v === 'number' && isNaN(v))) return '--';
    return v;
  }

  const cols0 = [
    '選手名','年度','チーム','試合','打数','得点','安打','二塁打','三塁打','本塁打',
    '打点','四球','三振','盗塁','盗塁死','打率','出塁率','OPS',
    '対左打率','得点圏打率','走力',
    '４シーム','シンカー/2シーム','チェンジアップ','スライダー','カーブ','カット','スプリット',
  ];
  const nStat = cols0.length;
  const hRow0 = [...cols0, ...positions.flatMap(p => [p, ''])];
  const hRow1 = [...Array(nStat).fill(''), ...positions.flatMap(() => ['Inn','DRS'])];

  function buildRow(yk) {
    const b = basic[yk], sp = splits[yk], pt = pitchBA[yk];
    return [
      playerName, yk, b.team, b.g, b.pa, b.r, b.h, b.d, b.t, b.hr,
      b.rbi, b.bb, b.so, b.sb, b.cs,
      b.avg.slice(1), b.obp.slice(1),
      // OPS が 1.000 以上の場合は先頭の "1" が消えないよう整数4桁で表示
      (s => parseFloat(s || 0) >= 1 ? String(Math.round(parseFloat(s) * 1000)) : (s || '--').slice(1))(b.ops),
      sp.vsLeft, sp.risp, yk === '通算' ? careerRawSpeed : getRawSpeedInput(yk),
      pt.ff, pt.si, pt.ch, pt.sl, pt.cu, pt.fc, pt.fs,
      ...positions.flatMap(p => [getF(yk, p, 'inn'), getF(yk, p, 'drs')]),
    ];
  }

  const allRows = [hRow0, hRow1, ...years.map(buildRow), buildRow('通算')];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(allRows);
  ws['!merges'] = [
    ...Array.from({ length: nStat }, (_, c) => ({ s:{r:0,c}, e:{r:1,c} })),
    ...positions.map((_, i) => { const c = nStat + i*2; return { s:{r:0,c}, e:{r:0,c:c+1} }; }),
  ];
  ws['!cols'] = [
    {wch:12},
    {wch:6},{wch:6},{wch:5},{wch:5},{wch:5},{wch:5},{wch:7},{wch:7},{wch:7},
    {wch:5},{wch:5},{wch:5},{wch:5},{wch:7},{wch:6},{wch:6},{wch:6},
    {wch:8},{wch:8},{wch:6},
    {wch:9},{wch:13},{wch:12},{wch:9},{wch:7},{wch:8},{wch:9},
    ...Array(16).fill({wch:7}),
  ];
  XLSX.utils.book_append_sheet(wb, ws, playerName + '成績');

  const note = [
    ['項目','説明'],
    ['打数','atBats（四球・死球・犠打飛を含まない）'],
    ['走力','①Savant≥50→そのまま ②Savant<50orなし→(MLBTS+③SB値)/2 ③MLBTSもなし→SB値のみ。SB計算: base(試行≥8→60, ≥3→50, ≥1→40, 0→30)+netSB×500/(打数+四球)。走らない選手は30台以下。通算は試合数加重平均'],
    ['スライダー','SL+ST(Sweeper) PA加重平均'],['球種別打率通算','PA加重平均。データなし年度は除外'],
    ['守備通算DRS','イニング加重平均（ROUND）'],['守備','FanGraphs Inn/DRS（pageitems=2000）'],
  ];
  const wsN = XLSX.utils.aoa_to_sheet(note);
  wsN['!cols'] = [{wch:20},{wch:60}];
  XLSX.utils.book_append_sheet(wb, wsN, 'データソース・備考');

  const outFile = path.join(OUT_DIR, playerName + '_成績.xlsx');
  XLSX.writeFile(wb, outFile);

  // xlsx ライブラリはフリーズペインを出力しないため ExcelJS で適用
  const ejWb = new ExcelJS.Workbook();
  await ejWb.xlsx.readFile(outFile);
  ejWb.worksheets[0].views = [{ state:'frozen', xSplit:2, ySplit:2, topLeftCell:'C3', activeCell:'C3' }];
  await ejWb.xlsx.writeFile(outFile);

  return outFile;
}

// buildExcel は async なので呼び出し元も await が必要

// ── Ability value formulas (from stats_tool) ──────────────────────────────────
function parseBA(val) {
  if (val == null) return 0;
  const s = String(val).trim();
  if (!s || s === '--') return 0;
  if (s.includes('.')) return Math.round(parseFloat(s) * 1000);
  return parseInt(s, 10) || 0;
}
function calcHRTier(hr, ab) {
  if (!ab) return 0; // IFERROR: ゼロ除算ガード
  const r = Math.round(500 * hr / ab);
  if (r >= 54) return 6;
  if (r >= 45) return 5;
  if (r >= 36) return 4;
  if (r >= 27) return 3;
  if (r >= 18) return 2;
  if (r >= 9)  return 1;
  if (r < 12)  return 0;
  return 0;
}
function calcMeet(avg, tier) {
  // HR能(tier)が上がるほど閾値を+10ずつ引き上げ、高HR選手のミートが過大評価されないよう補正
  // tier: 0=[329,142], 1=[339,152], 2=[349,162], 3=[359,172], 4=[369,182], 5=[379,192], 6=[389,202]
  const thresholds = [[329,142],[339,152],[349,162],[359,172],[369,182],[379,192],[389,202]];
  const [hi, lo] = thresholds[tier] ?? thresholds[0];
  return avg >= hi ? Math.round(85 + (avg - hi) / 4.17) : Math.round(40 + (avg - lo) / 4.2);
}
function calcPower(hr, ab) {
  if (!ab) return 40;
  const r = Math.round(500 * hr / ab);
  return r >= 30 ? r + 55 : Math.round(500 * hr / ab * 1.54 + 40);
}
function calcSpeed(u) { return u >= 50 ? u : Math.round((u + 100) / 3); }
function calcChance(diff) { return Math.round(70 + diff / 7.4); }
function calcEye(f) {
  return f >= 110 ? Math.round(70 + (f-110)/3.6) : f >= 78  ? Math.round(60 + (f-78)/3.4)
       : f >= 55  ? Math.round(50 + (f-55)/2.3)  : f >= 42  ? Math.round(40 + (f-42)/1.3)
       : f >= 33  ? Math.round(30 + (f-33)/0.9)  : Math.round(f / 1.1);
}
function calcSO(h) { return Math.round(100 - (h - 80) / 4); }
function calcVsLeft(g) { return Math.round(g / 7.4); }

// 盗塁能テーブル（守備.ods 盗塁能シート 実測値）
// 列: スピード 55, 60, 65, 70, 75, 80, 85, 90, 95, 100
// 値: 期待 netSBper500 = (盗塁成功 - 盗塁死) × 500 / (打数 + 四球)
// ※ テーブル範囲外のスピードは端値でクランプし、範囲内は線形補間
const STEAL_ABILITY_TABLE = [
  { ability: -10, vals: [ 0,  0,  0,  0,  1,  3,  5,  6,  8, 10] },
  { ability:   0, vals: [ 0,  1,  3,  4,  5,  6,  7,  8, 10, 13] },
  { ability:  10, vals: [ 1,  3,  6,  9, 12, 15, 18, 21, 24, 27] },
  { ability:  20, vals: [ 6,  8, 10, 12, 16, 18, 21, 24, 28, 32] },
  { ability:  30, vals: [ 9, 13, 16, 20, 24, 26, 28, 31, 34, 37] },
  { ability:  40, vals: [12, 16, 20, 24, 28, 32, 36, 40, 44, 48] },
  { ability:  50, vals: [15, 21, 27, 33, 39, 45, 51, 57, 63, 70] },
];
const STEAL_SPD_MIN  = 55;  // テーブル最小スピード
const STEAL_SPD_STEP =  5;  // スピード刻み
function calcStealAbility(speed, ab, bb, sb, cs) {
  const pa    = (ab || 0) + (bb || 0);
  const netSB = (sb || 0) - (cs || 0);
  const n     = pa > 0 ? netSB * 500 / pa : 0; // netSBper500

  // スピードを [55, 100] にクランプして線形補間
  const spd  = Math.max(STEAL_SPD_MIN, Math.min(100, speed));
  const raw  = (spd - STEAL_SPD_MIN) / STEAL_SPD_STEP;
  const lo   = Math.floor(raw);
  const hi   = Math.min(lo + 1, STEAL_ABILITY_TABLE[0].vals.length - 1);
  const frac = raw - lo;

  let bestAbility = -10, bestDist = Infinity;
  for (const row of STEAL_ABILITY_TABLE) {
    const expected = row.vals[lo] + frac * (row.vals[hi] - row.vals[lo]);
    const dist = Math.abs(n - expected);
    if (dist < bestDist) { bestDist = dist; bestAbility = row.ability; }
  }
  return bestAbility;
}

// キャッチャーデータ（フレーミング + 守備成績）を統合
function buildCatcherData(catcherFielding, catcherFraming) {
  const result = { byYear: {}, career: null };
  const allYears = new Set([
    ...Object.keys(catcherFielding?.byYear || {}),
    ...Object.keys(catcherFraming?.byYear  || {}),
  ]);
  for (const yr of allYears) {
    result.byYear[yr] = {
      fielding: catcherFielding?.byYear?.[yr] || null,
      framing:  catcherFraming?.byYear?.[yr]  || null,
    };
  }
  if (catcherFielding?.career || catcherFraming?.career) {
    result.career = {
      fielding: catcherFielding?.career || null,
      framing:  catcherFraming?.career  || null,
    };
  }
  return result;
}

const PITCH_BASE      = [277, 289, 238, 218, 210, 257, 215];
const PITCH_OUT_ORDER = [0, 1, 5, 3, 4, 2, 6];
function calcPitchRatings(pitchVals) {
  const mVals = pitchVals.map((r, i) => r ? (r - PITCH_BASE[i]) / 7 : null);
  const valid  = mVals.filter(m => m !== null);
  if (!valid.length) return Array(7).fill('');
  const n13 = valid.reduce((a, b) => a + b, 0) / valid.length;
  return PITCH_OUT_ORDER.map(i => (mVals[i] !== null ? Math.round(mVals[i] - n13) : ''));
}

const DEF_POSITIONS = [
  { label:'C',  innCol:29, drsCol:30 }, { label:'1B', innCol:31, drsCol:32 },
  { label:'2B', innCol:33, drsCol:34 }, { label:'3B', innCol:35, drsCol:36 },
  { label:'SS', innCol:37, drsCol:38 }, { label:'LF', innCol:39, drsCol:40 },
  { label:'CF', innCol:41, drsCol:42 }, { label:'RF', innCol:43, drsCol:44 },
];
function parseInn(val) {
  if (val == null) return null;
  const n = parseFloat(String(val).trim());
  return isNaN(n) ? null : n;
}
function parseDRS(val) {
  if (val == null) return null;
  const s = String(val).trim();
  if (!s || s === '--') return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}
function calcDefMain(inn, drs) {
  if (inn == null || drs == null) return null;
  return inn > 699 ? drs / inn * 1000 : inn < 700 ? drs * 1.5 : null;
}
function calcDefSub(inn, drs, mainInn) {
  if (inn == null || drs == null || !mainInn) return null;
  const pct = inn / mainInn * 100, penalty = pct < 20 ? 20 - pct : 0;
  if (inn > 499 || pct >= 50) return inn > 699 ? drs / inn * 1000 : inn < 700 ? drs * 1.5 : null;
  return drs < 0 ? drs * 500 / inn - penalty : drs - penalty;
}


const BAT_REDPURPLE_FILL = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFC00060' } };
function batRedPurpleCell(cell, value, fs) {
  cell.value = value; cell.fill = { ...BAT_REDPURPLE_FILL };
  cell.font = { bold:true, color:{ argb:'FFFFFFFF' }, size:fs };
  cell.alignment = { horizontal:'center', vertical:'middle' };
}

async function processFile(filePath, catcherData = null) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];
  const fontSize = ws.getCell(1, 1).font?.size || 11;

  // START_COL 45–55: ミート〜阻止率（11列）
  // PITCH_COL  56–62: FB〜SF（7列）
  // DEF_START_COL 63〜: 守備
  const NEW_HEADERS  = ['ミート','パワー','スピード','チャンス','選球眼','三振','HR','盗塁能','対左投手','リード','阻止率'];
  const START_COL    = 45;
  const PITCH_HEADERS = ['FB','2C','CT','SL','CB','CH','SF'];
  const PITCH_COL    = 56;
  const DEF_START_COL = 63;

  NEW_HEADERS.forEach((h, i)   => purpleCell(ws.getCell(1, START_COL + i), h, fontSize));
  PITCH_HEADERS.forEach((h, i) => batRedPurpleCell(ws.getCell(1, PITCH_COL + i), h, fontSize));

  // 全年合計Innでグローバル守備順序を決定
  const posTotalInn = {};
  DEF_POSITIONS.forEach(p => { posTotalInn[p.label] = 0; });
  ws.eachRow((row, rn) => {
    if (rn <= 2) return;
    if (!Number(row.getCell(5).value)) return;
    for (const pos of DEF_POSITIONS) {
      const inn = parseInn(row.getCell(pos.innCol).value);
      if (inn != null) posTotalInn[pos.label] += inn;
    }
  });
  const globalOrder = [...DEF_POSITIONS]
    .filter(p => posTotalInn[p.label] > 0)
    .sort((a, b) => posTotalInn[b.label] - posTotalInn[a.label]);

  if (globalOrder.length > 0) {
    purpleCell(ws.getCell(1, DEF_START_COL), '守備', fontSize);
    globalOrder.forEach((pos, i) => purpleCell(ws.getCell(2, DEF_START_COL + i), pos.label, fontSize));
  }

  const careerPosRatings = {}; // pos label → { sumWeighted, sumInn } for weighted-average career DRS
  let count = 0;
  ws.eachRow((row, rn) => {
    if (rn === 1) return;
    const ab = Number(row.getCell(5).value) || 0;
    if (!ab) return;

    const yr       = String(row.getCell(2).value || '').trim();
    const isCareer = yr === '通算';

    const hr     = Number(row.getCell(10).value) || 0;
    const walks  = Number(row.getCell(12).value) || 0;
    const k      = Number(row.getCell(13).value) || 0;
    const sb     = Number(row.getCell(14).value) || 0;
    const cs     = Number(row.getCell(15).value) || 0;
    const avg    = parseBA(row.getCell(16).value);
    const vsL    = parseBA(row.getCell(19).value);
    const clutch = parseBA(row.getCell(20).value);
    const spd    = Number(row.getCell(21).value) || 0;

    const tier     = calcHRTier(hr, ab);
    const walkRate = ab > 0 ? walks / ab * 1000 : 0;
    const soRate   = k / ab * 1000;
    const spdVal   = calcSpeed(spd);

    [calcMeet(avg,tier), calcPower(hr,ab), spdVal, calcChance(clutch-avg),
     calcEye(walkRate), calcSO(soRate), tier,
     calcStealAbility(spdVal, ab, walks, sb, cs),
     calcVsLeft(vsL-avg)]
      .forEach((v, i) => purpleCell(ws.getCell(rn, START_COL + i), v, fontSize));

    // リード (START_COL+9) と 阻止率 (START_COL+10) — 捕手のみ
    {
      const cInn    = parseInn(row.getCell(29).value); // C Inn列
      const yrData  = catcherData?.byYear?.[yr]  ?? null;
      const carData = catcherData?.career         ?? null;
      // 2002年以前は FanGraphs 対象外で C Inn 列が空になるため
      // MLB Stats API の捕手守備データがある年も捕手と判定する
      const hasCatcher = (cInn != null && cInn > 0)
        || (!isCareer && (yrData?.fielding?.g  ?? 0) > 0)
        || ( isCareer && (carData?.fielding?.g ?? 0) > 0);
      if (hasCatcher) {

        // リード: Baseball Savant フレーミング実データ（2018年以降）
        //         2017年以前 or データなし → CS% vs 時代別リーグ平均から仮想算出
        // ※年別行はキャリア framing にフォールバックしない（誤った通算値が入るのを防ぐ）
        const framingData = yrData?.framing ?? (isCareer ? carData?.framing : null);
        let leadVal = 0;
        if (framingData && framingData.pitches > 0) {
          // ── 実データ: pitches 1500換算フレーミングrun ──────────────────────
          leadVal = Math.round(framingData.runs * 1500 / framingData.pitches);
        } else {
          // ── 仮想リード: CS%盗塁阻止成分 + PB抑制ブロッキング成分 ─────────
          // 捕手の総合守備力（リード代替）を2成分で算出
          const fld = yrData?.fielding ?? carData?.fielding;
          if (fld) {
            const cs    = fld.cs || 0;
            const sb    = fld.sb || 0;
            const pb    = fld.pb || 0;
            const g     = fld.g  || 0;
            const total = cs + sb;
            const yrNum = isCareer ? 2000 : (parseInt(yr) || 2000);

            // ── 成分①: 盗塁阻止能力 (CS% vs 時代別リーグ平均) ────────────
            // 時代別リーグ平均CS%（歴史的統計から）
            const lgCsPct = yrNum < 1970 ? 0.38
                          : yrNum < 1985 ? 0.36
                          : yrNum < 1995 ? 0.33
                          : yrNum < 2005 ? 0.30
                          : yrNum < 2015 ? 0.28
                          :                0.26;
            // サンプル10以上のみ推計。係数0.35: CS%10%差×試合数で±run相当に換算
            const comp1 = total >= 10
              ? (cs / total - lgCsPct) * total * 0.35
              : 0;

            // ── 成分②: PB抑制能力（ブロッキング）─────────────────────────
            // 時代別リーグ平均PB/G（捕手の受け方技術・グローブ形状の歴史的変化）
            const lgPbPerG = yrNum < 1950 ? 0.15
                           : yrNum < 1970 ? 0.12
                           : yrNum < 1985 ? 0.10
                           : yrNum < 2000 ? 0.08
                           : yrNum < 2015 ? 0.06
                           :                0.05;
            // 20試合以上のみ推計。係数0.80: PB1個差×試合数で±run相当に換算
            // ※PBが少ない(pbPerG < lgPbPerG)ほど正のリード値になる（符号反転）
            const pbPerG = g >= 20 ? pb / g : null;
            const comp2  = pbPerG != null
              ? -(pbPerG - lgPbPerG) * g * 0.80
              : 0;

            // ±8 でキャップ（仮想算出の推計精度を考慮した保守的な上下限）
            leadVal = Math.max(-8, Math.min(8, Math.round(comp1 + comp2)));
          }
        }
        purpleCell(ws.getCell(rn, START_COL + 9), leadVal, fontSize);

        // 阻止率: CS÷(SB+CS)×100 round（年別→career で代替）
        const fieldingData = yrData?.fielding ?? carData?.fielding; // 年別なければ通算で代替
        let csRateVal = 0;
        if (fieldingData) {
          const tot = (fieldingData.sb || 0) + (fieldingData.cs || 0);
          csRateVal = tot > 0 ? Math.round((fieldingData.cs || 0) / tot * 100) : 0;
        }
        purpleCell(ws.getCell(rn, START_COL + 10), csRateVal, fontSize);
      }
    }

    const pitchRaw  = [22,23,24,25,26,27,28].map(c => parseBA(row.getCell(c).value));
    const pitchVals = calcPitchRatings(pitchRaw);
    pitchVals.forEach((v, i) => {
      if (v !== '') batRedPurpleCell(ws.getCell(rn, PITCH_COL + i), v, fontSize);
    });

    if (globalOrder.length > 0) {
      if (isCareer) {
        // 通算行: 年度別加重平均（Inn加重）で算出した値を書き込む（-30 下限補正）
        globalOrder.forEach((gpos, i) => {
          const data = careerPosRatings[gpos.label];
          if (!data || data.sumInn === 0) return;
          const careerRating = Math.max(-30, Math.round(data.sumWeighted / data.sumInn));
          purpleCell(ws.getCell(rn, DEF_START_COL + i), careerRating, fontSize);
        });
      } else {
        // 年別行: 計算して書き込みつつ加重平均用に累積
        const yearPos = DEF_POSITIONS
          .map(p => ({ label:p.label, inn:parseInn(row.getCell(p.innCol).value), drs:parseDRS(row.getCell(p.drsCol).value) }))
          .filter(p => p.inn != null && p.inn > 0)
          .sort((a, b) => b.inn - a.inn);
        if (yearPos.length > 0) {
          const mainInn = yearPos[0].inn;
          // メイン守備比率 < 12% → DH専属とみなし全ポジションに -15 修正
          const mainInnPct = mainInn / ((ab + walks) * 2) * 100;
          const dhPenalty  = mainInnPct < 12 ? -15 : 0;
          globalOrder.forEach((gpos, i) => {
            const yp = yearPos.find(p => p.label === gpos.label);
            if (!yp) return;
            // 個別出場比率 < 2% → 非表示（端的すぎる守備）
            const defInnPct = yp.inn / ((ab + walks) * 2) * 100;
            if (defInnPct < 2) return;
            const rating = yp === yearPos[0] ? calcDefMain(yp.inn, yp.drs) : calcDefSub(yp.inn, yp.drs, mainInn);
            if (rating != null) {
              const raw = Math.round(rating) + dhPenalty;
              // DH専属でない年のみ -30 下限補正（DH専属年は補正なし）
              const finalRating = dhPenalty === 0 ? Math.max(-30, raw) : raw;
              purpleCell(ws.getCell(rn, DEF_START_COL + i), finalRating, fontSize);
              // Inn加重平均用に累積
              if (!careerPosRatings[gpos.label]) careerPosRatings[gpos.label] = { sumWeighted: 0, sumInn: 0 };
              careerPosRatings[gpos.label].sumWeighted += finalRating * yp.inn;
              careerPosRatings[gpos.label].sumInn += yp.inn;
            }
          });
        }
      }
    }
    count++;
  });

  await wb.xlsx.writeFile(filePath);
  return count;
}

async function batRunCreateJob(jobId, params) {
  // ── ログ出力: コンソール + ファイル（mlb_create_tool_log.txt）────────────
  const logLines = [];
  const logFile  = path.join(OUT_DIR, 'mlb_create_tool_log.txt');
  const upd = msg => {
    const j = jobs.get(jobId);
    if (j) { j.progress = msg; }
    const line = `[${new Date().toLocaleTimeString('ja-JP')}] ${msg}`;
    console.log('[job]', msg);
    logLines.push(line);
    // 都度ファイルに追記（ツールが途中で止まっても確認できるよう）
    try { fs.appendFileSync(logFile, line + '\n', 'utf8'); } catch {}
  };
  // ログファイルをこのジョブ開始時にリセット
  try { fs.writeFileSync(logFile, `=== ${params.name || params.fullName} ${new Date().toLocaleString('ja-JP')} ===\n`, 'utf8'); } catch {}
  try {
    upd('MLB Stats API からデータ取得中...');
    const { years, basic, splitsRaw, catcherFielding, mlbApiFielding } = await batFetchMLBStats(params.id, params.y1, params.y2);

    upd('ブラウザを起動して Baseball Savant / FanGraphs を取得中...');
    const { sprintSpeed, rawPitch, fieldingByYear, mlbTheShowSpeed, catcherFraming, bbRefSplits, battingHand } =
      await batFetchBrowserData(params.slug, params.id, params.fullName, years, upd, splitsRaw);

    // BB-Ref 通算スプリットで MLB Stats API の欠損年を補完
    // （歴代選手など sitCodes API が空を返す場合に使用）
    for (const yr of years) {
      const d  = splitsRaw[yr];
      const bb = bbRefSplits?.[yr];
      if (!d || !bb) continue;
      if (!d.vsLAB  && bb.vsLAB)  { d.vsLAB  = bb.vsLAB;  d.vsLH   = bb.vsLH;  }
      if (!d.rispAB && bb.rispAB) { d.rispAB = bb.rispAB; d.rispH  = bb.rispH; }
    }

    // ── プラトーン推計: BB-Ref にも対左/得点圏データがない場合は打席情報から推計 ──
    // 対象: vsLAB === 0 の年（実データが一切取れなかった年）
    // 推計根拠:
    //   左打者(LHB) vs 左投手: 通算BAの約88% (同側投手に対し約-30〜35点が統計的平均)
    //   右打者(RHB) vs 左投手: 通算BAの約104% (対左投手に約+10〜15点有利)
    //   両打(Switch) vs 左投手: 右打席に入るため右打者に準じ約103%
    //   得点圏打率: 個人差が大きいが平均的には通算BAと同値 → そのまま使用
    {
      // 打席係数
      const platoonCoeff = battingHand === 'L' ? 0.88
                         : battingHand === 'R' ? 1.04
                         : battingHand === 'S' ? 1.03
                         : null;  // 不明時は推計しない

      const stillMissingVsL  = years.filter(yr => !(splitsRaw[yr]?.vsLAB));
      const stillMissingRisp = years.filter(yr => !(splitsRaw[yr]?.rispAB));

      if (stillMissingVsL.length > 0 || stillMissingRisp.length > 0) {
        for (const yr of years) {
          const d  = splitsRaw[yr];
          const b  = basic[yr];
          if (!d || !b) continue;
          const ab = b.pa || 0;   // basic.pa は atBats
          const h  = b.h  || 0;
          if (ab === 0) continue;
          const yearBA = h / ab;

          // 対左打率推計（打席情報がある年のみ・データ欠損年のみ）
          if (!d.vsLAB && platoonCoeff !== null) {
            const estBA  = Math.min(Math.max(yearBA * platoonCoeff, 0), 1);
            d.vsLAB = ab;                          // 実ABを使って加重平均を正確に
            d.vsLH  = Math.round(estBA * ab);
          }

          // 得点圏打率推計（打席情報不要・データ欠損年のみ）
          // RBI余剰ロジック:
          //   extraRBI = RBI - HR（本塁打以外で稼いだ打点）
          //   expectedExtraRBI = H × 0.30（安打あたり平均RBI率）
          //   surplus > 0 → RISP状況で多く打点を挙げている → 得点圏BAを上方修正
          //   surplus < 0 → RISP状況でRBI少ない → 得点圏BAを下方修正
          //   調整幅 ±0.030（MLB実測RISP差の範囲内）
          if (!d.rispAB) {
            const rbi = b.rbi || 0;
            const hr  = b.hr  || 0;
            const extraRBI         = rbi - hr;
            const expectedExtraRBI = h * 0.30;  // MLB平均: 安打の約30%がRBIにつながる
            const surplus          = extraRBI - expectedExtraRBI;
            // surplus / expectedExtraRBI = 相対的RBI余剰率
            // × 0.05 でBA調整値に変換（余剰率100% → +5点調整）
            const rispAdj = expectedExtraRBI > 0
              ? Math.max(-0.030, Math.min(0.030, (surplus / expectedExtraRBI) * 0.05))
              : 0;
            const rispBA = Math.max(0.100, Math.min(0.999, yearBA + rispAdj));
            d.rispAB = ab;
            d.rispH  = Math.round(rispBA * ab);
          }
        }
        if (platoonCoeff !== null) {
          const handLabel = battingHand === 'L' ? '左打' : battingHand === 'R' ? '右打' : '両打';
          upd(`プラトーン推計(${handLabel}): 対左=${stillMissingVsL.length}年, RISP=${stillMissingRisp.length}年 を補完`);
        } else {
          upd(`得点圏打率推計(RBI余剰): ${stillMissingRisp.length}年を補完（打席情報不明のため対左は推計スキップ）`);
        }
      }
    }

    // ── 2002年以前 守備フォールバック ──────────────────────────────────────────
    // FanGraphs は 2003 以降のみ。BB-Ref が取得できなかった年は
    // MLB Stats API の yearByYear fielding データで Inn/DRS を補完する。
    {
      // ポジション別リーグ平均 RF/9 と守備率（時代を問わず近似値）
      const LG_RF9 = { C:6.5, '1B':8.8, '2B':4.8, '3B':2.9, SS:4.5, LF:2.1, CF:2.7, RF:2.1, OF:2.3 };
      const LG_FLD = { C:.994, '1B':.994, '2B':.980, '3B':.958, SS:.969, LF:.987, CF:.988, RF:.986, OF:.987 };

      const PRE2003 = years.filter(y => parseInt(y) < 2003);
      for (const yr of PRE2003) {
        if (!fieldingByYear[yr]) fieldingByYear[yr] = {};
        const apiYr = mlbApiFielding.byYear[yr];
        if (!apiYr) continue;
        for (const [pos, d] of Object.entries(apiYr)) {
          if (fieldingByYear[yr][pos]) continue;   // FanGraphs / BB-Ref 優先
          if (!d.g || d.g === 0) continue;

          // ── イニング: API 値があれば使う、なければ試合数×推定値 ────────────
          let innFmt = null, innDec = 0;
          if (d.inn) {
            const s = String(d.inn).replace(/[,\s]/g, '');
            const m = s.match(/^(\d+)(?:\.(\d))?$/);
            if (m) {
              innDec = parseInt(m[1]) + parseInt(m[2] || '0') / 3;
              if (innDec >= 1) innFmt = s;
            }
          }
          if (!innFmt) {
            const avgInn = pos === 'C' ? 8.5 : 8.7;
            const outs   = Math.round(d.g * avgInn * 3);
            innDec  = d.g * avgInn;
            innFmt  = Math.floor(outs / 3) + '.' + (outs % 3);
          }

          // ── 推定 DRS ─────────────────────────────────────────────────────
          let drs = 0;
          if (pos === 'C') {
            // 捕手: CS% vs 時代別リーグ平均（BB-Ref 仮想DRS と同一係数）
            const total = (d.cs || 0) + (d.sb || 0);
            const yrNum = parseInt(yr);
            const lgCs  = yrNum < 1970 ? 0.38 : yrNum < 1985 ? 0.36
                        : yrNum < 1995 ? 0.33 : yrNum < 2005 ? 0.30
                        : yrNum < 2015 ? 0.28 : 0.26;
            drs = total >= 10 ? Math.round((d.cs / total - lgCs) * total * 0.40) : 0;
          } else {
            // 捕手以外: RF/9差 × イニング / 9 × 0.25 + エラー差成分
            // RF/9 = API直値 または (刺殺+補殺) / イニング × 9 で算出
            let rf9 = (typeof d.rf9 === 'number') ? d.rf9 : null;
            if (rf9 == null && d.po != null && d.a != null && innDec > 0) {
              rf9 = (d.po + d.a) / innDec * 9;
            }
            const lgRf9 = LG_RF9[pos] ?? 0;
            const rangeDRS = (rf9 != null && lgRf9 > 0)
              ? (rf9 - lgRf9) * innDec / 9 * 0.25
              : 0;
            const fldNum = d.fld ? parseFloat(String(d.fld)) : 0;
            const lgFld  = LG_FLD[pos] ?? 0;
            const ch     = d.ch ?? 0;
            const errorDRS = (ch > 0 && lgFld > 0 && fldNum > 0)
              ? ch * (fldNum - lgFld) * 0.5
              : 0;
            drs = Math.round(rangeDRS + errorDRS);
          }

          fieldingByYear[yr][pos] = { inn: innFmt, drs, g: d.g };
        }
        const added = Object.keys(fieldingByYear[yr]);
        if (added.length) upd(`MLB API 守備補完 ${yr}: ${added.join(', ')}`);
      }
    }

    upd('Excel ファイルを生成中...');
    const outFile = await batBuildExcel(params.name, years, basic, splitsRaw, sprintSpeed, mlbTheShowSpeed, rawPitch, fieldingByYear);

    upd('能力値を計算・追加中...');
    const catcherData = buildCatcherData(catcherFielding, catcherFraming);
    const rows = await processFile(outFile, catcherData);

    const j = jobs.get(jobId);
    if (j) { j.status = 'done'; j.result = path.basename(outFile); j.rows = rows; j.progress = '完了'; }
  } catch (e) {
    const j = jobs.get(jobId);
    if (j) { j.status = 'error'; j.error = e.message; j.progress = 'エラー'; }
    console.error('[job error]', e.message);
  }
}

// ── Job management ────────────────────────────────────────────────────────────
const jobs = new Map();


// ── HTML ──────────────────────────────────────────────────────────────────────
const HTML = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>MLB成績ツール</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Meiryo UI','Meiryo','Yu Gothic UI',sans-serif;background:#f0e8f5;
  min-height:100vh;display:flex;align-items:flex-start;justify-content:center;padding:16px 18px}
.card{background:white;border-radius:12px;box-shadow:0 4px 20px rgba(112,48,160,.15);
  padding:22px 26px;width:100%;max-width:640px}
h1{color:#7030A0;font-size:19px;margin-bottom:10px}
.mode-toggle{display:flex;gap:8px;margin-bottom:14px}
.mode-btn{padding:7px 17px;border:2px solid #7030A0;border-radius:20px;cursor:pointer;
  font-size:14px;font-weight:bold;background:white;color:#7030A0;transition:all .15s}
.mode-btn.active{background:#7030A0;color:white}
.mode-btn:hover:not(.active){background:#f3e5f5}
.mode-panel{display:none}.mode-panel.active{display:block}
.tabs{display:flex;gap:0;margin-bottom:16px;border-bottom:2px solid #e0c8f0}
.tab{padding:8px 20px;cursor:pointer;font-size:13px;font-weight:bold;color:#999;
  border-bottom:3px solid transparent;margin-bottom:-2px;transition:all .15s}
.tab.active{color:#7030A0;border-bottom-color:#7030A0}
.tab:hover:not(.active){color:#555}
.panel{display:none}.panel.active{display:block}
.sec{margin-bottom:13px}
label{display:block;font-size:12px;font-weight:bold;color:#555;margin-bottom:4px}
input{width:100%;padding:7px 11px;border:1px solid #ddd;border-radius:6px;
  font-size:13px;font-family:inherit}
input:focus{outline:none;border-color:#7030A0}
.row{display:flex;gap:8px}.row>div{flex:1}
button{padding:9px 20px;border:none;border-radius:6px;cursor:pointer;
  font-size:13px;font-family:inherit;font-weight:bold;transition:all .15s}
.btn-s{background:#555;color:white;white-space:nowrap}
.btn-s:hover{background:#333}
.btn-p{background:#7030A0;color:white}
.btn-p:hover:not(:disabled){background:#5a1e85}
.btn-p:disabled{background:#bbb;cursor:not-allowed}
.ri{padding:6px 10px;background:#f8f8f8;border:1px solid #eee;border-radius:4px;
  cursor:pointer;margin-bottom:3px;font-size:13px;transition:background .1s}
.ri:hover{background:#f0e8f5;border-color:#ce93d8}
.ri .n{font-weight:bold;color:#333}.ri .m{color:#888;font-size:11px;margin-left:8px}
.results{margin-top:5px}
.file-area{border:2px dashed #ce93d8;border-radius:8px;padding:10px 12px;min-height:44px;
  display:flex;align-items:center;gap:10px;background:#faf5ff;margin-bottom:8px}
.file-area.sel{border-color:#7030A0;background:#f3e5f5}
.file-icon{font-size:22px;flex-shrink:0}
.file-path{font-size:13px;color:#888;flex:1;word-break:break-all}
.file-path.has{color:#4a1470;font-weight:bold}
.pbox{margin-top:8px;padding:8px 12px;background:#f9f5ff;border:1px solid #ce93d8;
  border-radius:8px;display:none}
.ptxt{font-size:13px;color:#555}
.sp{display:inline-block;width:12px;height:12px;border:2px solid #ddd;
  border-top-color:#7030A0;border-radius:50%;animation:spin .7s linear infinite;
  vertical-align:middle;margin-right:6px}
@keyframes spin{to{transform:rotate(360deg)}}
.done{margin-top:10px;padding:11px 14px;background:#e8f5e9;border:1px solid #a5d6a7;
  border-radius:8px;display:none;font-size:13px;color:#2e7d32}
.err{margin-top:10px;padding:11px 14px;background:#ffebee;border:1px solid #ef9a9a;
  border-radius:8px;display:none;font-size:13px;color:#c62828;white-space:pre-wrap}
.note{font-size:11px;color:#aaa;margin-top:10px;line-height:1.6}
.badge-row{display:flex;flex-wrap:wrap;gap:4px;margin-bottom:10px}
.badge{background:#f3e5f5;color:#7030A0;border:1px solid #ce93d8;border-radius:4px;
  padding:3px 8px;font-size:11px;font-weight:bold}
.badge.red{background:#fce4ec;color:#c00060;border-color:#f48fb1}
.opt-label{font-size:11px;color:#888;margin-left:6px;font-weight:normal}
</style>
</head>
<body>
<div class="card">
  <h1>MLB成績ツール</h1>
  <div class="mode-toggle">
    <button class="mode-btn active" id="modeBtn-pitcher" onclick="setMode('pitcher')">⚾ 投手モード</button>
    <button class="mode-btn" id="modeBtn-batter" onclick="setMode('batter')">🏏 野手モード</button>
  </div>

  <!-- ── Pitcher mode ── -->
  <div id="mode-pitcher" class="mode-panel active">
    <div class="tabs" id="ptabs">
      <div class="tab active" onclick="switchTab('p-create',this,'ptabs')">新規作成</div>
      <div class="tab" onclick="switchTab('p-info',this,'ptabs')">ツール情報</div>
    </div>
    <div id="panel-p-create" class="panel active">
      <div class="sec">
        <label>① 選手検索（英語名）</label>
        <div class="row">
          <div style="flex:3"><input id="pq" type="text" placeholder="例: Yoshinobu Yamamoto"
            onkeydown="if(event.key==='Enter')doSearch('pitcher')"></div>
          <div style="flex:1"><button class="btn-s" onclick="doSearch('pitcher')">🔍 検索</button></div>
        </div>
        <div class="results" id="presults"></div>
      </div>
      <div class="sec">
        <label>② 選手情報</label>
        <div class="row">
          <div><label>英語スラッグ</label><input id="pslug" type="text" placeholder="yoshinobu-yamamoto"></div>
          <div><label>MLB ID</label><input id="ppid" type="number" placeholder="808982"></div>
        </div>
        <div class="row" style="margin-top:6px">
          <div><label>日本語名（ファイル名）</label><input id="pjaName" type="text" placeholder="山本由伸"></div>
          <div><label>英語フルネーム</label><input id="pfullName" type="text" placeholder="Yoshinobu Yamamoto"></div>
        </div>
        <div class="row" style="margin-top:6px">
          <div><label>開始年</label><input id="py1" type="number" placeholder="2024"></div>
          <div><label>終了年</label><input id="py2" type="number" placeholder="2026"></div>
        </div>
      </div>
      <div class="sec" style="margin-top:4px">
        <label>Claude APIキー <span class="opt-label">（省略可）pre-2017年のMLB The Show未収録選手に使用</span></label>
        <input id="papiKey" type="password" placeholder="sk-ant-..." oninput="saveApiKey(this.value)">
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
        <button class="btn-p" id="pbtnCreate" onclick="doCreate('pitcher')">▶ 成績ファイルを作成</button>
      </div>
      <div class="pbox" id="pcPbox"><div class="ptxt" id="pcPtxt"><span class="sp"></span>処理中...</div></div>
      <div class="done" id="pcDone"></div>
      <div class="err"  id="pcErr"></div>
      <div class="note">
        ※ Chromeが自動起動 / pre-2017年は FanGraphs→MLB The Show→Claude の順で補完 / AY〜BE 自動追加 / 出力先: 同フォルダ
      </div>
    </div>
    <div id="panel-p-info" class="panel">
      <div class="sec">
        <p style="font-size:13px;color:#555;line-height:1.7;margin-top:4px">
          <strong>投手モード</strong>は Baseball Savant・FanGraphs・MLB The Show のデータを組み合わせて球種・制球・スタミナ能力値を自動作成します。<br>
          球種データが取得できない選手（pre-2017年など）は Claude API キーを入力することで推定データを補完できます。
        </p>
      </div>
    </div>
  </div>

  <!-- ── Batter mode ── -->
  <div id="mode-batter" class="mode-panel">
    <div class="tabs" id="btabs">
      <div class="tab active" onclick="switchTab('b-create',this,'btabs')">新規作成</div>
      <div class="tab" onclick="switchTab('b-add',this,'btabs')">既存ファイルに追加</div>
    </div>
    <div id="panel-b-create" class="panel active">
      <div class="sec">
        <label>① 選手検索（英語名）</label>
        <div class="row">
          <div style="flex:3"><input id="bq" type="text" placeholder="例: Masataka Yoshida"
            onkeydown="if(event.key==='Enter')doSearch('batter')"></div>
          <div style="flex:1"><button class="btn-s" onclick="doSearch('batter')">🔍 検索</button></div>
        </div>
        <div class="results" id="bresults"></div>
      </div>
      <div class="sec">
        <label>② 選手情報</label>
        <div class="row">
          <div><label>英語スラッグ</label><input id="bslug" type="text" placeholder="masataka-yoshida"></div>
          <div><label>MLB ID</label><input id="bpid" type="number" placeholder="807799"></div>
        </div>
        <div class="row" style="margin-top:6px">
          <div><label>日本語名（ファイル名）</label><input id="bjaName" type="text" placeholder="吉田正尚"></div>
          <div><label>FanGraphs 表示名（英語）</label><input id="bfullName" type="text" placeholder="Masataka Yoshida"></div>
        </div>
        <div class="row" style="margin-top:6px">
          <div><label>開始年</label><input id="by1" type="number" placeholder="2023"></div>
          <div><label>終了年</label><input id="by2" type="number" placeholder="2026"></div>
        </div>
      </div>
      <div class="sec">
        <div class="badge-row">
          <span class="badge">成績データ取得</span>
          <span style="font-size:14px;color:#aaa;align-self:center">→</span>
          <span class="badge">Excel生成</span>
          <span style="font-size:14px;color:#aaa;align-self:center">→</span>
          <span class="badge">ミート・パワー等 自動追加</span>
          <span class="badge red">投球能力値 自動追加</span>
          <span class="badge">守備能力値 自動追加</span>
        </div>
        <button class="btn-p" id="bbtnCreate" onclick="doCreate('batter')">▶ 成績ファイルを作成</button>
      </div>
      <div class="pbox" id="bcPbox"><div class="ptxt" id="bcPtxt"><span class="sp"></span>処理中...</div></div>
      <div class="done" id="bcDone"></div>
      <div class="err"  id="bcErr"></div>
      <div class="note">※ Chromeが自動起動（Baseball Savant / FanGraphs） / 出力先: 同フォルダ</div>
    </div>
    <div id="panel-b-add" class="panel">
      <div class="sec">
        <div class="badge-row">
          <span class="badge">ミート</span><span class="badge">パワー</span>
          <span class="badge">スピード</span><span class="badge">チャンス</span>
          <span class="badge">選球眼</span><span class="badge">三振</span>
          <span class="badge">HR</span><span class="badge">盗塁能</span>
          <span class="badge">対左投手</span><span class="badge">リード</span>
          <span class="badge">阻止率</span>
          <span class="badge red">FB 2C CT SL CB CH SF</span>
          <span class="badge">守備</span>
        </div>
      </div>
      <div class="sec">
        <label>対象 Excel ファイル（成績.xlsx）</label>
        <div class="file-area" id="bfileArea">
          <div class="file-icon">📄</div>
          <div class="file-path" id="bfilePath">ファイルが選択されていません</div>
        </div>
        <div style="display:flex;gap:8px;align-items:center">
          <button class="btn-s" onclick="doBrowse()">📂 ファイルを参照...</button>
          <button class="btn-p" id="bbtnAdd" disabled onclick="doAdd()">✓ 能力値を追加</button>
          <span id="baddStatus" style="font-size:12px;color:#888"></span>
        </div>
      </div>
      <div class="pbox" id="baPbox"><div class="ptxt" id="baPtxt"><span class="sp"></span>処理中...</div></div>
      <div class="done" id="baDone"></div>
      <div class="err"  id="baErr"></div>
    </div>
  </div>
</div>

<script>
let pcTimer = null, bcTimer = null, selectedPath = '';

(function(){
  const k = localStorage.getItem('mlb_tool_apikey');
  if (k) { const el = document.getElementById('papiKey'); if (el) el.value = k; }
})();
function saveApiKey(v) {
  if (v && v.trim()) localStorage.setItem('mlb_tool_apikey', v.trim());
  else localStorage.removeItem('mlb_tool_apikey');
}

function setMode(mode) {
  ['pitcher','batter'].forEach(m => {
    document.getElementById('modeBtn-'+m).classList.toggle('active', m===mode);
    document.getElementById('mode-'+m).classList.toggle('active', m===mode);
  });
}

function switchTab(id, el, tabsId) {
  const container = document.getElementById(tabsId).parentElement;
  container.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  container.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  document.getElementById('panel-'+id).classList.add('active');
}

async function doSearch(mode) {
  const prefix = mode === 'pitcher' ? 'p' : 'b';
  const q = document.getElementById(prefix+'q').value.trim();
  if (!q) return;
  const el = document.getElementById(prefix+'results');
  el.innerHTML = '<div style="font-size:12px;color:#888">検索中...</div>';
  try {
    const r = await fetch('/api/search?name=' + encodeURIComponent(q));
    const data = await r.json();
    if (!data.length) { el.innerHTML = '<div style="font-size:12px;color:#888">見つかりませんでした</div>'; return; }
    el.innerHTML = data.map(p =>
      '<div class="ri" data-mode="' + mode + '" data-id="' + p.id + '" data-name="' + p.name.replace(/"/g,'&quot;') + '" data-debut="' + p.debut + '" onclick="pickEl(this)">' +
      '<span class="n">' + p.name + '</span>' +
      '<span class="m">' + p.position + ' · debut ' + p.debut + ' · ID: ' + p.id + '</span></div>'
    ).join('');
  } catch(e) { el.innerHTML = '<div style="font-size:12px;color:#c00">エラー: '+e.message+'</div>'; }
}

function pickEl(el) { pick(el.dataset.mode, parseInt(el.dataset.id), el.dataset.name, el.dataset.debut); }
function pick(mode, id, name, debut) {
  const prefix = mode === 'pitcher' ? 'p' : 'b';
  document.getElementById(prefix+'pid').value      = id;
  document.getElementById(prefix+'fullName').value = name;
  document.getElementById(prefix+'slug').value     = name.toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'');
  const debutStr = debut ? '、' + debut + '年デビュー' : '';
  document.getElementById(prefix+'results').innerHTML =
    '<div style="font-size:12px;color:#7030A0">✓ ' + name + '（ID: ' + id + debutStr + '）</div>';
}

async function doCreate(mode) {
  const prefix = mode === 'pitcher' ? 'p' : 'b';
  const slug     = document.getElementById(prefix+'slug').value.trim();
  const id       = parseInt(document.getElementById(prefix+'pid').value);
  const name     = document.getElementById(prefix+'jaName').value.trim();
  const fullName = document.getElementById(prefix+'fullName').value.trim();
  const y1       = parseInt(document.getElementById(prefix+'y1').value);
  const y2       = parseInt(document.getElementById(prefix+'y2').value);
  if (!slug||!id||!name||!fullName||!y1||!y2){alert('すべての項目を入力してください');return;}
  document.getElementById(prefix+'btnCreate').disabled=true;
  document.getElementById(prefix+'cPbox').style.display='block';
  document.getElementById(prefix+'cDone').style.display='none';
  document.getElementById(prefix+'cErr').style.display='none';
  setCP('処理を開始しています...',mode);
  const body = {slug,id,name,fullName,y1,y2,type:mode};
  if (mode==='pitcher') body.apiKey=(document.getElementById('papiKey').value||'').trim();
  const r=await fetch('/api/create',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify(body)});
  const {jobId}=await r.json();
  if (mode==='pitcher') pcTimer=setInterval(()=>pollCreate('pitcher',jobId),1500);
  else bcTimer=setInterval(()=>pollCreate('batter',jobId),1500);
}

async function pollCreate(mode, jobId) {
  const prefix = mode === 'pitcher' ? 'p' : 'b';
  const r=await fetch('/api/job/'+jobId); const j=await r.json();
  setCP(j.progress,mode);
  if (j.status==='done') {
    clearInterval(mode==='pitcher'?pcTimer:bcTimer);
    document.getElementById(prefix+'cPbox').style.display='none';
    document.getElementById(prefix+'btnCreate').disabled=false;
    const el=document.getElementById(prefix+'cDone'); el.style.display='block';
    if (mode==='pitcher') {
      let msg = '✓ 完了: ' + j.result + ' を作成しました';
      if (j.abilityRows > 0) msg += '（スタミナ・制球: ' + j.abilityRows + ' 行追加）';
      el.textContent = msg;
    } else {
      el.textContent='✓ 完了: '+j.result+'（'+j.rows+' 行に能力値追加済）';
    }
  } else if (j.status==='error') {
    clearInterval(mode==='pitcher'?pcTimer:bcTimer);
    document.getElementById(prefix+'cPbox').style.display='none';
    document.getElementById(prefix+'btnCreate').disabled=false;
    const el=document.getElementById(prefix+'cErr'); el.style.display='block';
    el.textContent='✗ エラー: '+j.error;
  }
}
function setCP(msg,mode){
  const prefix = mode === 'pitcher' ? 'p' : 'b';
  document.getElementById(prefix+'cPtxt').innerHTML='<span class="sp"></span>'+msg;
}

async function doBrowse() {
  document.getElementById('baddStatus').textContent='ダイアログを開いています...';
  const r=await fetch('/api/browse'); const data=await r.json();
  if (data.path) {
    selectedPath=data.path;
    const fp=document.getElementById('bfilePath');
    fp.textContent=data.path; fp.className='file-path has';
    document.getElementById('bfileArea').className='file-area sel';
    document.getElementById('bbtnAdd').disabled=false;
    document.getElementById('baddStatus').textContent='';
  } else { document.getElementById('baddStatus').textContent='選択されませんでした'; }
}
async function doAdd() {
  if (!selectedPath) return;
  document.getElementById('bbtnAdd').disabled=true;
  document.getElementById('baPbox').style.display='block';
  document.getElementById('baDone').style.display='none';
  document.getElementById('baErr').style.display='none';
  document.getElementById('baPtxt').innerHTML='<span class="sp"></span>処理中...';
  try {
    const r=await fetch('/api/process',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({filePath:selectedPath})});
    const data=await r.json();
    document.getElementById('baPbox').style.display='none';
    document.getElementById('bbtnAdd').disabled=false;
    if (data.success) {
      const el=document.getElementById('baDone'); el.style.display='block';
      el.textContent='✓ 完了: '+data.count+' 行に能力値を書き込みました';
    } else {
      const el=document.getElementById('baErr'); el.style.display='block';
      el.textContent='✗ エラー: '+data.error;
    }
  } catch(e) {
    document.getElementById('baPbox').style.display='none';
    document.getElementById('bbtnAdd').disabled=false;
    const el=document.getElementById('baErr'); el.style.display='block';
    el.textContent='✗ 通信エラー: '+e.message;
  }
}
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
        if (params.type === 'pitcher') {
          jobs.set(jobId, { status:'running', progress:'開始中...', result:null, abilityRows:0, error:null });
          pitRunCreateJob(jobId, params);
        } else {
          jobs.set(jobId, { status:'running', progress:'開始中...', result:null, rows:0, error:null });
          batRunCreateJob(jobId, params);
        }
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ jobId }));
      } catch (e) {
        res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ error: e.message }));
      }
    });
    return;
  }
  if (req.method === 'POST' && url.pathname === '/api/process') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', async () => {
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
      try {
        const { filePath } = JSON.parse(body);
        if (!filePath) throw new Error('ファイルパスが指定されていません');
        const count = await processFile(filePath);
        res.end(JSON.stringify({ success: true, count }));
      } catch (e) {
        res.end(JSON.stringify({ success: false, error: e.message }));
      }
    });
    return;
  }
  res.writeHead(404); res.end('Not found');
});

server.on('error', err => {
  if (err.code === 'EADDRINUSE') {
    const url = `http://localhost:${PORT}`;
    console.log('\n  ⚾  MLB成績ツール（既に起動済み）\n\n  URL: ' + url + '\n');
    try { spawn('cmd.exe', ['/c', 'start', '', url], { detached:true, shell:false, stdio:'ignore' }).unref(); } catch {}
    setTimeout(() => process.exit(0), 2000);
  } else {
    console.error('サーバーエラー:', err.message);
    process.exit(1);
  }
});

server.listen(PORT, '127.0.0.1', () => {
  const url = `http://localhost:${PORT}`;
  console.log('\n  ⚾  MLB成績ツール\n\n  URL: ' + url + '\n  Ctrl+C で停止\n');
  try { spawn('cmd.exe', ['/c', 'start', '', url], { detached:true, shell:false, stdio:'ignore' }).unref(); } catch {}
});
