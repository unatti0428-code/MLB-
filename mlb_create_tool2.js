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
・球種は最大5種類、使用率5%未満は省略`;

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
  if (ip >= 65) {
    const mult = (gs > 0 && (gs / g) > 0.5) ? 13 : 20;
    return Math.round(ratio * mult);
  }
  if (ip >= 50)  return Math.round(ratio * 21);
  return Math.round(ratio * 22);  // ip <= 49
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
  if (denom <= 0) return '';
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
const PITCH_GROUPS = [
  { idx: 0, name: 'フォーシーム',   startCol: 23 },
  { idx: 1, name: 'スライダー',     startCol: 27 },
  { idx: 2, name: 'チェンジアップ', startCol: 31 },
  { idx: 3, name: 'カーブ',         startCol: 35 },
  { idx: 4, name: 'カットボール',   startCol: 39 },
  { idx: 5, name: 'ツーシーム',     startCol: 43 },
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
function calcKankyuu(eraStr, seikyu, pitchData) {
  const v1 = _kERA(parseFloat(String(eraStr || '').trim()));
  if (v1 === null) return '';
  const v2 = _kSeikyu(seikyu !== '' && seikyu != null ? Number(seikyu) : NaN);
  if (v2 === null) return '';

  // 各変化球の球威を計算してテーブルで変換、最大値を③に採用
  const kyuiOf = (idx, pd) => {
    if (!pd) return null;
    const pctStr = String(pd.pct ?? '').trim();
    const pctNum = Number(pctStr);
    if (!pctStr || pctStr === '--' || isNaN(pctNum) || pctNum <= 0) return null;
    const baNum  = pd.ba  != null && String(pd.ba)  !== '--' ? Number(pd.ba)  : NaN;
    const slgNum = pd.slg != null && String(pd.slg) !== '--' ? Number(pd.slg) : NaN;
    const ah = calcAH_pitch(idx, baNum);
    const ai = calcAI_pitch(idx, slgNum);
    if (ah === '' || ai === '') return null;
    const aj = (Number(ah) + Number(ai)) / 2;
    const ki = calcKyuI(aj, calcAK_pitch(idx, aj, pctNum), pctNum);
    return ki !== '' ? Number(ki) : null;
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

async function addAbilityToFile(xlsxPath, showKyuiMap = {}) {
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
    redPurpleCell(ws.getCell(1, base), pg.name, fontSize);
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

    // 緩急 (①ERA + ②制球 + ③変化球威力MAX) ÷ 2
    const kankyuu = calcKankyuu(eraStr, seikyu, pitchData);
    if (kankyuu !== '') purpleCell(ws.getCell(rn, KANKYUU_COL), kankyuu, fontSize);

    // 精神
    if (eraStr && eraStr !== '--') {
      const seisin = calcSeisinFromERA(Number(eraStr));
      if (seisin !== '') purpleCell(ws.getCell(rn, SEISIN_COL), seisin, fontSize);
    }

    const sanshin = calcSanshinFromK9(so / ip * 9);
    if (sanshin !== '') purpleCell(ws.getCell(rn, SANSHIN_COL), sanshin, fontSize);

    const omosa = calcOmosaFromHR9(hr / ip * 9);
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
        // MLB The Show ゲームデータから算出した球威を使用 (pre-2017)
        const showKyui = showKyuiMap[yr]?.[pg.idx];
        if (showKyui !== undefined) kyui = showKyui;
      }
      if (kyui !== '') redPurpleCell(ws.getCell(rn, base + 1), kyui, fontSize);

      redPurpleCell(ws.getCell(rn, base + 2), pctNum, fontSize);
    });

    count++;
  }

  await wb.xlsx.writeFile(xlsxPath);
  return count;
}

// ── Pitching stats fetch ──────────────────────────────────────────────────────
async function fetchPitchingStats(id, y1, y2) {
  const yby = await mlbGet(
    `https://statsapi.mlb.com/api/v1/people/${id}/stats?stats=yearByYear&group=pitching&sportId=1`
  );
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
      teamStr = (primary.team?.abbreviation || primary.team?.name?.slice(0,3)?.toUpperCase() || '???') + named.length;
    } else {
      teamStr = row.team?.abbreviation || row.team?.name?.slice(0,3)?.toUpperCase() || '???';
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

  const careerData = await mlbGet(
    `https://statsapi.mlb.com/api/v1/people/${id}/stats?stats=career&group=pitching&sportId=1`
  );
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

  const vsLeftByYear = {};
  await Promise.all(years.map(async yr => {
    try {
      const vl = await mlbGet(
        `https://statsapi.mlb.com/api/v1/people/${id}/stats?stats=statSplits&group=pitching&sportId=1&sitCodes=vl&season=${yr}`
      );
      vsLeftByYear[yr] = vl.stats[0]?.splits[0]?.stat?.avg || '--';
    } catch { vsLeftByYear[yr] = '--'; }
  }));

  try {
    const carVL = await mlbGet(
      `https://statsapi.mlb.com/api/v1/people/${id}/stats?stats=careerStatSplits&group=pitching&sportId=1&sitCodes=vl`
    );
    vsLeftByYear['通算'] = carVL.stats[0]?.splits[0]?.stat?.avg || '--';
  } catch { vsLeftByYear['通算'] = '--'; }

  return { years, basic, vsLeftByYear };
}

// ── Pitch type config ─────────────────────────────────────────────────────────
const PITCH_KEYS     = ['ff', 'sl', 'ch', 'cu', 'fc', 'si', 'fs'];
const PITCH_NAMES_JA = ['4シーム', 'スライダー', 'チェンジアップ', 'カーブ', 'カット', 'シンカー', 'スプリット'];

// Baseball Savant pitch name → 球種キー
const PITCH_MAP_P    = {
  '4-Seam Fastball': 'ff', '4-seam Fastball': 'ff', 'Four-Seam Fastball': 'ff',
  'Four Seamer': 'ff', 'Four-Seamer': 'ff', '4-Seamer': 'ff', '4 Seamer': 'ff',
  'Fastball': 'ff',        // 初期Statcat年代の汎用分類
  'Riding Fastball': 'ff', 'Rising Fastball': 'ff',  // 高スピン4シームの新分類名
  'Slider': 'sl', 'Sweeper': 'sl', 'Hard Slider': 'sl',
  'Changeup': 'ch', 'Change-up': 'ch',
  'Curveball': 'cu', 'Knuckle Curve': 'cu', 'Knuckleball': 'cu', 'Slow Curve': 'cu',
  'Cutter': 'fc',
  'Sinker': 'si', 'Two-Seam Fastball': 'si', '2-Seam Fastball': 'si',
  'Split-Finger': 'fs', 'Splitter': 'fs', 'Split Finger': 'fs',
};

// ※ brooksbaseball.net は廃止のため PITCH_MAP_B 削除済み (2025-05)

// Baseball Savant JSON API pitch_type コード → 球種キー
const PITCH_TYPE_JSON = {
  'FF': 'ff', 'FA': 'ff', 'FT': 'ff',       // 4-Seam / generic / Two-seam（4S優先）
  'SI': 'si',                                  // Sinker
  'SL': 'sl', 'ST': 'sl', 'SV': 'sl',       // Slider / Sweeper / Slurve
  'CH': 'ch', 'SC': 'ch',                    // Changeup / Screwball
  'CU': 'cu', 'KC': 'cu', 'CS': 'cu',       // Curveball / Knuckle-curve
  'FC': 'fc',                                  // Cutter
  'FS': 'fs', 'FO': 'fs',                    // Split-finger / Forkball
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
  'Slider': 1, 'Sweeper': 1, 'Slurve': 1,
  'Changeup': 2, 'Circle Change': 2, 'Vulcan Change': 2,
  'Curveball': 3, '12-6 Curve': 3, 'Slow Curve': 3, 'Knuckle-Curve': 3, 'Power Curve': 3,
  'Cutter': 4,
  'Splitter': 6, 'Forkball': 6, 'Split-Finger': 6, 'Split Finger': 6,
};

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

// 球種数 → 推定投球割合
function estimateShowUsagePct(n) {
  const tables = { 1:[100], 2:[62,38], 3:[50,30,20], 4:[42,28,18,12], 5:[35,25,20,12,8] };
  return tables[Math.min(n, 5)] || tables[5];
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

// ── Baseball Savant browser scraping + MLB The Show API (pre-2017) ──────────
// 設計:
//  Step1: Baseball Savant JSON API を in-page fetch で直接叩く（2017年以降・HTMLパース不要）
//         → pct列は「同一年の合計≈100」ヒューリスティックで特定
//  Step2: 2016年以前の未取得年 → MLB The Show API で球種・球速・球威を取得
//         （Brooksbaseball.net は廃止のため削除）
async function fetchBrowserData(slug, id, years, onProgress, playerName = '', apiKey = '') {
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

    const rawPitch = {};
    for (const yr of years) rawPitch[yr] = emptyPitchP();

    // ── ヘルパー ─────────────────────────────────────────────────────────────
    const yearHasPct = (yr) => PITCH_KEYS.some(k => {
      const p = rawPitch[yr]?.[k]?.pct;
      return p && p !== '--' && !isNaN(parseFloat(String(p)));
    });

    // rawPitch[yr] に HTML パース結果をマージ（上書きしない）
    const mergeHtmlData = (htmlData, yr) => {
      for (const [ptName, vals] of Object.entries(htmlData)) {
        const key = PITCH_MAP_P[ptName];
        if (!key) continue;
        const cur = rawPitch[yr][key];
        rawPitch[yr][key] = {
          velo: (cur.velo && cur.velo !== '--') ? cur.velo : (vals.velo || '--'),
          ba:   (cur.ba   && cur.ba   !== '--') ? cur.ba   : (vals.ba   || '--'),
          slg:  (cur.slg  && cur.slg  !== '--') ? cur.slg  : (vals.slg  || '--'),
          pct:  (cur.pct  && cur.pct  !== '--') ? cur.pct  : (vals.pct  || '--'),
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
      await page.goto(savantUrl, { waitUntil: 'networkidle2', timeout: 60000 });

      // テーブル描画を待機（最大20秒）
      try {
        await page.waitForFunction(
          () => [...document.querySelectorAll('table')].some(t =>
            /\d{4}/.test(t.innerText || '') &&
            ['4-Seam','Fastball','Sinker','Slider','Riding'].some(k => (t.innerText||'').includes(k))
          ),
          { timeout: 20000 }
        );
      } catch { /* テーブルが見つからなくても続行 */ }
      await sleep(2000);

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

          // 投球割合: ヘッダーが正確に "%" の列のみ使用
          // ※ Baseball Savant の Pitch Movement テーブル (Table 6) には % 列がなく、
          //    MPH の左列は "#"（投球数）。veloIdx-1 フォールバックは廃止。
          // ※ Run Values テーブル (Table 7) に明示的な "%" ヘッダーがある。
          const pctIdx = hdr.findIndex(h => h === '%');

          // ── 行データを年別に抽出（複数テーブルのフィールドをマージ）──
          for (const cells of allRows) {
            const yr = (cells[yearCol]?.innerText || '').trim();
            if (!yrs.includes(yr)) continue;
            const pt = (cells[pitchCol]?.innerText || '').trim();
            if (!pt || !hasPK(pt)) continue;

            const g = (idx) => (idx >= 0 && idx < cells.length)
              ? (cells[idx]?.innerText.trim() || '--') : '--';

            // velo / ba / slg はヘッダーマッチした列から（既存値があれば上書きしない）
            if (!result[yr]) result[yr] = {};
            const cur = result[yr][pt] || { velo: '--', ba: '--', slg: '--', pct: '--' };
            result[yr][pt] = {
              velo: (cur.velo && cur.velo !== '--') ? cur.velo : g(veloIdx),
              ba:   (cur.ba   && cur.ba   !== '--') ? cur.ba   : g(baIdx),
              slg:  (cur.slg  && cur.slg  !== '--') ? cur.slg  : g(slgIdx),
              // pct: "%" ヘッダーがある列から直接読む（Run Values テーブルの "%" 列）
              // 既に他のテーブルで取得済みの場合は上書きしない
              pct:  (cur.pct  && cur.pct  !== '--') ? cur.pct  : (pctIdx >= 0 ? g(pctIdx) : '--'),
            };
          }
          // break しない: 複数テーブルから全フィールドを収集する（Pitch Movement で velo、Run Values で ba/slg/pct）
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

    // ── Step 2: pre-2017 年 → MLB The Show API で球種データ取得 ──────────────────────
    // Brooksbaseball.net は廃止のため削除。2016年以前の年かつ未取得分をThe Showで補完。
    const showKyuiMap = {};
    const preShowYears = years.filter(yr => +yr < 2017 && !yearHasPct(yr));
    if (preShowYears.length > 0) {
      onProgress(`MLB The Show API 検索中 (${preShowYears.length} 年分)...`);
      try {
        const showCard = await fetchMLBTheShowCard(playerName);
        if (showCard) {
          onProgress(`MLB The Show: "${showCard.name}" (${showCard.rarity}) カード発見`);
          const pcts = estimateShowUsagePct(showCard.pitches.length);
          for (const yr of preShowYears) {
            showCard.pitches.forEach((p, i) => {
              const idx = PITCH_MAP_SHOW[p.name];
              if (idx === undefined) return;
              const key = PITCH_KEYS[idx];
              const kyui = calcKyuiFromShow(p.speed, p.control, p.movement);
              rawPitch[yr][key] = { velo: String(p.speed), ba: '--', slg: '--', pct: String(pcts[i] || 5) };
              if (kyui !== '') {
                if (!showKyuiMap[yr]) showKyuiMap[yr] = {};
                showKyuiMap[yr][idx] = kyui;
              }
            });
          }
          onProgress(`MLB The Show: ${preShowYears.length} 年 × ${showCard.pitches.length} 球種 設定完了`);
        } else {
          // ── The Show 未収録 → Claude ウェブ検索でフォールバック ─────────────
          if (apiKey) {
            onProgress('MLB The Show: 未収録 → Claude ウェブ検索で球種を推定中...');
            try {
              const claudeData = await callClaudeForPitchData(apiKey, playerName, preShowYears);
              if (claudeData && Array.isArray(claudeData.pitches) && claudeData.pitches.length > 0) {
                for (const yr of preShowYears) {
                  claudeData.pitches.forEach(p => {
                    const idx = PITCH_MAP_SHOW[p.name];
                    if (idx === undefined) return;
                    const key = PITCH_KEYS[idx];
                    rawPitch[yr][key] = { velo: String(p.speed), ba: '--', slg: '--', pct: String(p.pct) };
                    // 球威: speed のみで推定（control=70, movement=70 を仮定）
                    const kyui = calcKyuiFromShow(p.speed, 70, 70);
                    if (kyui !== '') {
                      if (!showKyuiMap[yr]) showKyuiMap[yr] = {};
                      showKyuiMap[yr][idx] = kyui;
                    }
                  });
                }
                const noteStr = claudeData.note ? `（${claudeData.note}）` : '';
                onProgress(`Claude 推定完了: ${claudeData.pitches.length} 球種 × ${preShowYears.length} 年${noteStr}`);
              } else {
                onProgress('Claude: 球種データを取得できませんでした（pre-2017 球種データなし）');
              }
            } catch (e) {
              onProgress('⚠ Claude 球種推定失敗: ' + e.message);
            }
          } else {
            onProgress('MLB The Show: 該当カード未収録（pre-2017 球種データなし / Claude APIキー未設定）');
          }
        }
      } catch (e) {
        onProgress('⚠ MLB The Show API 取得失敗: ' + e.message);
      }
    }

    return { rawPitch, showKyuiMap };
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
    const { rawPitch, showKyuiMap } = await fetchBrowserData(params.slug, params.id, years, upd, params.name, apiKey);

    upd('Excel ファイルを生成中...');
    const outFile = await buildExcel(params.name, years, basic, vsLeftByYear, rawPitch);

    upd('スタミナ・制球を計算中...');
    let abilityRows = 0;
    try {
      abilityRows = await addAbilityToFile(outFile, showKyuiMap);
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
