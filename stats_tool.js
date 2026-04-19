'use strict';
const http = require('http');
const { spawnSync } = require('child_process');
const path = require('path');
const ExcelJS = require('exceljs');

const PORT = 3939;

// ── PowerShell helpers ──────────────────────────────────────────────────────
// PowerShell → Node.js 間の日本語文字化けを防ぐため、
// パスをUTF-8 Base64でエンコードしてから渡す
function runPSBase64(script) {
  const fullScript = `
${script}
`;
  const r = spawnSync(
    'powershell.exe',
    ['-NoProfile', '-NonInteractive', '-Command', fullScript],
    { encoding: 'buffer' }  // バイナリで受け取る
  );
  return (r.stdout || Buffer.alloc(0));
}

function browseFile() {
  const r = spawnSync(
    'powershell.exe',
    ['-NoProfile', '-NonInteractive', '-Command', `
[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
$d = New-Object System.Windows.Forms.OpenFileDialog
$d.Filter = "Excel Files (*.xlsx)|*.xlsx"
$d.Title = "Excel file"
if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($d.FileName)
    [Convert]::ToBase64String($bytes)
}
`],
    { encoding: 'buffer' }
  );
  const b64 = (r.stdout || Buffer.alloc(0)).toString('ascii').trim();
  if (!b64) return '';
  return Buffer.from(b64, 'base64').toString('utf8');
}

// ── Formula helpers ─────────────────────────────────────────────────────────
// 打率などは "170"(=.170) のような3桁整数文字列で保存されている
function parseBA(val) {
  if (val == null) return 0;
  const s = String(val).trim();
  if (!s || s === '--') return 0;
  if (s.includes('.')) return Math.round(parseFloat(s) * 1000);
  return parseInt(s, 10) || 0;
}

// 守備.ods : I2 = HR率ティア (0–5)
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

// 守備.ods : K2 = ミート
function calcMeet(avg, tier) {
  const thresholds = [
    [329, 142], [339, 152], [349, 162],
    [359, 172], [369, 182], [379, 192]
  ];
  const [hi, lo] = thresholds[tier] || thresholds[0];
  return avg >= hi
    ? Math.round(85 + (avg - hi) / 4.17)
    : Math.round(40 + (avg - lo) / 4.2);
}

// 守備.ods : L2 = パワー
function calcPower(hr, ab) {
  if (!ab) return 40;
  const r = Math.round(500 * hr / ab);
  return r >= 30
    ? r + 55
    : Math.round(500 * hr / ab * 1.54 + 40);
}

// AU列 : 走力変換
function calcSpeed(u) {
  return u >= 50 ? u : Math.round((u + 100) / 3);
}

// 守備.ods : N2 = チャンス
function calcChance(diff) {
  return Math.round(70 + diff / 7.4);
}

// 守備.ods : O2 = 選球眼
function calcEye(f) {
  return f >= 110 ? Math.round(70 + (f - 110) / 3.6)
       : f >= 78  ? Math.round(60 + (f - 78)  / 3.4)
       : f >= 55  ? Math.round(50 + (f - 55)  / 2.3)
       : f >= 42  ? Math.round(40 + (f - 42)  / 1.3)
       : f >= 33  ? Math.round(30 + (f - 33)  / 0.9)
       : Math.round(f / 1.1);
}

// 守備.ods : P2 = 三振
function calcSO(h) {
  return Math.round(100 - (h - 80) / 4);
}

// 守備.ods : R2 = 対左投手
function calcVsLeft(g) {
  return Math.round(g / 7.4);
}

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
const STEAL_SPD_MIN  = 55;
const STEAL_SPD_STEP =  5;
function calcStealAbility(speed, ab, bb, sb, cs) {
  const pa    = (ab || 0) + (bb || 0);
  const netSB = (sb || 0) - (cs || 0);
  const n     = pa > 0 ? netSB * 500 / pa : 0;

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

// ── 守備.ods 投球能力値 (BA-BG列) ──────────────────────────────────────────
// 入力 R5-R11 に対応する基準被打率 (ODS N14-N20)
//   R5=4シーム, R6=シンカー, R7=チェンジアップ, R8=スライダー,
//   R9=カーブ,  R10=カット,  R11=スプリット
const PITCH_BASE = [277, 289, 238, 218, 210, 257, 215];

// 出力列(BA-BG = FB,2C,CT,SL,CB,CH,SF)は L14-L20 の順に対応
// L14=M[0](FB), L15=M[1](2C), L16=M[5](CT), L17=M[3](SL),
// L18=M[4](CB), L19=M[2](CH), L20=M[6](SF)
const PITCH_OUT_ORDER = [0, 1, 5, 3, 4, 2, 6];

function calcPitchRatings(pitchVals) {
  // pitchVals: [V,W,X,Y,Z,AA,AB] = [FB,2C,CH,SL,CB,CT,SF] batting avg ×1000
  // 0 は "--"(データなし) と同様に null 扱い (ODS の IF(R5,...) と同じ)
  const mVals = pitchVals.map((r, i) => r ? (r - PITCH_BASE[i]) / 7 : null);

  const valid = mVals.filter(m => m !== null);
  if (valid.length === 0) return Array(7).fill('');

  const n13 = valid.reduce((a, b) => a + b, 0) / valid.length; // ODS N13

  return PITCH_OUT_ORDER.map(i => {
    const m = mVals[i];
    return m !== null ? Math.round(m - n13) : '';
  });
}

// ── 守備.ods 守備能力値 (BH列以降) ──────────────────────────────────────────
// ポジション名と対応する列インデックス (ExcelJS 1-based)
const DEF_POSITIONS = [
  { label: 'C',  innCol: 29, drsCol: 30 },
  { label: '1B', innCol: 31, drsCol: 32 },
  { label: '2B', innCol: 33, drsCol: 34 },
  { label: '3B', innCol: 35, drsCol: 36 },
  { label: 'SS', innCol: 37, drsCol: 38 },
  { label: 'LF', innCol: 39, drsCol: 40 },
  { label: 'CF', innCol: 41, drsCol: 42 },
  { label: 'RF', innCol: 43, drsCol: 44 },
];

// Inn値をパース ("--" / 空 → null)
function parseInn(val) {
  if (val == null) return null;
  const s = String(val).trim();
  if (!s || s === '--') return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

// DRS値をパース ("--" / 空 → null、0 は 0 として扱う)
function parseDRS(val) {
  if (val == null) return null;
  const s = String(val).trim();
  if (!s || s === '--') return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

// 守備.ods O40 (主ポジション) : N36=0 なので調整項なし
// IFERROR(IF(O37>699, O38/O37*1000, IF(O37<700, O38*1.5, "")), "")
function calcDefMain(inn, drs) {
  if (inn == null || drs == null) return null;
  if (inn > 699) return drs / inn * 1000;
  if (inn < 700) return drs * 1.5;
  return null; // inn === 700 は "" (ODS と同様)
}

// 守備.ods P40 (サブポジション) : N36=0 なので調整項なし
// P39 = inn/mainInn*100 (守備割合%)
// P41 = IF(P39<20, 20-P39, 0) (ペナルティ)
// IF(inn > 499 OR pct >= 50)
//   IF(inn>699) → drs/inn*1000  ELSE drs*1.5
// ELSE
//   IF(drs<0) → drs*500/inn - penalty  ELSE drs - penalty
function calcDefSub(inn, drs, mainInn) {
  if (inn == null || drs == null || !mainInn) return null;
  const pct     = inn / mainInn * 100; // P39
  const penalty = pct < 20 ? 20 - pct : 0; // P41
  if (inn > 499 || pct >= 50) {
    if (inn > 699) return drs / inn * 1000;
    if (inn < 700) return drs * 1.5;
    return null;
  } else {
    if (drs < 0) return drs * 500 / inn - penalty;
    return drs - penalty; // drs >= 0
  }
}

// ── Excel処理 ───────────────────────────────────────────────────────────────
const PURPLE_FILL = {
  type: 'pattern', pattern: 'solid',
  fgColor: { argb: 'FF7030A0' }
};
const REDPURPLE_FILL = {
  type: 'pattern', pattern: 'solid',
  fgColor: { argb: 'FFC00060' }
};

function styledCell(cell, value, fill, fontSize) {
  cell.value = value;
  cell.fill = { ...fill };
  cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: fontSize };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
}
function purpleCell(cell, value, fontSize)   { styledCell(cell, value, PURPLE_FILL,    fontSize); }
function redPurpleCell(cell, value, fontSize) { styledCell(cell, value, REDPURPLE_FILL, fontSize); }

async function processFile(filePath, catcherData = null) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];

  // 既存セルのフォントサイズを取得
  const fontSize = ws.getCell(1, 1).font?.size || 11;

  // START_COL 45–55: ミート〜阻止率（11列）
  // PITCH_COL  56–62: FB〜SF（7列）
  // DEF_START_COL 63〜: 守備
  const NEW_HEADERS = ['ミート', 'パワー', 'スピード', 'チャンス', '選球眼', '三振', 'HR', '盗塁能', '対左投手', 'リード', '阻止率'];
  const START_COL = 45; // AS列

  const PITCH_HEADERS = ['FB', '2C', 'CT', 'SL', 'CB', 'CH', 'SF'];
  const PITCH_COL = 56; // BD列（リード・阻止率挿入により+2）

  const DEF_START_COL = 63; // BK列（同上）

  // ヘッダー行 (Row 1) に項目名を書き込む
  NEW_HEADERS.forEach((h, i)   => purpleCell(ws.getCell(1, START_COL + i), h, fontSize));
  PITCH_HEADERS.forEach((h, i) => redPurpleCell(ws.getCell(1, PITCH_COL + i), h, fontSize));

  // ── 守備: 事前スキャンで全年合計Innを集計しグローバル順序を決定 ──
  const posTotalInn = {};
  DEF_POSITIONS.forEach(p => { posTotalInn[p.label] = 0; });

  ws.eachRow((row, rn) => {
    if (rn <= 2) return; // Row1=カテゴリ, Row2=サブヘッダーをスキップ
    const ab = Number(row.getCell(5).value) || 0;
    if (!ab) return;
    for (const pos of DEF_POSITIONS) {
      const inn = parseInn(row.getCell(pos.innCol).value);
      if (inn != null) posTotalInn[pos.label] += inn;
    }
  });

  // 合計Inn降順でソート → 守備列のグローバル順序
  const globalOrder = [...DEF_POSITIONS]
    .filter(p => posTotalInn[p.label] > 0)
    .sort((a, b) => posTotalInn[b.label] - posTotalInn[a.label]);

  // 守備セクションのヘッダー書き込み
  // Row 1: "守備" (BH1)、Row 2: ポジション名 (BH2, BI2, ...)
  if (globalOrder.length > 0) {
    purpleCell(ws.getCell(1, DEF_START_COL), '守備', fontSize);
    globalOrder.forEach((pos, i) => {
      purpleCell(ws.getCell(2, DEF_START_COL + i), pos.label, fontSize);
    });
  }

  // ── メインループ: 打撃・投球・守備 ──
  const careerPosRatings = {}; // pos label → { sumWeighted, sumInn } for weighted-average career DRS
  let count = 0;
  ws.eachRow((row, rn) => {
    if (rn === 1) return; // ヘッダースキップ

    const ab = Number(row.getCell(5).value)  || 0; // E: 打数
    if (!ab) return;

    const yr       = String(row.getCell(2).value || '').trim();
    const isCareer = yr === '通算';

    const hr     = Number(row.getCell(10).value) || 0; // J: 本塁打
    const walks  = Number(row.getCell(12).value) || 0; // L: 四球
    const k      = Number(row.getCell(13).value) || 0; // M: 三振
    const sb     = Number(row.getCell(14).value) || 0; // N: 盗塁
    const cs     = Number(row.getCell(15).value) || 0; // O: 盗塁死
    const avg    = parseBA(row.getCell(16).value);     // P: 打率
    const vsL    = parseBA(row.getCell(19).value);     // S: 対左打率
    const clutch = parseBA(row.getCell(20).value);     // T: 得点圏打率
    const spd    = Number(row.getCell(21).value) || 0; // U: 走力

    const tier     = calcHRTier(hr, ab);
    const walkRate = ab > 0 ? walks / ab * 1000 : 0;
    const soRate   = k / ab * 1000;
    const spdVal   = calcSpeed(spd);

    const vals = [
      calcMeet(avg, tier),                        // AS: ミート    (K2)
      calcPower(hr, ab),                          // AT: パワー    (L2)
      spdVal,                                     // AU: スピード
      calcChance(clutch - avg),                   // AV: チャンス  (N2)
      calcEye(walkRate),                          // AW: 選球眼    (O2)
      calcSO(soRate),                             // AX: 三振      (P2)
      tier,                                       // AY: HR        (I2)
      calcStealAbility(spdVal, ab, walks, sb, cs),// AZ: 盗塁能
      calcVsLeft(vsL - avg),                      // BA: 対左投手  (R2)
    ];
    vals.forEach((v, i) => purpleCell(ws.getCell(rn, START_COL + i), v, fontSize));

    // リード (START_COL+9=54) と 阻止率 (START_COL+10=55) — 捕手のみ
    {
      const cInn = parseInn(row.getCell(29).value); // C Inn列
      if (cInn != null && cInn > 0) {
        const yrData  = catcherData?.byYear?.[yr]  ?? null;
        const carData = catcherData?.career         ?? null;

        // リード: pitches 1500換算フレーミングrun
        // 年別→通算 の順でフォールバック（歴代選手も通算値で類推）
        const framingData = yrData?.framing ?? carData?.framing;
        const leadVal = (framingData && framingData.pitches > 0)
          ? Math.round(framingData.runs * 1500 / framingData.pitches)
          : 0; // 取得不能→0
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

    // BD-BJ: 投球能力値 (FB,2C,CT,SL,CB,CH,SF)
    const pitchRaw = [22, 23, 24, 25, 26, 27, 28].map(c => parseBA(row.getCell(c).value));
    const pitchVals = calcPitchRatings(pitchRaw);
    pitchVals.forEach((v, i) => {
      const cell = ws.getCell(rn, PITCH_COL + i);
      if (v !== '') redPurpleCell(cell, v, fontSize);
    });

    // BK+: 守備能力値 (グローバル順序に従い BK列以降に数値書き込み)
    if (globalOrder.length > 0) {
      if (isCareer) {
        // 通算行: 年度別加重平均（Inn加重）で算出した値を書き込む
        globalOrder.forEach((gpos, i) => {
          const data = careerPosRatings[gpos.label];
          if (!data || data.sumInn === 0) return;
          purpleCell(ws.getCell(rn, DEF_START_COL + i), Math.round(data.sumWeighted / data.sumInn), fontSize);
        });
      } else {
        // 年別行: 計算して書き込みつつ加重平均用に累積
        const yearPositions = DEF_POSITIONS
          .map(pos => ({
            label: pos.label,
            inn:   parseInn(row.getCell(pos.innCol).value),
            drs:   parseDRS(row.getCell(pos.drsCol).value),
          }))
          .filter(p => p.inn != null && p.inn > 0)
          .sort((a, b) => b.inn - a.inn);

        if (yearPositions.length > 0) {
          const mainInn = yearPositions[0].inn; // ODS O37
          // メイン守備比率 < 12% → DH専属とみなし全ポジションに -15 修正
          const mainInnPct = mainInn / ((ab + walks) * 2) * 100;
          const dhPenalty  = mainInnPct < 12 ? -15 : 0;

          globalOrder.forEach((gpos, i) => {
            const yp = yearPositions.find(p => p.label === gpos.label);
            if (!yp) return; // このyearにデータなし

            // 個別出場比率 < 2% → 非表示（端的すぎる守備）
            const defInnPct = yp.inn / ((ab + walks) * 2) * 100;
            if (defInnPct < 2) return;

            const isMain = yp === yearPositions[0]; // 最多Innポジション = 主
            const rating = isMain
              ? calcDefMain(yp.inn, yp.drs)
              : calcDefSub(yp.inn, yp.drs, mainInn);

            if (rating != null) {
              const finalRating = Math.round(rating) + dhPenalty;
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

// ── HTML UI ─────────────────────────────────────────────────────────────────
const HTML = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>MLB成績 能力値追加ツール</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: 'Meiryo UI', 'Meiryo', 'Yu Gothic UI', sans-serif;
  background: #f0e8f5;
  min-height: 100vh;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 20px;
}
.card {
  background: white;
  border-radius: 12px;
  box-shadow: 0 4px 20px rgba(112,48,160,0.15);
  padding: 32px;
  width: 100%;
  max-width: 560px;
}
h1 {
  color: #7030A0;
  font-size: 20px;
  margin-bottom: 6px;
  display: flex;
  align-items: center;
  gap: 8px;
}
h1::before { content: "⚾"; font-size: 22px; }
.subtitle {
  color: #888;
  font-size: 12px;
  margin-bottom: 20px;
}
.section { margin-bottom: 20px; }
.section-label {
  font-size: 12px;
  font-weight: bold;
  color: #555;
  margin-bottom: 6px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
}
.cols-badge {
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
  margin-bottom: 20px;
}
.badge {
  background: #f3e5f5;
  color: #7030A0;
  border: 1px solid #ce93d8;
  border-radius: 4px;
  padding: 3px 8px;
  font-size: 11px;
  font-weight: bold;
}
.file-area {
  border: 2px dashed #ce93d8;
  border-radius: 8px;
  padding: 16px;
  min-height: 60px;
  display: flex;
  align-items: center;
  gap: 12px;
  background: #faf5ff;
  margin-bottom: 12px;
  transition: border-color 0.2s;
}
.file-area.selected {
  border-color: #7030A0;
  background: #f3e5f5;
}
.file-icon { font-size: 28px; flex-shrink: 0; }
.file-path {
  font-size: 13px;
  color: #555;
  word-break: break-all;
  flex: 1;
}
.file-path.has-file { color: #4a1470; font-weight: bold; }
.btn-row { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
button {
  padding: 10px 22px;
  border: none;
  border-radius: 6px;
  cursor: pointer;
  font-size: 14px;
  font-family: inherit;
  font-weight: bold;
  transition: all 0.15s;
}
.btn-browse {
  background: #555;
  color: white;
}
.btn-browse:hover { background: #333; }
.btn-browse:active { transform: scale(0.97); }
.btn-confirm {
  background: #7030A0;
  color: white;
  min-width: 120px;
}
.btn-confirm:hover:not(:disabled) { background: #5a1e85; }
.btn-confirm:active:not(:disabled) { transform: scale(0.97); }
.btn-confirm:disabled { background: #ccc; cursor: not-allowed; }
.status-msg {
  font-size: 12px;
  color: #888;
  flex: 1;
  min-width: 100px;
}
.result {
  margin-top: 16px;
  padding: 14px 16px;
  border-radius: 8px;
  font-size: 14px;
  display: none;
}
.result.ok {
  background: #e8f5e9;
  color: #2e7d32;
  border: 1px solid #a5d6a7;
}
.result.err {
  background: #ffebee;
  color: #c62828;
  border: 1px solid #ef9a9a;
}
.spinner {
  display: inline-block;
  width: 14px; height: 14px;
  border: 2px solid #ddd;
  border-top-color: #7030A0;
  border-radius: 50%;
  animation: spin 0.7s linear infinite;
  vertical-align: middle;
  margin-right: 6px;
}
@keyframes spin { to { transform: rotate(360deg); } }
.formula-note {
  font-size: 11px;
  color: #999;
  margin-top: 20px;
  border-top: 1px solid #eee;
  padding-top: 12px;
  line-height: 1.7;
}
</style>
</head>
<body>
<div class="card">
  <h1>MLB選手成績 能力値追加ツール</h1>
  <div class="subtitle">守備.ods の計算式を元に、成績.xlsx へ能力値列を自動追加します</div>

  <div class="section">
    <div class="section-label">追加される列</div>
    <div class="cols-badge">
      <span class="badge">AS ミート</span>
      <span class="badge">AT パワー</span>
      <span class="badge">AU スピード</span>
      <span class="badge">AV チャンス</span>
      <span class="badge">AW 選球眼</span>
      <span class="badge">AX 三振</span>
      <span class="badge">AY HR</span>
      <span class="badge">AZ 盗塁能</span>
      <span class="badge">BA 対左投手</span>
      <span class="badge">BB リード</span>
      <span class="badge">BC 阻止率</span>
      <span class="badge" style="background:#fce4ec;color:#c00060;border-color:#f48fb1">BD-BJ 投球</span>
      <span class="badge">BK+ 守備</span>
    </div>
  </div>

  <div class="section">
    <div class="section-label">ファイル選択</div>
    <div class="file-area" id="fileArea">
      <div class="file-icon">📄</div>
      <div class="file-path" id="filePath">ファイルが選択されていません</div>
    </div>
    <div class="btn-row">
      <button class="btn-browse" onclick="browse()">📂 ファイルを参照...</button>
      <button class="btn-confirm" id="btnConfirm" disabled onclick="confirm_process()">✓ 確 認</button>
      <span class="status-msg" id="status"></span>
    </div>
    <div class="result" id="result"></div>
  </div>

  <div class="formula-note">
    <strong>計算ロジック（守備.ods 参照）</strong><br>
    ミート: 打率 → K2 ／ パワー: 本塁打+打数 → L2 ／ 走力: U≥50はそのまま・未満は(U+100)÷3<br>
    チャンス: 得点圏打率-打率 → N2 ／ 選球眼: 四球÷打数×1000 → O2<br>
    三振: 三振÷(打数-四球)×1000 → P2 ／ HR: 本塁打+打数 → I2 ／ 対左: 対左打率-打率 → R2<br>
    盗塁能: 守備.ods 盗塁能シート線形近似でnetSBper500と照合し最近傍ability値<br>
    リード: Baseball Savant framing runs × 1500 ÷ pitches（捕手出場あり/データなし→0, 出場なし→空白）<br>
    阻止率: MLB Stats API Cポジション CS÷(SB+CS)×100（捕手出場なし→空白）
  </div>
</div>

<script>
let selectedPath = '';

async function browse() {
  setStatus('<span class="spinner"></span>ダイアログを開いています...');
  try {
    const res = await fetch('/api/browse');
    const data = await res.json();
    if (data.path) {
      selectedPath = data.path;
      const fp = document.getElementById('filePath');
      fp.textContent = data.path;
      fp.className = 'file-path has-file';
      document.getElementById('fileArea').className = 'file-area selected';
      document.getElementById('btnConfirm').disabled = false;
      setStatus('');
    } else {
      setStatus('ファイルが選択されませんでした');
    }
  } catch (e) {
    setStatus('エラー: ' + e.message);
  }
}

async function confirm_process() {
  if (!selectedPath) return;
  const btn = document.getElementById('btnConfirm');
  btn.disabled = true;
  setStatus('<span class="spinner"></span>処理中...');
  showResult('', '');

  try {
    const res = await fetch('/api/process', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filePath: selectedPath })
    });
    const data = await res.json();
    if (data.success) {
      showResult('ok', '✓ 完了: ' + data.count + ' 行の能力値を書き込みました（紫セル追加済）');
    } else {
      showResult('err', '✗ エラー: ' + data.error);
    }
  } catch (e) {
    showResult('err', '✗ 通信エラー: ' + e.message);
  }
  btn.disabled = false;
  setStatus('');
}

function setStatus(html) {
  document.getElementById('status').innerHTML = html;
}

function showResult(type, msg) {
  const r = document.getElementById('result');
  if (!type) { r.style.display = 'none'; return; }
  r.className = 'result ' + type;
  r.textContent = msg;
  r.style.display = 'block';
}
</script>
</body>
</html>`;

// ── HTTP Server ──────────────────────────────────────────────────────────────
const server = http.createServer(async (req, res) => {
  const parsedUrl = new URL(req.url, `http://localhost:${PORT}`);
  const pathname  = parsedUrl.pathname;

  if (req.method === 'GET' && pathname === '/') {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(HTML);
    return;
  }

  if (req.method === 'GET' && pathname === '/api/browse') {
    const fp = browseFile();
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ path: fp }));
    return;
  }

  if (req.method === 'POST' && pathname === '/api/process') {
    let body = '';
    req.on('data', chunk => { body += chunk; });
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

  res.writeHead(404);
  res.end('Not found');
});

server.on('error', err => {
  if (err.code === 'EADDRINUSE') {
    console.error(`ポート ${PORT} は既に使用中です。ブラウザで http://localhost:${PORT} を開いてください。`);
  } else {
    console.error('サーバーエラー:', err.message);
  }
  process.exit(1);
});

server.listen(PORT, '127.0.0.1', () => {
  const url = `http://localhost:${PORT}`;
  console.log('');
  console.log('  Baseball  MLB Seiseki Tool');
  console.log('');
  console.log('  URL: ' + url);
  console.log('');
  console.log('  Ctrl+C to stop.');
  console.log('');
  // ブラウザを開く（複数の方法でフォールバック）
  const { spawn } = require('child_process');
  try {
    spawn('cmd.exe', ['/c', 'start', '', url], { detached: true, shell: false, stdio: 'ignore' }).unref();
  } catch (e) {
    console.log('  (ブラウザを手動で開いてください: ' + url + ')');
  }
});
