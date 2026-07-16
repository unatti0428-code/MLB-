'use strict';
const http    = require('http');
const ExcelJS = require('exceljs');
const path    = require('path');
const { spawnSync } = require('child_process');

const PORT = 3942;

function browseFile(filter) {
  const r = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', `
[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
$d = New-Object System.Windows.Forms.OpenFileDialog
$d.Filter = "${filter || 'All Files (*.*)|*.*'}"
if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($d.FileName)
  [Convert]::ToBase64String($bytes)
}`], { encoding: 'buffer' });
  const b64 = (r.stdout || Buffer.alloc(0)).toString('ascii').trim();
  return b64 ? Buffer.from(b64, 'base64').toString('utf8') : '';
}

// ── イニング換算: "200.1" → 200.333..., "200.2" → 200.667 ─────────────────
function parseIP(ipStr) {
  const s = String(ipStr || '').trim();
  if (!s || s === '--') return 0;
  const [whole, frac] = s.split('.');
  return (parseInt(whole) || 0) + (parseInt(frac || 0)) / 3;
}

// ── スタミナ計算式 (守備.ods AC3 の等価実装 + GS補正) ────────────────────
// 守備.ods AC3 の数式:
//   IFERROR(IFS(V3>=230, ROUND(V3/W3*12.5), V3>=210, ROUND(V3/W3*13.1),
//               V3>=86,  ROUND(V3/W3*13.5), V3>=65,  ROUND(V3/W3*20),
//               V3>=50,  ROUND(V3/W3*21),   V3<=49,  ROUND(V3/W3*22)), "")
// V3 = 換算イニング(K列), W3 = 試合数(G列)
//
// GS補正: IP>=65 の区間で H列(GS)/G列(試合数) > 0.5 の場合は
//         係数を 20 → 13 に変更（先発投手寄りのスタミナ算出）
function calcStaminaFromIP(ip, g, gs) {
  if (!ip || isNaN(ip) || ip < 0) return '';
  if (!g  || isNaN(g)  || g  <= 0) return '';
  const ratio = ip / g;
  if (ip >= 230) return Math.round(ratio * 12.5);
  if (ip >= 210) return Math.round(ratio * 13.1);
  if (ip >= 86)  return Math.round(ratio * 13.5);
  if (ip >= 65) {
    // GS/G > 0.5 → 先発寄り: 係数 13、それ以外: 係数 20
    const mult = (gs > 0 && (gs / g) > 0.5) ? 13 : 20;
    return Math.round(ratio * mult);
  }
  if (ip >= 50) return Math.round(ratio * 21);
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
// 守備.ods AE2 の数式:
//   IFERROR(IFS(W2>=8.2,ROUND(55-(W2-8.2)/0.35),W2>=6.6,ROUND(60-(W2-6.6)/0.32),
//               W2>=5.2,ROUND(65-(W2-5.2)/0.28),W2>=4,ROUND(70-(W2-4)/0.24),
//               W2>=3.2,ROUND(75-(W2-3.2)/0.16),W2>=2.5,ROUND(80-(W2-2.5)/0.14),
//               W2>=1.9,ROUND(85-(W2-1.9)/0.12),W2>=1.4,ROUND(90-(W2-1.4)/0.1),
//               W2>=1,ROUND(95-(W2-1)/0.08),W2>=0,ROUND(100-(W2-0.7)/0.06)),"")
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
  return Math.round(100 - (era - 0.7) / 0.06);  // era >= 0
}

// ── 奪三振計算式 (守備.ods AF2 の等価実装) ────────────────────────────────────
// 守備.ods AF2 の数式:
//   IFERROR(IFS(X2<=6,ROUND(40+(X2-6)/0.2),X2<=10,ROUND(80+(X2-10)/0.1),
//               X2<=30,ROUND(100+(X2-14)/0.2)),"")
// X2 = 奪三振(P列) / 換算イニング(K列) × 9  (= K/9)
function calcSanshinFromK9(k9) {
  if (k9 == null || isNaN(k9) || k9 < 0) return '';
  if (k9 <= 6)  return Math.round(40 + (k9 - 6)  / 0.2);
  if (k9 <= 10) return Math.round(80 + (k9 - 10) / 0.1);
  if (k9 <= 30) return Math.round(100 + (k9 - 14) / 0.2);
  return '';
}

// ── 重さ計算式 (守備.ods AG2 の等価実装) ─────────────────────────────────────
// 守備.ods AG2 の数式:
//   IFERROR(IFS(Y2>=2.2,ROUND(50-(Y2-2.2)/0.1),Y2>=1.8,ROUND(55-(Y2-1.8)/0.08),
//               Y2>=1.5,ROUND(60-(Y2-1.5)/0.06),Y2>=1.3,ROUND(65-(Y2-1.3)/0.04),
//               Y2>=1,ROUND(80-(Y2-1)/0.02),Y2>=0.25,ROUND(105-(Y2-0.25)/0.03),
//               Y2>=0.1,ROUND(110-(Y2-0.1)/0.03)),"")
// Y2 = 被本塁打(N列) / 換算イニング(K列) × 9  (= HR/9)
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
// 守備.ods AH2 の数式:
//   IFERROR(IFS(Z2<-60,ROUND(-15+(60+Z2)/8),Z2>60,ROUND(15+(Z2-60)/8),
//               Z2>=-60,ROUND(Z2/4),Z2<=60,ROUND(Z2/4)),"")
// Z2 = 被打率(Q列) - 対左被打率(S列)  ※どちらも整数形式（.275→275）
function calcTaiHidariFromDiff(z) {
  if (z == null || isNaN(z)) return '';
  if (z < -60) return Math.round(-15 + (60 + z) / 8);
  if (z > 60)  return Math.round(15 + (z - 60) / 8);
  return Math.round(z / 4);
}

// ── 対盗塁計算式 (守備.ods AI2=AJ2+AK2 の等価実装) ──────────────────────────
// AJ2: IFERROR(IFS(AA2/AB2*9>=1,ROUND(-7),AA2/AB2*9>=0,ROUND(11-(AA2/AB2*9*18))),"")
// AK2: IFERROR(IFS((AA2-AA3)/(AA2+AB3)>=0.85,-10,(AA2-AA3)/(AA2+AB3)<=0.35,18,
//              (AA2-AA3)/(AA2+AB3)<0.85,ROUND((0.65-(AA2-AA3)/(AA2+AB3))*60)),"")
// AI2: IFERROR(AJ2+AK2,"")
// AA2=SB(T列), AA3=PK(U列), AB2=換算IP(K列), AB3=CS(V列)
function calcTaiTouruiFromSBData(sb, pk, ip, cs) {
  if (!ip || ip <= 0) return '';
  // AJ2: SB per 9 innings
  const sb9 = (sb / ip) * 9;
  let aj;
  if (sb9 >= 1)      aj = -7;
  else if (sb9 >= 0) aj = Math.round(11 - sb9 * 18);
  else               return '';
  // AK2: (SB-PK)/(SB+CS) ratio
  const denom = sb + cs;
  if (denom <= 0) return '';   // 分母0 → IFERROR で "" と同等
  const ratio = (sb - pk) / denom;
  let ak;
  if (ratio >= 0.85)   ak = -10;
  else if (ratio <= 0.35) ak = 18;
  else ak = Math.round((0.65 - ratio) * 60);
  return aj + ak;
}

// ── セルスタイル ─────────────────────────────────────────────────────────────
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
  { idx: 0, name: 'フォーシーム',   startCol: 23 }, // W-Z
  { idx: 1, name: 'スライダー',     startCol: 27 }, // AA-AD
  { idx: 2, name: 'チェンジアップ', startCol: 31 }, // AE-AH
  { idx: 3, name: 'カーブ',         startCol: 35 }, // AI-AL
  { idx: 4, name: 'カットボール',   startCol: 39 }, // AM-AP
  { idx: 5, name: 'ツーシーム',     startCol: 43 }, // AQ-AT
  { idx: 6, name: 'スプリット',     startCol: 47 }, // AU-AX
];
const PITCH_ABILITY_START_COL = 59; // BG〜

// 球速能力 = S: IFERROR(IFS(O<10,""),IFS(O>11,ROUND(O*1.6+4),""))
function calcKyuSoku(velo) {
  if (velo == null || isNaN(velo) || velo < 10) return '';
  if (velo > 11) return Math.round(velo * 1.6 + 4);
  return '';
}

// AH: 被打率ベーススコア（球種別・守備.ods AH14〜AH20 等価実装）
function calcAH_pitch(idx, p) {
  if (p == null || isNaN(p)) return '';
  switch (idx) {
    case 0: // フォーシーム (AH14)
      if (p >= 300) return Math.round(55  - (p-300)/4);
      if (p >= 250) return Math.round(80  - (p-250)/2);
      if (p >= 235) return Math.round(85  - (p-235)/3);
      if (p >= 215) return Math.round(90  - (p-215)/4);
      if (p >= 190) return Math.round(95  - (p-190)/5);
      if (p >= 150) return Math.round(100 - (p-150)/8);
      if (p >= 1)   return Math.round(105 - (p-70)/16);
      return '';
    case 1: // スライダー (AH15)
      if (p >= 300) return Math.round(55  - (p-300)/6);
      if (p >= 200) return Math.round(80  - (p-200)/4);
      if (p >= 150) return Math.round(90  - (p-150)/5);
      if (p >= 120) return Math.round(95  - (p-120)/6);
      if (p >= 1)   return Math.round(105 - (p-60)/6);
      return '';
    case 2: // チェンジアップ (AH16)
    case 3: // カーブ (AH17)
      if (p >= 300) return Math.round(55  - (p-300)/6);
      if (p >= 200) return Math.round(80  - (p-200)/4);
      if (p >= 150) return Math.round(90  - (p-150)/5);
      if (p >= 120) return Math.round(95  - (p-120)/6);
      if (p >= 80)  return Math.round(100 - (p-80)/8);
      if (p >= 1)   return Math.round(100 - (p-80)/12);
      return '';
    case 4: // カットボール (AH18)
      if (p >= 290) return Math.round(55  - (p-290)/4);
      if (p >= 240) return Math.round(80  - (p-240)/2);
      if (p >= 225) return Math.round(85  - (p-225)/3);
      if (p >= 205) return Math.round(90  - (p-205)/4);
      if (p >= 180) return Math.round(95  - (p-180)/5);
      if (p >= 1)   return Math.round(105 - (p-80)/10);
      return '';
    case 5: // ツーシーム (AH19)
      if (p >= 330) return Math.round(50  - (p-330)/6);
      if (p >= 310) return Math.round(55  - (p-310)/4);
      if (p >= 260) return Math.round(80  - (p-260)/2);
      if (p >= 245) return Math.round(85  - (p-245)/3);
      if (p >= 220) return Math.round(90  - (p-220)/5);
      if (p >= 195) return Math.round(95  - (p-195)/5);
      if (p >= 150) return Math.round(100 - (p-150)/9);
      if (p >= 1)   return Math.round(105 - (p-70)/16);
      return '';
    case 6: // スプリット (AH20)
      if (p >= 285) return Math.round(55  - (p-285)/5);
      if (p >= 245) return Math.round(65  - (p-245)/4);
      if (p >= 200) return Math.round(80  - (p-200)/3);
      if (p >= 110) return Math.round(95  - (p-110)/6);
      if (p >= 1)   return Math.round(105 - (p-30)/8);
      return '';
    default: return '';
  }
}

// AI: SLGベーススコア（球種別・守備.ods AI14〜AI20 等価実装）
function calcAI_pitch(idx, q) {
  if (q == null || isNaN(q)) return '';
  switch (idx) {
    case 0: // フォーシーム (AI14)
    case 4: // カットボール (AI18)
    case 5: // ツーシーム (AI19)
      if (q >= 570) return Math.round(60  - (q-570)/10);
      if (q >= 500) return Math.round(70  - (q-500)/7);
      if (q >= 300) return Math.round(90  - (q-300)/10);
      if (q >= 100) return Math.round(100 - (q-100)/20);
      if (q >= 1)   return Math.round(105 - (q-50)/10);
      return '';
    case 1: // スライダー (AI15)
    case 2: // チェンジアップ (AI16)
    case 3: // カーブ (AI17)
      if (q >= 470) return Math.round(65  - (q-470)/8);
      if (q >= 420) return Math.round(70  - (q-420)/10);
      if (q >= 300) return Math.round(80  - (q-300)/12);
      if (q >= 250) return Math.round(90  - (q-250)/5);
      if (q >= 215) return Math.round(95  - (q-215)/7);
      if (q >= 180) return Math.round(100 - (q-180)/7);
      if (q >= 100) return Math.round(105 - (q-100)/16);
      if (q >= 1)   return Math.round(110 - (q-50)/10);
      return '';
    case 6: // スプリット (AI20)
      if (q >= 310) return Math.round(80  - (q-310)/6);
      if (q >= 270) return Math.round(85  - (q-270)/8);
      if (q >= 1)   return Math.round(102 - (q-32)/14);
      return '';
    default: return '';
  }
}

// AK: 投球率ボーナス（守備.ods AK14〜AK20 等価実装）
// TypeA (idx 0,4,5): 最大ボーナス8  TypeB (idx 1,2,3,6): 最大ボーナス18
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

// T = 球威: IFERROR(IFS(AND(R<=8,AJ+AK>=85),85,AJ,ROUNDUP(AJ+AK)),"")
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
function calcKankyuu(eraStr, seikyu, pitchData) {
  const v1 = _kERA(parseFloat(String(eraStr || '').trim()));
  if (v1 === null) return '';
  const v2 = _kSeikyu(seikyu !== '' && seikyu != null ? Number(seikyu) : NaN);
  if (v2 === null) return '';
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

// 列定義
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

// ── メイン処理 ───────────────────────────────────────────────────────────────
async function processFile(xlsxPath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(xlsxPath);
  const ws = wb.worksheets[0];
  const fontSize = ws.getCell(1, 1).font?.size || 11;

  // ヘッダー書き込み（Row 1）
  purpleCell(ws.getCell(1, STAMINA_COL),   'スタミナ', fontSize);
  purpleCell(ws.getCell(1, SEIKYU_COL),    '制球',     fontSize);
  purpleCell(ws.getCell(1, KANKYUU_COL),   '緩急',     fontSize);
  purpleCell(ws.getCell(1, SEISIN_COL),    '精神',     fontSize);
  purpleCell(ws.getCell(1, SANSHIN_COL),   '奪三振',   fontSize);
  purpleCell(ws.getCell(1, OMOSA_COL),     '重さ',     fontSize);
  purpleCell(ws.getCell(1, TAILEFT_COL),   '対左',     fontSize);
  purpleCell(ws.getCell(1, TAITOURUI_COL), '対盗塁',   fontSize);

  // データ行収集 + 球種アクティブ判定
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
    // 球種データ収集
    const pitchData = {};
    for (const pg of PITCH_GROUPS) {
      const velo = row.getCell(pg.startCol + 0).value;
      const ba   = row.getCell(pg.startCol + 1).value;
      const slg  = row.getCell(pg.startCol + 2).value;
      const pct  = row.getCell(pg.startCol + 3).value;
      const pctStr = String(pct == null ? '' : pct).trim();
      const pctNum = Number(pctStr);
      if (pctStr && pctStr !== '--' && !isNaN(pctNum) && pctNum > 0) {
        pitchActiveSet.add(pg.idx);
      }
      pitchData[pg.idx] = { velo, ba, slg, pct };
    }
    dataRows.push({ rn, ipRaw, g, gs, bb, eraRaw, hr, so, avgRaw, vsLRaw, sb, pk, cs, pitchData });
  });

  // アクティブ球種リスト（元の順序を保持）
  const activePitchList = PITCH_GROUPS.filter(pg => pitchActiveSet.has(pg.idx));

  // 球種能力 ヘッダー行書き込み
  activePitchList.forEach((pg, i) => {
    const base = PITCH_ABILITY_START_COL + i * 3;
    redPurpleCell(ws.getCell(1, base), pg.name, fontSize);
    try { ws.mergeCells(1, base, 1, base + 2); } catch {}
    redPurpleCell(ws.getCell(2, base + 0), '球速', fontSize);
    redPurpleCell(ws.getCell(2, base + 1), '球威', fontSize);
    redPurpleCell(ws.getCell(2, base + 2), '割合', fontSize);
  });

  let count = 0;
  for (const { rn, ipRaw, g, gs, bb, eraRaw, hr, so, avgRaw, vsLRaw, sb, pk, cs, pitchData } of dataRows) {
    const ip = parseIP(ipRaw);
    if (!ip) continue;

    // スタミナ
    const stamina = calcStaminaFromIP(ip, g, gs);
    if (stamina !== '') purpleCell(ws.getCell(rn, STAMINA_COL), stamina, fontSize);

    // 制球
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

    // 奪三振
    const sanshin = calcSanshinFromK9(so / ip * 9);
    if (sanshin !== '') purpleCell(ws.getCell(rn, SANSHIN_COL), sanshin, fontSize);

    // 重さ
    const omosa = calcOmosaFromHR9(hr / ip * 9);
    if (omosa !== '') purpleCell(ws.getCell(rn, OMOSA_COL), omosa, fontSize);

    // 対左
    const avgStr = String(avgRaw || '').trim();
    const vsLStr = String(vsLRaw || '').trim();
    if (avgStr && avgStr !== '--' && vsLStr && vsLStr !== '--') {
      const taileft = calcTaiHidariFromDiff(Number(avgStr) - Number(vsLStr));
      if (taileft !== '') purpleCell(ws.getCell(rn, TAILEFT_COL), taileft, fontSize);
    }

    // 対盗塁
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

      // 球速
      const kyusoku = calcKyuSoku(veloNum);
      if (kyusoku !== '') redPurpleCell(ws.getCell(rn, base + 0), kyusoku, fontSize);

      // 球威 (AH→AI→AJ→AK→T)
      const ah = calcAH_pitch(pg.idx, baNum);
      const ai = calcAI_pitch(pg.idx, slgNum);
      const aj = (ah !== '' && ai !== '') ? (Number(ah) + Number(ai)) / 2 : '';
      const ak = calcAK_pitch(pg.idx, aj, pctNum);
      const kyui = calcKyuI(aj, ak, pctNum);
      if (kyui !== '') redPurpleCell(ws.getCell(rn, base + 1), kyui, fontSize);

      // 割合（直接転記）
      redPurpleCell(ws.getCell(rn, base + 2), pctNum, fontSize);
    });

    count++;
  }

  await wb.xlsx.writeFile(xlsxPath);
  return count;
}

// ── HTML ──────────────────────────────────────────────────────────────────────
const HTML = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>MLB投手成績 能力値追加ツール</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Meiryo UI','Meiryo','Yu Gothic UI',sans-serif;background:#f0e8f5;
  min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}
.card{background:white;border-radius:12px;box-shadow:0 4px 20px rgba(112,48,160,.15);
  padding:32px;width:100%;max-width:560px}
h1{color:#7030A0;font-size:20px;margin-bottom:6px;display:flex;align-items:center;gap:8px}
h1::before{content:"⚾";font-size:22px}
.subtitle{color:#888;font-size:12px;margin-bottom:22px}
.sec{margin-bottom:20px}
.sec-label{font-size:12px;font-weight:bold;color:#555;margin-bottom:6px;text-transform:uppercase;letter-spacing:.5px}
.file-area{border:2px dashed #ce93d8;border-radius:8px;padding:14px;min-height:56px;
  display:flex;align-items:center;gap:12px;background:#faf5ff;margin-bottom:10px;transition:border-color .2s}
.file-area.sel{border-color:#7030A0;background:#f3e5f5}
.file-icon{font-size:26px;flex-shrink:0}
.file-path{font-size:13px;color:#888;flex:1;word-break:break-all}
.file-path.has{color:#4a1470;font-weight:bold}
.btn-row{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
button{padding:9px 20px;border:none;border-radius:6px;cursor:pointer;
  font-size:13px;font-family:inherit;font-weight:bold;transition:all .15s}
.btn-browse{background:#555;color:white}
.btn-browse:hover{background:#333}
.btn-run{background:#7030A0;color:white;min-width:120px}
.btn-run:hover:not(:disabled){background:#5a1e85}
.btn-run:disabled{background:#ccc;cursor:not-allowed}
.status{font-size:12px;color:#888;flex:1}
.result{margin-top:14px;padding:13px 16px;border-radius:8px;font-size:14px;display:none}
.result.ok{background:#e8f5e9;color:#2e7d32;border:1px solid #a5d6a7}
.result.err{background:#ffebee;color:#c62828;border:1px solid #ef9a9a}
.sp{display:inline-block;width:12px;height:12px;border:2px solid #ddd;border-top-color:#7030A0;
  border-radius:50%;animation:spin .7s linear infinite;vertical-align:middle;margin-right:6px}
@keyframes spin{to{transform:rotate(360deg)}}
.note{font-size:11px;color:#aaa;margin-top:18px;border-top:1px solid #eee;padding-top:12px;line-height:1.8}
.badge{background:#f3e5f5;color:#7030A0;border:1px solid #ce93d8;border-radius:4px;
  padding:3px 8px;font-size:11px;font-weight:bold;display:inline-block;margin:2px}
.formula-box{background:#f9f5ff;border:1px solid #e0c8f0;border-radius:6px;
  padding:10px 14px;font-size:11px;color:#555;line-height:1.9;margin-top:10px}
.formula-box code{font-family:monospace;color:#7030A0;font-size:11px}
.formula-box .col-label{font-weight:bold;color:#333;min-width:40px;display:inline-block}
</style>
</head>
<body>
<div class="card">
  <h1>MLB投手成績 能力値追加ツール</h1>
  <div class="subtitle">投手成績.xlsx に <strong>スタミナ・制球・精神・奪三振・重さ・対左・対盗塁</strong> を自動追加します</div>

  <div class="sec">
    <div class="sec-label">追加される列</div>
    <span class="badge">AY スタミナ</span>
    <span class="badge">AZ 制球</span>
    <span class="badge">BA 精神</span>
    <span class="badge">BB 奪三振</span>
    <span class="badge">BC 重さ</span>
    <span class="badge">BD 対左</span>
    <span class="badge">BE 対盗塁</span>
    <div class="formula-box">
      <span class="col-label">AY</span>スタミナ = AC3相当 (IP/G×係数、GS/G&gt;0.5時IP≥65は×13)<br>
      <span class="col-label">AZ</span>制球 = AC2相当 (BB9=四死球/IP×9)<br>
      <span class="col-label">BA</span>精神 = AE2相当 (W2=防御率(F列))<br>
      <span class="col-label">BB</span>奪三振 = AF2相当 (X2=K9=奪三振(P列)/IP×9)<br>
      <span class="col-label">BC</span>重さ = AG2相当 (Y2=HR9=被本塁打(N列)/IP×9)<br>
      <span class="col-label">BD</span>対左 = AH2相当 (Z2=被打率(Q列)−対左被打率(S列))<br>
      <span class="col-label">BE</span>対盗塁 = AI2=AJ2+AK2相当 (SB/IP×9 + CS率)
    </div>
  </div>

  <div class="sec">
    <div class="sec-label">投手成績 Excel ファイル</div>
    <div class="file-area" id="fa1">
      <div class="file-icon">📊</div>
      <div class="file-path" id="fp1">ファイルが選択されていません</div>
    </div>
    <div class="btn-row">
      <button class="btn-browse" onclick="browse()">📂 ファイルを参照...</button>
      <button class="btn-run" id="btnRun" disabled onclick="run()">▶ 能力値を追加</button>
      <span class="status" id="status"></span>
    </div>
    <div class="result" id="result"></div>
  </div>

  <div class="note">
    <strong>イニング換算</strong>: X.1 → X＋1/3（≈X.33）、X.2 → X＋2/3（≈X.67）<br>
    <strong>入力列</strong>: F=防御率、G=試合数、H=GS、K=イニング、N=被本塁打、O=四死球、P=奪三振、Q=被打率、S=対左被打率、T=SB、U=PK、V=CS<br>
    <strong>出力列</strong>: AY=スタミナ、AZ=制球、BA=精神、BB=奪三振、BC=重さ、BD=対左、BE=対盗塁（紫塗・白文字・ボールド）
  </div>
</div>

<script>
let selectedPath = '';

async function browse() {
  setStatus('<span class="sp"></span>ダイアログを開いています...');
  try {
    const r = await fetch('/api/browse');
    const d = await r.json();
    if (d.path) {
      selectedPath = d.path;
      const fp = document.getElementById('fp1');
      fp.textContent = d.path;
      fp.className = 'file-path has';
      document.getElementById('fa1').className = 'file-area sel';
      document.getElementById('btnRun').disabled = false;
    }
    setStatus('');
  } catch(e) { setStatus('エラー: ' + e.message); }
}

async function run() {
  if (!selectedPath) return;
  const btn = document.getElementById('btnRun');
  btn.disabled = true;
  setStatus('<span class="sp"></span>処理中...');
  showResult('', '');
  try {
    const r = await fetch('/api/process', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ xlsxPath: selectedPath })
    });
    const d = await r.json();
    if (d.success) showResult('ok', '✓ 完了: ' + d.count + ' 行にスタミナ・制球を書き込みました');
    else showResult('err', '✗ エラー: ' + d.error);
  } catch(e) { showResult('err', '✗ 通信エラー: ' + e.message); }
  btn.disabled = false;
  setStatus('');
}

function setStatus(html) { document.getElementById('status').innerHTML = html; }
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

// ── HTTP Server ───────────────────────────────────────────────────────────────
const server = http.createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');

  if (req.method === 'GET' && url.pathname === '/') {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    return res.end(HTML);
  }
  if (req.method === 'GET' && url.pathname === '/api/browse') {
    const fp = browseFile('Excel Files (*.xlsx)|*.xlsx');
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    return res.end(JSON.stringify({ path: fp }));
  }
  if (req.method === 'POST' && url.pathname === '/api/process') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', async () => {
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
      try {
        const { xlsxPath } = JSON.parse(body);
        if (!xlsxPath) throw new Error('ファイルパスが指定されていません');
        const count = await processFile(xlsxPath);
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
    console.error(`ポート ${PORT} は既に使用中です。http://localhost:${PORT} をブラウザで開いてください。`);
  } else {
    console.error('サーバーエラー:', err.message);
  }
  process.exit(1);
});

server.listen(PORT, '127.0.0.1', () => {
  const url = `http://localhost:${PORT}`;
  console.log('\n  ⚾  MLB投手成績 能力値追加ツール\n\n  URL: ' + url + '\n  Ctrl+C で停止\n');
  const { spawn } = require('child_process');
  try { spawn('cmd.exe', ['/c', 'start', '', url], { detached: true, shell: false, stdio: 'ignore' }).unref(); } catch {}
});
