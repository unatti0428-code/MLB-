const XLSX = require('xlsx');

// ========== HELPERS ==========
// Remove leading period from rate strings: ".262" -> "262", "--" stays "--"
function noLeadPeriod(val) {
  if (typeof val === 'string' && val.startsWith('.')) return val.slice(1);
  return val;
}

// PA-weighted batting average (returns "NNN" string without period)
function weightedBA(entries) { // [{ba: '.xxx', pa: N}, ...]
  let sumH = 0, sumPA = 0;
  for (const e of entries) {
    if (!e || e.ba === '--' || !e.pa) continue;
    sumH  += parseFloat(e.ba) * e.pa;
    sumPA += e.pa;
  }
  if (sumPA === 0) return '--';
  return (sumH / sumPA).toFixed(3).split('.')[1]; // e.g. "276"
}

// Baseball innings addition: addInnings('306.2', '1275.2') -> "1582.1"
function addInnings(a, b) {
  const parse = s => {
    const [full, frac] = String(s).split('.');
    return parseInt(full) * 3 + (parseInt(frac || 0));
  };
  const total = parse(a) + parse(b);
  return Math.floor(total / 3) + '.' + (total % 3);
}
function innToOuts(s) {
  const [f, r] = String(s).split('.');
  return parseInt(f) * 3 + parseInt(r || 0);
}

// ========== SOURCE DATA ==========

// Basic stats (MLB.com)
const years = ['2024', '2025', '2026'];
const basic = {
  '2024': { team:'SF', g:37,  pa:158, r:15, h:38,  d:4,  t:0,  hr:2, rbi:8,  bb:10, so:13, sb:2,  cs:3, avg:'.262', obp:'.310', ops:'.641' },
  '2025': { team:'SF', g:150, pa:617, r:73, h:149, d:31, t:12, hr:8, rbi:55, bb:47, so:71, sb:10, cs:3, avg:'.266', obp:'.327', ops:'.734' },
  '2026': { team:'SF', g:15,  pa:57,  r:4,  h:10,  d:4,  t:0,  hr:1, rbi:7,  bb:5,  so:10, sb:0,  cs:0, avg:'.200', obp:'.263', ops:'.603' },
  '通算': { team:'SF', g:202, pa:832, r:92, h:197, d:39, t:12, hr:11,rbi:70, bb:62, so:94, sb:12, cs:6, avg:'.261', obp:'.320', ops:'.708' },
};

// Splits (MLB Stats API): vs Left AB/H, RISP AB/H
const splitsRaw = {
  '2024': { vsLAB:44,  vsLH:10, rispAB:26,  rispH:6  },
  '2025': { vsLAB:158, vsLH:38, rispAB:117, rispH:30 },
  '2026': { vsLAB:13,  vsLH:2,  rispAB:14,  rispH:4  },
};
const splits = {};
for (const yr of years) {
  const d = splitsRaw[yr];
  splits[yr] = {
    vsLeft: (d.vsLH / d.vsLAB).toFixed(3).split('.')[1],
    risp:   (d.rispH / d.rispAB).toFixed(3).split('.')[1],
  };
}
// Career splits
const totVsLAB = Object.values(splitsRaw).reduce((s,d)=>s+d.vsLAB,0);
const totVsLH  = Object.values(splitsRaw).reduce((s,d)=>s+d.vsLH,0);
const totRispAB= Object.values(splitsRaw).reduce((s,d)=>s+d.rispAB,0);
const totRispH = Object.values(splitsRaw).reduce((s,d)=>s+d.rispH,0);
splits['通算'] = {
  vsLeft: (totVsLH / totVsLAB).toFixed(3).split('.')[1],
  risp:   (totRispH / totRispAB).toFixed(3).split('.')[1],
};

// Sprint Speed percentile (Baseball Savant bar chart: percent_speed_order)
const sprintSpeed = { '2024': 79, '2025': 70, '2026': 52 };
// Career game-weighted average
const sprintCareer = Math.round(
  (79*basic['2024'].g + 70*basic['2025'].g + 52*basic['2026'].g) /
  (basic['2024'].g + basic['2025'].g + basic['2026'].g)
);
sprintSpeed['通算'] = sprintCareer;

// Pitch type BA from Baseball Savant
// Slider = SL + ST (Sweeper) combined, PA-weighted
const rawPitch = {
  '2024': {
    ff: {ba:'.264',pa:60}, si: {ba:'.208',pa:26}, ch: {ba:'.294',pa:18},
    sl: {ba:'.313',pa:18}, st: {ba:'.250',pa:13},   // sl+st = slider
    cu: {ba:'.250',pa:8},  fc: {ba:'.231',pa:13}, fs: {ba:'--', pa:0},
  },
  '2025': {
    ff: {ba:'.284',pa:208},si: {ba:'.316',pa:86}, ch: {ba:'.261',pa:73},
    sl: {ba:'.197',pa:68}, st: {ba:'.244',pa:44},
    cu: {ba:'.207',pa:62}, fc: {ba:'.351',pa:42}, fs: {ba:'.240',pa:26},
  },
  '2026': {
    ff: {ba:'.214',pa:16}, si: {ba:'.083',pa:15}, ch: {ba:'.000',pa:6},
    sl: {ba:'.500',pa:3},  st: {ba:'.500',pa:4},
    cu: {ba:'.200',pa:6},  fc: {ba:'.000',pa:3},  fs: {ba:'--', pa:0},
  },
};

// Build combined slider (SL+ST) per year
const pitchBA = {};
for (const yr of years) {
  const d = rawPitch[yr];
  pitchBA[yr] = {
    ff: d.ff.ba.slice(1) || '--',
    si: d.si.ba.slice(1) || '--',
    ch: d.ch.ba === '--' ? '--' : d.ch.ba.slice(1),
    sl: weightedBA([d.sl, d.st]),
    cu: d.cu.ba.slice(1) || '--',
    fc: d.fc.ba === '--' ? '--' : d.fc.ba.slice(1),
    fs: d.fs.ba === '--' ? '--' : d.fs.ba.slice(1),
  };
}
// Career pitch type BA
pitchBA['通算'] = {
  ff: weightedBA(years.map(yr => rawPitch[yr].ff)),
  si: weightedBA(years.map(yr => rawPitch[yr].si)),
  ch: weightedBA(years.map(yr => rawPitch[yr].ch)),
  sl: weightedBA(years.flatMap(yr => [rawPitch[yr].sl, rawPitch[yr].st])),
  cu: weightedBA(years.map(yr => rawPitch[yr].cu)),
  fc: weightedBA(years.map(yr => rawPitch[yr].fc)),
  fs: weightedBA(years.map(yr => rawPitch[yr].fs)),
};

// Fielding by position (FanGraphs)
// Lee: 2024 CF, 2025 CF, 2026 RF
const positions = ['C','1B','2B','3B','SS','LF','CF','RF'];
const fieldingRaw = {
  '2024': { CF: { inn:'306.2', drs:-2  } },
  '2025': { CF: { inn:'1275.2',drs:-18 } },
  '2026': { RF: { inn:'127.0', drs:-1  } },
};
// Career fielding aggregation (Inn-weighted average DRS)
const fieldingCareer = {};
for (const pos of positions) {
  const entries = years.map(yr => fieldingRaw[yr]?.[pos]).filter(Boolean);
  if (!entries.length) continue;
  const totalOuts = entries.reduce((s, e) => s + innToOuts(e.inn), 0);
  const weightedDRS = totalOuts === 0 ? 0
    : entries.reduce((s, e) => s + e.drs * innToOuts(e.inn), 0) / totalOuts;
  let careerInn = entries[0].inn;
  for (let i = 1; i < entries.length; i++) careerInn = addInnings(careerInn, entries[i].inn);
  fieldingCareer[pos] = {
    inn: careerInn,
    drs: Math.round(weightedDRS),
  };
}

function getFieldVal(yearKey, pos, field) {
  const src = yearKey === '通算' ? fieldingCareer : (fieldingRaw[yearKey] || {});
  return src[pos]?.[field] ?? '--';
}

// ========== BUILD WORKSHEET DATA ==========

// Two header rows
const statsColsRow0 = [
  '選手名','年度','チーム','試合','打数','得点','安打','二塁打','三塁打','本塁打',
  '打点','四球','三振','盗塁','盗塁死',
  '打率','出塁率','OPS',
  '対左打率','得点圏打率',
  '走力',
  '４シーム','シンカー/2シーム','チェンジアップ','スライダー','カーブ','カット','スプリット',
];
const numStatsCols = statsColsRow0.length; // 27

// header row 0: stats cols + position names (each spans 2 cols)
const headerRow0 = [
  ...statsColsRow0,
  ...positions.flatMap(p => [p, '']),
];

// header row 1: blanks for stats cols + Inn/DRS for each position
const headerRow1 = [
  ...new Array(numStatsCols).fill(''),
  ...positions.flatMap(() => ['Inn','DRS']),
];

function buildDataRow(yearKey) {
  const b  = basic[yearKey];
  const sp = splits[yearKey];
  const pt = pitchBA[yearKey];

  const statsVals = [
    'イ・ジョンフ',
    yearKey,
    b.team,
    b.g, b.pa, b.r, b.h, b.d, b.t, b.hr,
    b.rbi, b.bb, b.so, b.sb, b.cs,
    b.avg.slice(1), b.obp.slice(1), b.ops.slice(1),
    sp.vsLeft, sp.risp,
    sprintSpeed[yearKey],
    pt.ff, pt.si, pt.ch, pt.sl, pt.cu, pt.fc, pt.fs,
  ];

  const fieldVals = positions.flatMap(pos => [
    getFieldVal(yearKey, pos, 'inn'),
    getFieldVal(yearKey, pos, 'drs'),
  ]);

  return [...statsVals, ...fieldVals];
}

const allRows = [
  headerRow0,
  headerRow1,
  ...years.map(yr => buildDataRow(yr)),
  buildDataRow('通算'),
];

// ========== CREATE WORKBOOK ==========
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet(allRows);

// Merged cells
const merges = [];

// Merge rows 0-1 for each stats column (vertical merge)
for (let c = 0; c < numStatsCols; c++) {
  merges.push({ s:{r:0,c}, e:{r:1,c} });
}

// Merge cols horizontally for each position (row 0 only)
positions.forEach((_, i) => {
  const c = numStatsCols + i * 2;
  merges.push({ s:{r:0,c}, e:{r:0,c:c+1} });
});

ws['!merges'] = merges;

// Column widths
const colWidths = [
  {wch:12},
  {wch:6},{wch:6},{wch:5},{wch:5},{wch:5},{wch:5},{wch:7},{wch:7},{wch:7},
  {wch:5},{wch:5},{wch:5},{wch:5},{wch:7},
  {wch:6},{wch:6},{wch:6},
  {wch:8},{wch:8},
  {wch:6},
  {wch:9},{wch:13},{wch:12},{wch:9},{wch:7},{wch:8},{wch:9},
  // 8 positions × 2 cols
  ...Array(16).fill({wch:7}),
];
ws['!cols'] = colWidths;

// Freeze first 2 header rows + first 2 columns (選手名, 年度)
ws['!freeze'] = { xSplit: 2, ySplit: 2 };

XLSX.utils.book_append_sheet(wb, ws, 'イ・ジョンフ成績');

// Notes sheet
const noteData = [
  ['データソース','内容'],
  ['MLB.com (MLB Stats API)','基本成績・対左打率・得点圏打率 (2024-2026)'],
  ['Baseball Savant','球種別打率・Sprint Speed (2024-2026)'],
  ['FanGraphs','守備成績 Pos/Inn/DRS (2024-2026)'],
  ['',''],
  ['備考',''],
  ['打率表記','頭の.を除去 (例: .262 → 262)'],
  ['スライダー','SL(従来型)+ST(スイーパー) PA加重平均'],
  ['球種別打率通算','PA加重平均による近似値'],
  ['走力','Baseball Savant棒グラフのパーセンタイル値(percent_speed_order)。通算は試合数加重平均'],
  ['守備','各ポジションのInn(守備イニング)とDRS(守備貢献値)'],
  ['CF通算守備','2024+2025合算: Inn=' + fieldingCareer['CF'].inn + ' DRS=' + fieldingCareer['CF'].drs + '（Inn加重平均）'],
  ['2026守備','RF出場のみ: Inn=' + fieldingCareer['RF'].inn + ' DRS=' + fieldingCareer['RF'].drs],
  ['2026','2026/4/12時点 (シーズン途中)'],
  ['対左打率通算','AB=' + totVsLAB + ' H=' + totVsLH + ' → .' + splits['通算'].vsLeft],
  ['得点圏打率通算','AB=' + totRispAB + ' H=' + totRispH + ' → .' + splits['通算'].risp],
];
const wsNote = XLSX.utils.aoa_to_sheet(noteData);
wsNote['!cols'] = [{wch:25},{wch:65}];
XLSX.utils.book_append_sheet(wb, wsNote, 'データソース・備考');

const outPath = 'イ_ジョンフ_成績.xlsx';
XLSX.writeFile(wb, outPath);
console.log('✓ Created: ' + outPath);
console.log('  Rows: ' + (allRows.length - 2) + ' data rows (+ 2 header rows)');
console.log('  Cols: ' + headerRow0.length);
