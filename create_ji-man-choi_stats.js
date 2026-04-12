const XLSX = require('xlsx');

// ===== ヘルパー関数 =====
function weightedBA(entries) {
  let sumH = 0, sumPA = 0;
  for (const e of entries) {
    if (!e || e.ba === '--' || !e.pa) continue;
    sumH  += parseFloat(e.ba) * e.pa;
    sumPA += e.pa;
  }
  if (sumPA === 0) return '--';
  return (sumH / sumPA).toFixed(3).split('.')[1];
}

function addInnings(list) {
  const total = list.reduce((acc, s) => {
    if (!s || s === '--') return acc;
    const [full, frac] = String(s).split('.');
    return acc + parseInt(full) * 3 + parseInt(frac || 0);
  }, 0);
  return Math.floor(total / 3) + '.' + (total % 3);
}

// ===== 基本成績 (MLB Stats API) =====
const years = ['2016','2017','2018','2019','2020','2021','2022','2023'];

const basic = {
  // pa = atBats（打数）
  '2016': { team:'LAA',  g:54,  pa:112,  r:9,  h:19,  d:4,  t:0, hr:5,  rbi:12, bb:16, so:27,  sb:2, cs:4, avg:'.170', obp:'.271', ops:'.610' },
  '2017': { team:'NYY',  g:6,   pa:15,   r:2,  h:4,   d:1,  t:0, hr:2,  rbi:5,  bb:2,  so:5,   sb:0, cs:0, avg:'.267', obp:'.333', ops:'1.066' },
  '2018': { team:'TB2',  g:61,  pa:190,  r:25, h:50,  d:14, t:1, hr:10, rbi:32, bb:26, so:55,  sb:2, cs:0, avg:'.263', obp:'.357', ops:'.862' },
  '2019': { team:'TB',   g:127, pa:410,  r:54, h:107, d:20, t:2, hr:19, rbi:63, bb:64, so:108, sb:2, cs:3, avg:'.261', obp:'.363', ops:'.822' },
  '2020': { team:'TB',   g:42,  pa:122,  r:16, h:28,  d:13, t:0, hr:3,  rbi:16, bb:20, so:36,  sb:0, cs:0, avg:'.230', obp:'.331', ops:'.741' },
  '2021': { team:'TB',   g:83,  pa:258,  r:36, h:59,  d:14, t:0, hr:11, rbi:45, bb:45, so:87,  sb:0, cs:0, avg:'.229', obp:'.348', ops:'.759' },
  '2022': { team:'TB',   g:113, pa:356,  r:36, h:83,  d:22, t:0, hr:11, rbi:52, bb:58, so:123, sb:0, cs:0, avg:'.233', obp:'.341', ops:'.729' },
  '2023': { team:'PIT2', g:39,  pa:104,  r:12, h:17,  d:5,  t:0, hr:6,  rbi:13, bb:10, so:35,  sb:0, cs:0, avg:'.163', obp:'.239', ops:'.624' },
  '通算': { team:'TB',   g:525, pa:1567, r:190,h:367, d:93, t:3, hr:67, rbi:238,bb:241,so:476, sb:6, cs:7, avg:'.234', obp:'.338', ops:'.764' },
};

// ===== スプリット (MLB Stats API statSplits) =====
const splitsRaw = {
  '2016': { vsLAB:3,  vsLH:0,  rispAB:24, rispH:3  },
  '2017': { vsLAB:2,  vsLH:0,  rispAB:4,  rispH:0  },
  '2018': { vsLAB:18, vsLH:2,  rispAB:46, rispH:7  },
  '2019': { vsLAB:81, vsLH:17, rispAB:92, rispH:23 },
  '2020': { vsLAB:17, vsLH:2,  rispAB:34, rispH:8  },
  '2021': { vsLAB:70, vsLH:13, rispAB:75, rispH:20 },
  '2022': { vsLAB:51, vsLH:15, rispAB:93, rispH:30 },
  '2023': { vsLAB:15, vsLH:2,  rispAB:13, rispH:3  },
};

const splits = {};
for (const yr of years) {
  const d = splitsRaw[yr];
  splits[yr] = {
    vsLeft: d.vsLAB === 0 ? '--' : (d.vsLH / d.vsLAB).toFixed(3).split('.')[1],
    risp:   d.rispAB === 0 ? '--' : (d.rispH / d.rispAB).toFixed(3).split('.')[1],
  };
}
const totVsLAB  = Object.values(splitsRaw).reduce((s,d) => s + d.vsLAB,  0);
const totVsLH   = Object.values(splitsRaw).reduce((s,d) => s + d.vsLH,   0);
const totRispAB = Object.values(splitsRaw).reduce((s,d) => s + d.rispAB, 0);
const totRispH  = Object.values(splitsRaw).reduce((s,d) => s + d.rispH,  0);
splits['通算'] = {
  vsLeft: (totVsLH  / totVsLAB ).toFixed(3).split('.')[1],
  risp:   (totRispH / totRispAB).toFixed(3).split('.')[1],
};

// ===== 走力パーセンタイル (Baseball Savant percent_speed_order) =====
const sprintSpeed = { '2016':31, '2017':22, '2018':27, '2019':20, '2020':7, '2021':21, '2022':8, '2023':6 };
sprintSpeed['通算'] = Math.round(
  years.reduce((sum, yr) => sum + sprintSpeed[yr] * basic[yr].g, 0) /
  years.reduce((sum, yr) => sum + basic[yr].g, 0)
);

// ===== 球種別打率 (Baseball Savant DOMテーブル) =====
// 2016: データなし
const rawPitch = {
  '2016': {
    ff:{ba:'--',pa:0}, si:{ba:'--',pa:0}, ch:{ba:'--',pa:0},
    sl:{ba:'--',pa:0}, st:{ba:'--',pa:0}, cu:{ba:'--',pa:0},
    fc:{ba:'--',pa:0}, fs:{ba:'--',pa:0},
  },
  '2017': {
    ff:{ba:'.500',pa:7},  si:{ba:'.167',pa:6},  ch:{ba:'.333',pa:3},
    sl:{ba:'--', pa:0},   st:{ba:'--', pa:0},   cu:{ba:'--', pa:0},
    fc:{ba:'--', pa:0},   fs:{ba:'.000',pa:2},
  },
  '2018': {
    ff:{ba:'.210',pa:76}, si:{ba:'.471',pa:40}, ch:{ba:'.222',pa:29},
    sl:{ba:'.185',pa:30}, st:{ba:'.500',pa:2},  cu:{ba:'.200',pa:16},
    fc:{ba:'.250',pa:11}, fs:{ba:'.333',pa:13},
  },
  '2019': {
    ff:{ba:'.255',pa:178},si:{ba:'.288',pa:60}, ch:{ba:'.217',pa:78},
    sl:{ba:'.279',pa:78}, st:{ba:'.000',pa:3},  cu:{ba:'.366',pa:44},
    fc:{ba:'.200',pa:29}, fs:{ba:'.286',pa:11},
  },
  '2020': {
    ff:{ba:'.191',pa:56}, si:{ba:'.222',pa:21}, ch:{ba:'.273',pa:23},
    sl:{ba:'.267',pa:18}, st:{ba:'.000',pa:1},  cu:{ba:'.250',pa:6},
    fc:{ba:'.429',pa:9},  fs:{ba:'.125',pa:9},
  },
  '2021': {
    ff:{ba:'.255',pa:117},si:{ba:'.171',pa:46}, ch:{ba:'.179',pa:46},
    sl:{ba:'.237',pa:45}, st:{ba:'.000',pa:4},  cu:{ba:'.273',pa:24},
    fc:{ba:'.278',pa:19}, fs:{ba:'.333',pa:3},
  },
  '2022': {
    ff:{ba:'.194',pa:152},si:{ba:'.366',pa:49}, ch:{ba:'.281',pa:66},
    sl:{ba:'.170',pa:54}, st:{ba:'.286',pa:16}, cu:{ba:'.188',pa:33},
    fc:{ba:'.231',pa:30}, fs:{ba:'.286',pa:15},
  },
  '2023': {
    ff:{ba:'.129',pa:36}, si:{ba:'.286',pa:16}, ch:{ba:'.077',pa:14},
    sl:{ba:'.133',pa:15}, st:{ba:'.000',pa:7},  cu:{ba:'.182',pa:12},
    fc:{ba:'.200',pa:11}, fs:{ba:'.400',pa:6},
  },
};

const pitchBA = {};
for (const yr of years) {
  const d = rawPitch[yr];
  pitchBA[yr] = {
    ff: d.ff.ba === '--' ? '--' : d.ff.ba.slice(1),
    si: d.si.ba === '--' ? '--' : d.si.ba.slice(1),
    ch: d.ch.ba === '--' ? '--' : d.ch.ba.slice(1),
    sl: weightedBA([d.sl, d.st]),
    cu: d.cu.ba === '--' ? '--' : d.cu.ba.slice(1),
    fc: d.fc.ba === '--' ? '--' : d.fc.ba.slice(1),
    fs: d.fs.ba === '--' ? '--' : d.fs.ba.slice(1),
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

// ===== 守備成績 (FanGraphs API pageitems=2000) =====
const fieldingRaw = {
  '2016': { '1B':{inn:'152',  drs:2},  'LF':{inn:'113',  drs:-2} },
  '2017': { '1B':{inn:'40',   drs:-1} },
  '2018': { '1B':{inn:'21',   drs:0},  'LF':{inn:'1',    drs:0}  },
  '2019': { '1B':{inn:'842',  drs:-1} },
  '2020': { '1B':{inn:'277.2',drs:1}  },
  '2021': { '1B':{inn:'587.1',drs:-3} },
  '2022': { '1B':{inn:'792.2',drs:-2} },
  '2023': { '1B':{inn:'123',  drs:0}  },
};

const positions = ['C','1B','2B','3B','SS','LF','CF','RF'];
const fieldingCareer = {};
for (const pos of positions) {
  const entries = years.map(yr => fieldingRaw[yr]?.[pos]).filter(Boolean);
  if (entries.length === 0) continue;
  fieldingCareer[pos] = {
    inn: addInnings(entries.map(e => e.inn)),
    drs: entries.reduce((acc, e) => acc + e.drs, 0),
  };
}

function getFieldVal(yearKey, pos, field) {
  const src = yearKey === '通算' ? fieldingCareer : (fieldingRaw[yearKey] || {});
  return src[pos]?.[field] ?? '--';
}

// ===== ワークシート構築 =====
const statsColsRow0 = [
  '年度','チーム','試合','打数','得点','安打','二塁打','三塁打','本塁打',
  '打点','四球','三振','盗塁','盗塁死',
  '打率','出塁率','OPS',
  '対左打率','得点圏打率',
  '走力',
  '４シーム','シンカー/2シーム','チェンジアップ','スライダー','カーブ','カット','スプリット',
];
const numStatsCols = statsColsRow0.length;

const headerRow0 = [...statsColsRow0, ...positions.flatMap(p => [p, ''])];
const headerRow1 = [...new Array(numStatsCols).fill(''), ...positions.flatMap(() => ['Inn','DRS'])];

function buildDataRow(yearKey) {
  const b  = basic[yearKey];
  const sp = splits[yearKey];
  const pt = pitchBA[yearKey];
  const statsVals = [
    yearKey, b.team,
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

// ===== Excelファイル生成 =====
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet(allRows);

const merges = [];
for (let c = 0; c < numStatsCols; c++) {
  merges.push({ s:{r:0,c}, e:{r:1,c} });
}
positions.forEach((_, i) => {
  const c = numStatsCols + i * 2;
  merges.push({ s:{r:0,c}, e:{r:0,c:c+1} });
});
ws['!merges'] = merges;

ws['!cols'] = [
  {wch:6},{wch:6},{wch:5},{wch:5},{wch:5},{wch:5},{wch:7},{wch:7},{wch:7},
  {wch:5},{wch:5},{wch:5},{wch:5},{wch:7},
  {wch:6},{wch:6},{wch:6},
  {wch:8},{wch:8},
  {wch:6},
  {wch:9},{wch:13},{wch:12},{wch:9},{wch:7},{wch:8},{wch:9},
  ...Array(16).fill({wch:7}),
];
ws['!freeze'] = { xSplit: 1, ySplit: 2 };

XLSX.utils.book_append_sheet(wb, ws, 'チェ_ジマン成績');

// ===== データソース・備考シート =====
const noteData = [
  ['データソース','内容'],
  ['MLB.com (MLB Stats API)','基本成績・対左打率・得点圏打率'],
  ['Baseball Savant','球種別打率・走力パーセンタイル(percent_speed_order)'],
  ['FanGraphs','守備成績 Pos/Inn/DRS (pageitems=2000で取得)'],
  ['',''],
  ['備考',''],
  ['対象年度','2016〜2023（MLB在籍期間）'],
  ['2016球種別打率','Baseball Savantにデータなし → 全て--'],
  ['2018チーム','MIL(12G)+TB(49G)の合算。TB優位のためTB2と表記'],
  ['2023チーム','PIT(23G)+SD(16G)の合算。PIT優位のためPIT2と表記'],
  ['打席','atBats（打数）を使用'],
  ['打率表記','頭の.を除去 (例: .261 → 261)'],
  ['スライダー','SL(従来型)+ST(スイーパー) PA加重平均'],
  ['球種別打率通算','PA加重平均による近似値（2016はデータなしのため除外）'],
  ['走力','Baseball Savant棒グラフのパーセンタイル値。通算は試合数加重平均'],
  ['守備','各ポジションのInn(守備イニング)とDRS(守備貢献値)'],
  ['FanGraphs取得方法','pageitems=500では取得不可。pageitems=2000で全件取得後フィルタ'],
];
const wsNote = XLSX.utils.aoa_to_sheet(noteData);
wsNote['!cols'] = [{wch:28},{wch:65}];
XLSX.utils.book_append_sheet(wb, wsNote, 'データソース・備考');

const outPath = 'チェ_ジマン_成績.xlsx';
XLSX.writeFile(wb, outPath);
console.log('✓ Created: ' + outPath);
console.log('  Rows: ' + (allRows.length - 2) + ' data rows (+ 2 header rows)');
console.log('  Cols: ' + headerRow0.length);
