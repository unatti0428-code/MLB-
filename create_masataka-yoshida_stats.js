const XLSX = require('xlsx');

function weightedBA(entries) {
  let sumH = 0, sumPA = 0;
  for (const e of entries) {
    if (!e || e.ba === '--' || !e.pa) continue;
    sumH  += parseFloat(e.ba.replace('__D__', '0.')) * e.pa;
    sumPA += e.pa;
  }
  if (sumPA === 0) return '--';
  return (sumH / sumPA).toFixed(3).split('.')[1];
}

function addInnings(list) {
  const total = list.filter(Boolean).reduce((acc, s) => {
    const [f, r] = String(s).split('.');
    return acc + parseInt(f) * 3 + parseInt(r || 0);
  }, 0);
  return Math.floor(total / 3) + '.' + (total % 3);
}

// ===== Step 1: MLB Stats API =====
const years = ["2023","2024","2025","2026"];
const basic = {
  "2023": { team:"BOS", g:140, pa:537, r:71,  h:155, d:33, t:3, hr:15, rbi:72,  bb:34, so:81,  sb:8,  cs:0, avg:".289", obp:".338", ops:".783" },
  "2024": { team:"BOS", g:108, pa:378, r:45,  h:106, d:21, t:0, hr:10, rbi:56,  bb:27, so:52,  sb:2,  cs:0, avg:".280", obp:".349", ops:".764" },
  "2025": { team:"BOS", g:55,  pa:188, r:16,  h:50,  d:11, t:0, hr:4,  rbi:26,  bb:10, so:24,  sb:3,  cs:0, avg:".266", obp:".307", ops:".695" },
  "2026": { team:"BOS", g:9,   pa:19,  r:2,   h:5,   d:2,  t:0, hr:0,  rbi:3,   bb:8,  so:4,   sb:1,  cs:0, avg:".263", obp:".500", ops:".868" },
  "通算": { team:"BOS", g:312, pa:1122,r:134, h:316, d:67, t:3, hr:29, rbi:157, bb:79, so:161, sb:14, cs:0, avg:".282", obp:".340", ops:".764" },
};
const splitsRaw = {
  "2023": { vsLAB:128, vsLH:35, rispAB:143, rispH:38 },
  "2024": { vsLAB:94,  vsLH:18, rispAB:91,  rispH:28 },
  "2025": { vsLAB:31,  vsLH:7,  rispAB:52,  rispH:12 },
  "2026": { vsLAB:6,   vsLH:2,  rispAB:5,   rispH:2  },
};

// ===== Step 2a: Baseball Savant 走力パーセンタイル =====
const sprintSpeed = { "2023":22, "2024":20, "2025":25, "2026":41 };

// ===== Step 2b: Baseball Savant 球種別打率 =====
// 注: Sweeper(ST) 2026は2PAのみ(BA=1.000)のため除外
const rawPitch = {
  "2023": {
    ff:{ba:'__D__370',pa:199}, si:{ba:'__D__287',pa:99},  ch:{ba:'__D__274',pa:68},
    sl:{ba:'__D__222',pa:69},  st:{ba:'__D__190',pa:21},  cu:{ba:'__D__269',pa:55},
    fc:{ba:'__D__119',pa:46},  fs:{ba:'__D__313',pa:17},
  },
  "2024": {
    ff:{ba:'__D__292',pa:126}, si:{ba:'__D__320',pa:87},  ch:{ba:'__D__286',pa:36},
    sl:{ba:'__D__256',pa:48},  st:{ba:'__D__150',pa:22},  cu:{ba:'__D__212',pa:34},
    fc:{ba:'__D__342',pa:40},  fs:{ba:'__D__231',pa:26},
  },
  "2025": {
    ff:{ba:'__D__338',pa:86},  si:{ba:'__D__263',pa:19},  ch:{ba:'__D__130',pa:23},
    sl:{ba:'__D__238',pa:23},  st:{ba:'__D__250',pa:8},   cu:{ba:'__D__143',pa:17},
    fc:{ba:'__D__167',pa:14},  fs:{ba:'__D__333',pa:13},
  },
  "2026": {
    ff:{ba:'__D__250',pa:10},  si:{ba:'__D__500',pa:5},   ch:{ba:'__D__200',pa:6},
    sl:{ba:'__D__000',pa:1},   st:{ba:'--',      pa:0},   cu:{ba:'__D__000',pa:3},
    fc:{ba:'__D__000',pa:1},   fs:{ba:'--',      pa:0},
  },
};

// ===== Step 2c: FanGraphs 守備成績 =====
const fieldingRaw = {
  "2023": { "LF":{ inn:"713.1", drs:-4 } },
  "2024": { "LF":{ inn:"1",     drs:0  } },
  "2025": { "LF":{ inn:"33",    drs:0  }, "RF":{ inn:"8", drs:0 } },
  "2026": { "LF":{ inn:"15",    drs:-2 } },
};

// ===== 計算処理 =====
const splits = {};
for (const yr of years) {
  const d = splitsRaw[yr];
  splits[yr] = {
    vsLeft: d.vsLAB === 0 ? '--' : (d.vsLH / d.vsLAB).toFixed(3).split('.')[1],
    risp:   d.rispAB === 0 ? '--' : (d.rispH / d.rispAB).toFixed(3).split('.')[1],
  };
}
const totVsLAB  = Object.values(splitsRaw).reduce((s,d) => s+d.vsLAB,  0);
const totVsLH   = Object.values(splitsRaw).reduce((s,d) => s+d.vsLH,   0);
const totRispAB = Object.values(splitsRaw).reduce((s,d) => s+d.rispAB, 0);
const totRispH  = Object.values(splitsRaw).reduce((s,d) => s+d.rispH,  0);
splits['通算'] = {
  vsLeft: (totVsLH  / totVsLAB ).toFixed(3).split('.')[1],
  risp:   (totRispH / totRispAB).toFixed(3).split('.')[1],
};

sprintSpeed['通算'] = Math.round(
  years.reduce((s, yr) => s + sprintSpeed[yr] * basic[yr].g, 0) /
  years.reduce((s, yr) => s + basic[yr].g, 0)
);

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
  const entries = years.map(yr => fieldingRaw[yr]?.[pos]).filter(Boolean);
  if (!entries.length) continue;
  fieldingCareer[pos] = {
    inn: addInnings(entries.map(e => e.inn)),
    drs: entries.reduce((s, e) => s + e.drs, 0),
  };
}
function getF(yk, pos, f) {
  return (yk === '通算' ? fieldingCareer : (fieldingRaw[yk]||{}))[pos]?.[f] ?? '--';
}

// ===== シート構築 =====
const cols0 = [
  '年度','チーム','試合','打数','得点','安打','二塁打','三塁打','本塁打',
  '打点','四球','三振','盗塁','盗塁死','打率','出塁率','OPS',
  '対左打率','得点圏打率','走力',
  '４シーム','シンカー/2シーム','チェンジアップ','スライダー','カーブ','カット','スプリット',
];
const nStat = cols0.length;
const hRow0 = [...cols0, ...positions.flatMap(p => [p,''])];
const hRow1 = [...Array(nStat).fill(''), ...positions.flatMap(() => ['Inn','DRS'])];

function buildRow(yk) {
  const b = basic[yk], sp = splits[yk], pt = pitchBA[yk];
  return [
    yk, b.team, b.g, b.pa, b.r, b.h, b.d, b.t, b.hr,
    b.rbi, b.bb, b.so, b.sb, b.cs,
    b.avg.slice(1), b.obp.slice(1), b.ops.slice(1),
    sp.vsLeft, sp.risp, sprintSpeed[yk],
    pt.ff, pt.si, pt.ch, pt.sl, pt.cu, pt.fc, pt.fs,
    ...positions.flatMap(p => [getF(yk,p,'inn'), getF(yk,p,'drs')]),
  ];
}

const allRows = [hRow0, hRow1, ...years.map(buildRow), buildRow('通算')];

const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet(allRows);
const merges = [
  ...Array.from({length: nStat}, (_, c) => ({s:{r:0,c},e:{r:1,c}})),
  ...positions.map((_, i) => { const c = nStat+i*2; return {s:{r:0,c},e:{r:0,c:c+1}}; }),
];
ws['!merges'] = merges;
ws['!cols'] = [
  {wch:6},{wch:6},{wch:5},{wch:5},{wch:5},{wch:5},{wch:7},{wch:7},{wch:7},
  {wch:5},{wch:5},{wch:5},{wch:5},{wch:7},{wch:6},{wch:6},{wch:6},
  {wch:8},{wch:8},{wch:6},
  {wch:9},{wch:13},{wch:12},{wch:9},{wch:7},{wch:8},{wch:9},
  ...Array(16).fill({wch:7}),
];
ws['!freeze'] = { xSplit:2, ySplit:2 };
XLSX.utils.book_append_sheet(wb, ws, '吉田正尚成績');

const note = [
  ['項目','説明'],
  ['打数','atBats（四球・死球・犠打飛を含まない）'],
  ['走力','Baseball Savant percent_speed_order（パーセンタイル）。通算は試合数加重平均'],
  ['スライダー','SL+ST(Sweeper) PA加重平均。2026 Sweeperは2PAのみ(BA=1.000)のため除外'],
  ['球種別打率通算','PA加重平均。データなし球種は除外'],
  ['守備','FanGraphs Inn/DRS（pageitems=2000で取得）'],
];
const wsN = XLSX.utils.aoa_to_sheet(note);
wsN['!cols'] = [{wch:20},{wch:65}];
XLSX.utils.book_append_sheet(wb, wsN, 'データソース・備考');

XLSX.writeFile(wb, '吉田正尚_成績.xlsx');
console.log('✓ Created: 吉田正尚_成績.xlsx');
console.log('  Rows: ' + (allRows.length - 2) + ' data rows (+ 2 header rows)');
console.log('  Cols: ' + hRow0.length);
