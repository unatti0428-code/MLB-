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
const browseFile    = () => browseFileWithFilter('Excel Files (*.xlsx)|*.xlsx');
const browseOdsFile = () => browseFileWithFilter('ODS Files (*.ods)|*.ods');

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

// 投球率: normalize all 7 values to sum to 100 (largest-remainder method)
// Input: array of 7 raw pct values (can be '--', "0.251", "25.1", "25.1%", etc.)
// Output: array of 7 integer strings summing to 100 (or '--' for missing)
function normalizePctToSum100(pctVals) {
  const nums = pctVals.map(v => {
    if (!v || v === '--') return null;
    const s = String(v).replace('%', '').trim();
    const n = parseFloat(s);
    if (isNaN(n) || n < 0) return null;
    if (n === 0) return 0;
    if (n <= 1.0) return n * 100;   // decimal fraction: 0.251 → 25.1
    return n;                         // already percentage: 25.1
  });

  const anyValid = nums.some(n => n !== null && n > 0);
  if (!anyValid) return pctVals.map(() => '--');

  const sum = nums.reduce((s, n) => s + (n || 0), 0);
  if (sum === 0) return pctVals.map(() => '--');

  const scaled  = nums.map(n => n !== null ? n / sum * 100 : null);
  const floors  = scaled.map(n => n !== null ? Math.floor(n) : 0);
  let remainder = 100 - floors.reduce((s, n) => s + n, 0);

  // Distribute remainder to items with largest fractional part
  const byFrac = scaled
    .map((n, i) => n !== null ? { i, frac: n - Math.floor(n) } : null)
    .filter(Boolean)
    .sort((a, b) => b.frac - a.frac);

  for (const { i } of byFrac) {
    if (remainder <= 0) break;
    floors[i]++;
    remainder--;
  }

  return nums.map((n, i) => n !== null ? String(floors[i]) : '--');
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

// ── ODS interaction (stats_tool2 logic) ──────────────────────────────────────
function findLibreOffice() {
  const candidates = [
    'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
    'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
  ];
  return candidates.find(p => { try { return fs.existsSync(p); } catch { return false; } }) || null;
}

function splitFormulaParts(str) {
  const parts = []; let depth = 0, cur = '';
  for (const ch of str) {
    if (ch === '(') { depth++; cur += ch; }
    else if (ch === ')') { depth--; cur += ch; }
    else if (ch === ',' && depth === 0) { parts.push(cur.trim()); cur = ''; }
    else cur += ch;
  }
  if (cur.trim()) parts.push(cur.trim());
  return parts;
}

function evalSpreadsheetFormula(formula, vars) {
  try {
    let f = String(formula);
    for (const [cell, val] of Object.entries(vars)) {
      f = f.replace(new RegExp(`\\b${cell}\\b`, 'g'), String(val));
    }
    for (let i = 0; i < 30; i++) {
      const prev = f;
      f = f.replace(/\bIF\(([^()]+)\)/gi, (_, inner) => {
        const p = splitFormulaParts(inner);
        return p.length === 3 ? `((${p[0]}) ? (${p[1]}) : (${p[2]}))` : _;
      });
      if (f === prev) break;
    }
    f = f.replace(/\bROUND\(([^,)]+),\s*\d+\)/gi, 'Math.round($1)');
    f = f.replace(/\bMAX\(([^)]+)\)/gi, 'Math.max($1)');
    f = f.replace(/\bMIN\(([^)]+)\)/gi, 'Math.min($1)');
    f = f.replace(/<>/g, '!==');
    f = f.replace(/([^<>!=])=([^=])/g, '$1==$2');
    // eslint-disable-next-line no-new-func
    return Function('"use strict"; return (' + f + ')')();
  } catch { return null; }
}

async function getControlRating(odsPath, bb9Value) {
  const wb = XLSX.readFile(odsPath, { cellFormula: true, cellDates: false, type: 'file' });
  const wsName = wb.SheetNames[0];
  const ws = wb.Sheets[wsName];
  const ac2Formula = ws['AC2']?.f || null;

  ws['V2'] = { t: 'n', v: bb9Value };
  XLSX.writeFile(wb, odsPath);

  const lo = findLibreOffice();
  if (lo) {
    const dir = path.dirname(odsPath);
    spawnSync(lo, [
      '--headless', '--norestore', '--infilter=calc8',
      '--convert-to', 'ods', '--outdir', dir, odsPath
    ], { timeout: 30000 });
    try {
      const wb2 = XLSX.readFile(odsPath, { cellFormula: true });
      const ws2 = wb2.Sheets[wb2.SheetNames[0]];
      const v = ws2['AC2']?.v;
      if (v != null) return v;
    } catch {}
  }

  if (ac2Formula) {
    const result = evalSpreadsheetFormula(ac2Formula, { V2: bb9Value });
    if (result != null) return result;
  }
  return ws['AC2']?.v ?? 0;
}

// ── Cell styling ──────────────────────────────────────────────────────────────
const PURPLE_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7030A0' } };
function purpleCell(cell, value, fs) {
  cell.value = value;
  cell.fill  = { ...PURPLE_FILL };
  cell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: fs };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
}

// ── Add 制球 column to pitcher Excel (stats_tool2 logic) ─────────────────────
// Col 51 = AY = 制球 (after 22 main + 7×4=28 pitch cols)
const SEIKYU_COL = 51;

async function addSeikyuToFile(xlsxPath, odsPath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(xlsxPath);
  const ws = wb.worksheets[0];
  const fontSize = ws.getCell(1, 1).font?.size || 11;

  purpleCell(ws.getCell(1, SEIKYU_COL), '制球', fontSize);

  const dataRows = [];
  ws.eachRow((row, rn) => {
    if (rn <= 2) return;
    const yr = row.getCell(2).value;
    if (!yr) return;
    const ipStr = String(row.getCell(11).value ?? '').trim();
    if (!ipStr || ipStr === '--') return;
    const bb = Number(row.getCell(15).value) || 0;
    dataRows.push({ rn, ipStr, bb });
  });

  let count = 0;
  for (const { rn, ipStr, bb } of dataRows) {
    const ip = parseIP(ipStr);
    if (!ip) continue;
    const bb9 = bb / ip * 9;
    const controlVal = await getControlRating(odsPath, bb9);
    purpleCell(ws.getCell(rn, SEIKYU_COL), Math.round(controlVal), fontSize);
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
const PITCH_MAP_P    = {
  '4-Seam Fastball': 'ff', '4-seam Fastball': 'ff', 'Four-Seam Fastball': 'ff',
  'Slider': 'sl', 'Sweeper': 'sl',
  'Changeup': 'ch', 'Change-up': 'ch',
  'Curveball': 'cu', 'Knuckle Curve': 'cu', 'Knuckleball': 'cu',
  'Cutter': 'fc',
  'Sinker': 'si', 'Two-Seam Fastball': 'si', '2-Seam Fastball': 'si',
  'Split-Finger': 'fs', 'Splitter': 'fs', 'Split Finger': 'fs',
};

const emptyPitchP = () => Object.fromEntries(
  PITCH_KEYS.map(k => [k, { velo: '--', ba: '--', slg: '--', pct: '--' }])
);

// ── Baseball Savant browser scraping ─────────────────────────────────────────
async function fetchBrowserData(slug, id, years, onProgress) {
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

    const y1 = years[0], y2 = years[years.length - 1];
    const rawPitch = {};
    for (const yr of years) rawPitch[yr] = emptyPitchP();

    try {
      onProgress('Baseball Savant (投手) を読み込み中...');
      const savantUrl = `https://baseballsavant.mlb.com/savant-player/${slug}-${id}` +
        `?stats=statcast&player_type=pitcher&startSeason=${y1}&endSeason=${y2}`;
      await page.goto(savantUrl, { waitUntil: 'networkidle2', timeout: 60000 });

      const savantRaw = await page.evaluate(() => {
        try {
          const tables = document.querySelectorAll('table');
          let pitchTable = null;
          for (const t of tables) {
            const txt = t.innerText || '';
            if (txt.includes('4-Seam') || txt.includes('Sinker') || txt.includes('Slider') ||
                txt.includes('Fastball') || txt.includes('Curveball')) {
              pitchTable = t; break;
            }
          }
          if (!pitchTable) return { pitchData: {}, debug: 'no_table' };

          const headers = [...pitchTable.querySelectorAll('thead th,thead td')]
            .map(h => h.innerText.trim().toLowerCase());

          let yearIdx  = headers.findIndex(h => h === 'year' || h === 'season');
          let pitchIdx = headers.findIndex(h => h === 'pitch' || h === 'pitch name' || h === 'pitch type' || h === 'type');
          let veloIdx  = headers.findIndex(h => h.includes('velo') || h.includes('velocity') || h === 'avg velo' || h === 'mph');
          let baIdx    = headers.findIndex(h => h === 'ba' || h === 'avg' || h === 'batting avg' || h === 'batting average');
          let slgIdx   = headers.findIndex(h => h === 'slg' || h === 'slugging' || h.startsWith('slg'));
          let pctIdx   = headers.findIndex(h => h === '%' || h === 'usage' || h === 'usage%' || h === 'pct' ||
                           (h.includes('%') && !h.includes('k%') && !h.includes('bb%') && !h.includes('zone')));

          if (pitchIdx < 0) pitchIdx = 1;
          if (yearIdx < 0)  yearIdx  = 0;

          const pitchData = {};
          for (const row of pitchTable.querySelectorAll('tbody tr')) {
            const cells = [...row.querySelectorAll('td')];
            if (cells.length < 2) continue;
            const yr = cells[yearIdx]?.innerText.trim();
            const pt = cells[pitchIdx]?.innerText.trim();
            if (!yr || !pt) continue;
            if (!pitchData[yr]) pitchData[yr] = {};
            pitchData[yr][pt] = {
              velo: veloIdx >= 0 ? cells[veloIdx]?.innerText.trim() : '--',
              ba:   baIdx   >= 0 ? cells[baIdx]?.innerText.trim()   : '--',
              slg:  slgIdx  >= 0 ? cells[slgIdx]?.innerText.trim()  : '--',
              pct:  pctIdx  >= 0 ? cells[pctIdx]?.innerText.trim()  : '--',
            };
          }
          return { pitchData };
        } catch (e) {
          return { pitchData: {}, error: e.message };
        }
      });

      if (savantRaw.error) onProgress('⚠ Baseball Savant エラー: ' + savantRaw.error);

      for (const yr of years) {
        const yrData = savantRaw.pitchData?.[yr] || {};
        for (const [ptName, vals] of Object.entries(yrData)) {
          const key = PITCH_MAP_P[ptName];
          if (!key) continue;
          rawPitch[yr][key] = {
            velo: vals.velo || '--',
            ba:   vals.ba   || '--',
            slg:  vals.slg  || '--',
            pct:  vals.pct  || '--',
          };
        }
      }
    } catch (e) {
      onProgress('⚠ Baseball Savant 取得失敗（空データで続行）: ' + e.message);
    }

    return { rawPitch };
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
      pct:  wAvg('pct',  true),
    };
  }

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

    upd('ブラウザを起動して Baseball Savant から球種データを取得中...');
    const { rawPitch } = await fetchBrowserData(params.slug, params.id, years, upd);

    upd('Excel ファイルを生成中...');
    const outFile = await buildExcel(params.name, years, basic, vsLeftByYear, rawPitch);

    let seikyuRows = 0;
    if (params.odsPath) {
      upd('制球を計算中（守備.ods と連携）...');
      try {
        seikyuRows = await addSeikyuToFile(outFile, params.odsPath);
        upd(`制球追加完了: ${seikyuRows} 行`);
      } catch (e) {
        upd('⚠ 制球追加失敗: ' + e.message);
      }
    }

    const j = jobs.get(jobId);
    if (j) {
      j.status   = 'done';
      j.result   = path.basename(outFile);
      j.seikyuRows = seikyuRows;
      j.progress = '完了';
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

    <div class="sec">
      <label>③ 守備.ods <span class="opt-label">（制球算出用・任意）</span></label>
      <div class="ods-area" id="odsArea">
        <div class="ods-icon">📋</div>
        <div class="ods-path" id="odsPathDisp">未選択（制球列をスキップ）</div>
      </div>
      <button class="btn-s" style="font-size:13px;padding:8px 16px" onclick="browseOds()">
        📂 守備.ods を選択...
      </button>
    </div>

    <div class="sec" style="margin-top:4px">
      <div class="badge-row">
        <span class="badge">投手成績取得</span>
        <span style="font-size:14px;color:#aaa;align-self:center">→</span>
        <span class="badge">Excel生成</span>
        <span style="font-size:14px;color:#aaa;align-self:center">→</span>
        <span class="badge red">7球種 × 4項目</span>
        <span style="font-size:14px;color:#aaa;align-self:center">→</span>
        <span class="badge" id="seikyuBadge">制球（守備.ods連携）</span>
      </div>
      <button class="btn-p" id="btnCreate" onclick="doCreate()">▶ 成績ファイルを作成</button>
    </div>
    <div class="pbox" id="cPbox"><div class="ptxt" id="cPtxt"><span class="sp"></span>処理中...</div></div>
    <div class="done" id="cDone"></div>
    <div class="err"  id="cErr"></div>
    <div class="note">
      ※ Chromeが自動起動します（Baseball Savant へのアクセス）<br>
      ※ 守備.ods を選択すると制球（AY列）が自動計算・追加されます<br>
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
let cTimer = null, odsFilePath = '';

function switchTab(id, el) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  document.getElementById('panel-' + id).classList.add('active');
}

async function browseOds() {
  try {
    const r = await fetch('/api/browse-ods');
    const d = await r.json();
    if (d.path) {
      odsFilePath = d.path;
      const el = document.getElementById('odsPathDisp');
      el.textContent = d.path;
      el.className = 'ods-path has';
      document.getElementById('odsArea').className = 'ods-area sel';
      document.getElementById('seikyuBadge').style.background = '#7030A0';
      document.getElementById('seikyuBadge').style.color = 'white';
    }
  } catch(e) { alert('ODS選択エラー: ' + e.message); }
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
  if (!slug||!id||!name||!fullName||!y1||!y2){alert('すべての項目を入力してください');return;}
  document.getElementById('btnCreate').disabled=true;
  document.getElementById('cPbox').style.display='block';
  document.getElementById('cDone').style.display='none';
  document.getElementById('cErr').style.display='none';
  setCP('処理を開始しています...');
  const r=await fetch('/api/create',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({slug,id,name,fullName,y1,y2,odsPath:odsFilePath})});
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
    if (j.seikyuRows > 0) msg += '（制球: ' + j.seikyuRows + ' 行追加）';
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
  if (req.method === 'GET' && url.pathname === '/api/browse-ods') {
    const fp = browseOdsFile();
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
        jobs.set(jobId, { status:'running', progress:'開始中...', result:null, seikyuRows:0, error:null });
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
