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

const PORT    = 3940;
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

// ── PowerShell file browser ───────────────────────────────────────────────────
function browseFile() {
  const r = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', `
[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
$d = New-Object System.Windows.Forms.OpenFileDialog
$d.Filter = "Excel Files (*.xlsx)|*.xlsx"
$d.Title = "Excel file"
if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($d.FileName)
  [Convert]::ToBase64String($bytes)
}`], { encoding: 'buffer' });
  const b64 = (r.stdout || Buffer.alloc(0)).toString('ascii').trim();
  return b64 ? Buffer.from(b64, 'base64').toString('utf8') : '';
}

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

async function fetchMLBStats(id, y1, y2) {
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
      teamStr = (primary.team?.abbreviation || primary.team?.name?.slice(0,3)?.toUpperCase() || '???') + named.length;
    } else {
      teamStr = row.team?.abbreviation || row.team?.name?.slice(0,3)?.toUpperCase() || '???';
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

  // キャッチャー守備成績（CS/SB率算出用）
  const catcherFielding = { byYear: {}, career: null };
  try {
    const [fldYby, fldCar] = await Promise.all([
      mlbGet(`https://statsapi.mlb.com/api/v1/people/${id}/stats?stats=yearByYear&group=fielding&sportId=1`),
      mlbGet(`https://statsapi.mlb.com/api/v1/people/${id}/stats?stats=career&group=fielding&sportId=1`),
    ]);
    for (const s of (fldYby.stats?.[0]?.splits || [])) {
      if (s.position?.abbreviation === 'C' && s.season && s.sport?.id === 1) {
        catcherFielding.byYear[s.season] = {
          sb: s.stat?.stolenBases ?? 0,
          cs: s.stat?.caughtStealing ?? 0,
        };
      }
    }
    const carCat = (fldCar.stats?.[0]?.splits || []).find(s => s.position?.abbreviation === 'C');
    if (carCat) catcherFielding.career = { sb: carCat.stat?.stolenBases ?? 0, cs: carCat.stat?.caughtStealing ?? 0 };
  } catch {}

  return { years, basic, splitsRaw, catcherFielding };
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

async function fetchBrowserData(slug, id, playerFullName, years, onProgress, splitsRaw = {}) {
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

      onProgress('FanGraphs から守備データを取得中...');
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
      }, years, playerFullName);

      for (const yr of years) {
        const entries = Array.isArray(fieldingRaw[yr]) ? fieldingRaw[yr] : [];
        for (const { pos, inn, drs } of entries) {
          fieldingByYear[yr][pos] = { inn: String(inn), drs: Number(drs) };
        }
      }
    } catch (e) {
      onProgress('⚠ FanGraphs 取得失敗（空データで続行）: ' + e.message);
    }

    // ── Baseball Reference フォールバック（歴代選手・DRS欠損/スプリット欠損時）────
    // 発動条件: FanGraphs で守備データが取れない年あり、または MLB API スプリット全欠損
    const bbRefSplits = {};
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
            // networkidle2 で JS 遅延ロードのテーブルも確実に取得する
            await page.goto(
              `https://www.baseball-reference.com/players/${bbSlug[0]}/${bbSlug}.shtml`,
              { waitUntil: 'networkidle2', timeout: 30000 }
            );
          }

          // ── 診断ログ（スキップ原因を特定）────────────────────────────────────
          onProgress(`[診断] bbSlug=${bbSlug}, missingFieldingYears=${missingFieldingYears.length}/${years.length}, allSplitsEmpty=${allSplitsEmpty}`);

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

                  // ヘッダーの data-stat（デバッグ用）
                  const headerStats = [...table.querySelectorAll('[data-stat]')]
                    .slice(0, 30).map(el => el.getAttribute('data-stat')).filter(Boolean);

                  const result = {};
                  for (const row of table.querySelectorAll('tbody tr')) {
                    if (row.classList.contains('thead') || row.classList.contains('minors_table')) continue;
                    const yearTh = row.querySelector('th[data-stat="year_ID"]');
                    const yr = yearTh ? yearTh.textContent.replace(/\D/g, '').trim() : '';
                    if (!yr || yr.length !== 4) continue;
                    const get = s => {
                      const el = row.querySelector(`td[data-stat="${s}"]`);
                      return el ? el.textContent.trim() : '';
                    };
                    const pos = get('pos');
                    if (!pos || pos === 'Pos') continue;
                    result[yr] = result[yr] || {};
                    result[yr][pos] = {
                      inn:   get('Inn') || get('inn_outs') || get('inn'),
                      ch:    get('chances') || '0',
                      e:     get('e')       || '0',
                      fld:   get('fielding_perc')        || '0',
                      lgFld: get('lg_fielding_perc')     || '0',
                      rf9:   get('range_factor_9inn')    || '0',
                      lgRf9: get('lg_range_factor_9inn') || '0',
                    };
                  }
                  return { result, headerStats, tableIds };
                } catch (e) { return { err: e.message }; }
              }, cleanHtml);

              if (bbResult?.err) {
                onProgress(`⚠ BB-Ref 守備テーブル未検出 (${bbResult.err})`);
                onProgress(`  検出テーブルIDs: [${(bbResult.tableIds||[]).slice(0, 12).join(', ')}]`);
              } else if (bbResult?.result) {
                onProgress(`BB-Ref 守備ヘッダー: [${(bbResult.headerStats||[]).slice(0, 12).join(', ')}]`);
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
                    const ch    = parseInt(d.ch)    || 0;
                    const fld   = parseFloat(d.fld)   || 0;
                    const lgFld = parseFloat(d.lgFld) || 0;
                    const rf9   = parseFloat(d.rf9)   || 0;
                    const lgRf9 = parseFloat(d.lgRf9) || 0;
                    if (innDec < 1 || !lgRf9) continue;
                    const rangeDRS = (rf9 - lgRf9) * innDec / 9 * 0.75;
                    const errorDRS = ch > 0 ? ch * (fld - lgFld) * 0.5 : 0;
                    fieldingByYear[yr][pos] = { inn: innFmt, drs: Math.round(rangeDRS + errorDRS) };
                  }
                  if (Object.keys(fieldingByYear[yr]).length > 0)
                    onProgress(`BB-Ref 守備推定 ${yr}: ${Object.keys(fieldingByYear[yr]).join(', ')}`);
                }
              }
            } catch (e) {
              onProgress(`⚠ BB-Ref 守備HTML取得失敗: ${e.message}`);
            }
          }

          // ── スプリット取得（通算 → 年別欠損の補完値として使用）─────────────
          try {
            const splitUrl = `https://www.baseball-reference.com/players/split.fcgi?id=${bbSlug}&year=all&t=b`;
            await page.goto(splitUrl, { waitUntil: 'networkidle2', timeout: 25000 });

            // スプリットページも HTML コメント内にテーブルが入っている場合があるため
            // Node.js 側で page.content() を取得してコメント除去後に DOMParser で解析
            const splitRawHtml = await page.content();
            const splitCleanHtml = splitRawHtml.replace(/<!--([\s\S]*?)-->/g, '$1');
            onProgress(`BB-Ref スプリットHTML: ${splitRawHtml.length.toLocaleString()} chars`);

            const bbSplit = await page.evaluate((html) => {
              try {
                const doc2 = new DOMParser().parseFromString(html, 'text/html');
                const parseSplit = (doc) => {
                  for (const table of doc.querySelectorAll('table')) {
                    let vsLRow = null, rispRow = null;
                    for (const row of table.querySelectorAll('tbody tr, tr')) {
                      // data-stat="split" の th か、最初の th/td のテキストで判定
                      const splitCell = row.querySelector('[data-stat="split"]') ||
                                        row.querySelector('th') ||
                                        row.querySelector('td');
                      const txt = (splitCell?.textContent || '').trim().toLowerCase();
                      if (!vsLRow  && (txt === 'vs. lhp' || txt === 'lhp' || txt === 'left' || txt === 'vs lhp'))
                        vsLRow = row;
                      if (!rispRow && (txt === 'risp' || txt.includes('scoring position') || txt === 'bases loaded'))
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
                      };
                      if (result.vsLAB > 0 || result.rispAB > 0) return result;
                    }
                  }
                  return null;
                };
                // まずコメント除去済みHTMLを試し、なければ元のページを試す
                return parseSplit(doc2) || parseSplit(document);
              } catch (e) { return null; }
            }, splitCleanHtml);

            if (bbSplit && (bbSplit.vsLAB > 0 || bbSplit.rispAB > 0)) {
              for (const yr of years) bbRefSplits[yr] = bbSplit;
              onProgress(`BB-Ref スプリット(通算): 対左 ${bbSplit.vsLAB}AB / RISP ${bbSplit.rispAB}AB`);
            } else {
              onProgress('BB-Ref スプリット: データなし');
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

    return { sprintSpeed, rawPitch, fieldingByYear, mlbTheShowSpeed, catcherFraming, bbRefSplits };
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

async function buildExcel(playerName, years, basic, splitsRaw, sprintSpeed, mlbTheShowSpeed, rawPitch, fieldingByYear) {
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
    return (yk === '通算' ? fieldingCareer : (fieldingByYear[yk] || {}))[pos]?.[field] ?? '--';
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
  const thresholds = [[329,142],[339,152],[349,162],[359,172],[369,182],[379,192]];
  const [hi, lo] = thresholds[tier] || thresholds[0];
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

const PURPLE_FILL    = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF7030A0' } };
const REDPURPLE_FILL = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFC00060' } };
function styledCell(cell, value, fill, fontSize) {
  cell.value = value;
  cell.fill  = { ...fill };
  cell.font  = { bold:true, color:{ argb:'FFFFFFFF' }, size:fontSize };
  cell.alignment = { horizontal:'center', vertical:'middle' };
}
function purpleCell(cell, value, fs)    { styledCell(cell, value, PURPLE_FILL,    fs); }
function redPurpleCell(cell, value, fs) { styledCell(cell, value, REDPURPLE_FILL, fs); }

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
  PITCH_HEADERS.forEach((h, i) => redPurpleCell(ws.getCell(1, PITCH_COL + i), h, fontSize));

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

    const pitchRaw  = [22,23,24,25,26,27,28].map(c => parseBA(row.getCell(c).value));
    const pitchVals = calcPitchRatings(pitchRaw);
    pitchVals.forEach((v, i) => {
      if (v !== '') redPurpleCell(ws.getCell(rn, PITCH_COL + i), v, fontSize);
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

// ── Job management ────────────────────────────────────────────────────────────
const jobs = new Map();

async function runCreateJob(jobId, params) {
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
    const { years, basic, splitsRaw, catcherFielding } = await fetchMLBStats(params.id, params.y1, params.y2);

    upd('ブラウザを起動して Baseball Savant / FanGraphs を取得中...');
    const { sprintSpeed, rawPitch, fieldingByYear, mlbTheShowSpeed, catcherFraming, bbRefSplits } =
      await fetchBrowserData(params.slug, params.id, params.fullName, years, upd, splitsRaw);

    // BB-Ref 通算スプリットで MLB Stats API の欠損年を補完
    // （歴代選手など sitCodes API が空を返す場合に使用）
    for (const yr of years) {
      const d  = splitsRaw[yr];
      const bb = bbRefSplits?.[yr];
      if (!d || !bb) continue;
      if (!d.vsLAB  && bb.vsLAB)  { d.vsLAB  = bb.vsLAB;  d.vsLH   = bb.vsLH;  }
      if (!d.rispAB && bb.rispAB) { d.rispAB = bb.rispAB; d.rispH  = bb.rispH; }
    }

    upd('Excel ファイルを生成中...');
    const outFile = await buildExcel(params.name, years, basic, splitsRaw, sprintSpeed, mlbTheShowSpeed, rawPitch, fieldingByYear);

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

// ── HTML ──────────────────────────────────────────────────────────────────────
const HTML = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>MLB成績ツール</title>
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
.file-area{border:2px dashed #ce93d8;border-radius:8px;padding:16px;min-height:56px;
  display:flex;align-items:center;gap:12px;background:#faf5ff;margin-bottom:12px}
.file-area.sel{border-color:#7030A0;background:#f3e5f5}
.file-icon{font-size:26px;flex-shrink:0}
.file-path{font-size:13px;color:#888;flex:1;word-break:break-all}
.file-path.has{color:#4a1470;font-weight:bold}
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
</style>
</head>
<body>
<div class="card">
  <h1>MLB成績ツール</h1>
  <div class="tabs">
    <div class="tab active" onclick="switchTab('create',this)">新規作成</div>
    <div class="tab" onclick="switchTab('add',this)">既存ファイルに追加</div>
  </div>

  <!-- ── Tab 1: 新規作成 ── -->
  <div id="panel-create" class="panel active">
    <div class="sec">
      <label>① 選手検索（英語名）</label>
      <div class="row">
        <div style="flex:3"><input id="q" type="text" placeholder="例: Masataka Yoshida"
          onkeydown="if(event.key==='Enter')doSearch()"></div>
        <div style="flex:1"><button class="btn-s" onclick="doSearch()">🔍 検索</button></div>
      </div>
      <div class="results" id="results"></div>
    </div>
    <div class="sec">
      <label>② 選手情報</label>
      <div class="row">
        <div><label>英語スラッグ</label><input id="slug" type="text" placeholder="masataka-yoshida"></div>
        <div><label>MLB ID</label><input id="pid" type="number" placeholder="807799"></div>
      </div>
      <div class="row" style="margin-top:10px">
        <div><label>日本語名（ファイル名）</label><input id="jaName" type="text" placeholder="吉田正尚"></div>
        <div><label>FanGraphs 表示名（英語）</label><input id="fullName" type="text" placeholder="Masataka Yoshida"></div>
      </div>
      <div class="row" style="margin-top:10px">
        <div><label>開始年</label><input id="y1" type="number" placeholder="2023"></div>
        <div><label>終了年</label><input id="y2" type="number" placeholder="2026"></div>
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
      <button class="btn-p" id="btnCreate" onclick="doCreate()">▶ 成績ファイルを作成</button>
    </div>
    <div class="pbox" id="cPbox"><div class="ptxt" id="cPtxt"><span class="sp"></span>処理中...</div></div>
    <div class="done" id="cDone"></div>
    <div class="err"  id="cErr"></div>
    <div class="note">※ Chromeが自動起動します（Baseball Savant / FanGraphsへのアクセス）<br>
      ※ 出力先: このツールと同じフォルダ</div>
  </div>

  <!-- ── Tab 2: 既存ファイルに追加 ── -->
  <div id="panel-add" class="panel">
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
      <div class="file-area" id="fileArea">
        <div class="file-icon">📄</div>
        <div class="file-path" id="filePath">ファイルが選択されていません</div>
      </div>
      <div style="display:flex;gap:8px;align-items:center">
        <button class="btn-s" onclick="doBrowse()">📂 ファイルを参照...</button>
        <button class="btn-p" id="btnAdd" disabled onclick="doAdd()">✓ 能力値を追加</button>
        <span id="addStatus" style="font-size:12px;color:#888"></span>
      </div>
    </div>
    <div class="pbox" id="aPbox"><div class="ptxt" id="aPtxt"><span class="sp"></span>処理中...</div></div>
    <div class="done" id="aDone"></div>
    <div class="err"  id="aErr"></div>
  </div>
</div>

<script>
let cTimer = null, selectedPath = '';

function switchTab(id, el) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  document.getElementById('panel-' + id).classList.add('active');
}

// ── Tab1: 新規作成 ────────────────────────────────────────────
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
    body:JSON.stringify({slug,id,name,fullName,y1,y2})});
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
    el.textContent='✓ 完了: '+j.result+'（'+j.rows+' 行に能力値追加済）';
  } else if (j.status==='error') {
    clearInterval(cTimer);
    document.getElementById('cPbox').style.display='none';
    document.getElementById('btnCreate').disabled=false;
    const el=document.getElementById('cErr'); el.style.display='block';
    el.textContent='✗ エラー: '+j.error;
  }
}
function setCP(msg){document.getElementById('cPtxt').innerHTML='<span class="sp"></span>'+msg;}

// ── Tab2: 既存ファイルに追加 ──────────────────────────────────
async function doBrowse() {
  document.getElementById('addStatus').textContent='ダイアログを開いています...';
  const r=await fetch('/api/browse'); const data=await r.json();
  if (data.path) {
    selectedPath=data.path;
    const fp=document.getElementById('filePath');
    fp.textContent=data.path; fp.className='file-path has';
    document.getElementById('fileArea').className='file-area sel';
    document.getElementById('btnAdd').disabled=false;
    document.getElementById('addStatus').textContent='';
  } else { document.getElementById('addStatus').textContent='選択されませんでした'; }
}
async function doAdd() {
  if (!selectedPath) return;
  document.getElementById('btnAdd').disabled=true;
  document.getElementById('aPbox').style.display='block';
  document.getElementById('aDone').style.display='none';
  document.getElementById('aErr').style.display='none';
  document.getElementById('aPtxt').innerHTML='<span class="sp"></span>処理中...';
  try {
    const r=await fetch('/api/process',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({filePath:selectedPath})});
    const data=await r.json();
    document.getElementById('aPbox').style.display='none';
    document.getElementById('btnAdd').disabled=false;
    if (data.success) {
      const el=document.getElementById('aDone'); el.style.display='block';
      el.textContent='✓ 完了: '+data.count+' 行に能力値を書き込みました';
    } else {
      const el=document.getElementById('aErr'); el.style.display='block';
      el.textContent='✗ エラー: '+data.error;
    }
  } catch(e) {
    document.getElementById('aPbox').style.display='none';
    document.getElementById('btnAdd').disabled=false;
    const el=document.getElementById('aErr'); el.style.display='block';
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
        jobs.set(jobId, { status:'running', progress:'開始中...', result:null, rows:0, error:null });
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
