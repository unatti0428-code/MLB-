'use strict';
const http    = require('http');
const ExcelJS = require('exceljs');
const XLSX    = require('xlsx');
const path    = require('path');
const fs      = require('fs');
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

// "200.1" → 200.333..., "200.2" → 200.667
function parseIP(ipStr) {
  const s = String(ipStr || '').trim();
  if (!s || s === '--') return 0;
  const [whole, frac] = s.split('.');
  return (parseInt(whole) || 0) + (parseInt(frac || 0)) / 3;
}

function findLibreOffice() {
  const candidates = [
    'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
    'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
  ];
  return candidates.find(p => { try { return fs.existsSync(p); } catch { return false; } }) || null;
}

// Split formula parts respecting nested parentheses
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

// Minimal spreadsheet formula evaluator (handles IF, arithmetic, comparisons)
function evalSpreadsheetFormula(formula, vars) {
  try {
    let f = String(formula);
    // Replace cell references (e.g. V2)
    for (const [cell, val] of Object.entries(vars)) {
      f = f.replace(new RegExp(`\\b${cell}\\b`, 'g'), String(val));
    }
    // Convert nested IF(cond,then,else) to ternary
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
    // Careful = → == (don't double-convert >=, <=, ==, !=)
    f = f.replace(/([^<>!=])=([^=])/g, '$1==$2');
    // eslint-disable-next-line no-new-func
    return Function('"use strict"; return (' + f + ')')();
  } catch { return null; }
}

// Write bb9Value to V2 in 守備.ods, recalculate via LibreOffice, read AC2
// Falls back to formula parsing if LibreOffice not found
async function getControlRating(odsPath, bb9Value) {
  const wb = XLSX.readFile(odsPath, { cellFormula: true, cellDates: false, type: 'file' });
  const wsName = wb.SheetNames[0];
  const ws = wb.Sheets[wsName];

  // Save original AC2 formula (before overwriting)
  const ac2Formula = ws['AC2']?.f || null;

  // Write V2
  ws['V2'] = { t: 'n', v: bb9Value };
  XLSX.writeFile(wb, odsPath);

  // Try LibreOffice headless recalculation
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

  // Fallback: evaluate formula from ODS
  if (ac2Formula) {
    const result = evalSpreadsheetFormula(ac2Formula, { V2: bb9Value });
    if (result != null) return result;
  }

  // Last resort: return cached value
  return ws['AC2']?.v ?? 0;
}

const PURPLE_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7030A0' } };
function purpleCell(cell, value, fs) {
  cell.value = value;
  cell.fill  = { ...PURPLE_FILL };
  cell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: fs };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
}

// Col 51 = 制球 (after 22 main stats + 7×4=28 pitch cols)
const SEIKYU_COL = 51;

async function processFile(xlsxPath, odsPath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(xlsxPath);
  const ws = wb.worksheets[0];
  const fontSize = ws.getCell(1, 1).font?.size || 11;

  // Write 制球 header in row 1 (merged header row)
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
  padding:32px;width:100%;max-width:580px}
h1{color:#7030A0;font-size:20px;margin-bottom:6px;display:flex;align-items:center;gap:8px}
h1::before{content:"⚾";font-size:22px}
.subtitle{color:#888;font-size:12px;margin-bottom:20px}
.sec{margin-bottom:20px}
.sec-label{font-size:12px;font-weight:bold;color:#555;margin-bottom:6px;text-transform:uppercase;letter-spacing:.5px}
.file-area{border:2px dashed #ce93d8;border-radius:8px;padding:14px;min-height:54px;
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
.btn-run{background:#7030A0;color:white;min-width:110px}
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
</style>
</head>
<body>
<div class="card">
  <h1>MLB投手成績 能力値追加ツール</h1>
  <div class="subtitle">守備.ods の計算式を元に、投手成績.xlsx へ <strong>制球</strong> 列を自動追加します</div>

  <div class="sec">
    <div class="sec-label">追加される列</div>
    <span class="badge">AY 制球</span>
    <p style="font-size:11px;color:#888;margin-top:8px">
      四死球 ÷ 換算イニング × 9 → 守備.ods V2 へ入力 → AC2 の値を取得
    </p>
  </div>

  <div class="sec">
    <div class="sec-label">① 投手成績 Excel ファイル</div>
    <div class="file-area" id="fa1">
      <div class="file-icon">📊</div>
      <div class="file-path" id="fp1">ファイルが選択されていません</div>
    </div>
    <button class="btn-browse" onclick="browse(1)">📂 参照...</button>
  </div>

  <div class="sec">
    <div class="sec-label">② 守備.ods ファイル</div>
    <div class="file-area" id="fa2">
      <div class="file-icon">📋</div>
      <div class="file-path" id="fp2">ファイルが選択されていません</div>
    </div>
    <button class="btn-browse" onclick="browse(2)">📂 参照...</button>
  </div>

  <div class="sec">
    <div class="btn-row">
      <button class="btn-run" id="btnRun" disabled onclick="run()">▶ 制球を追加</button>
      <span class="status" id="status"></span>
    </div>
    <div class="result" id="result"></div>
  </div>

  <div class="note">
    <strong>計算ロジック</strong><br>
    イニング変換: X.1 → X+1/3, X.2 → X+2/3<br>
    BB9 = 四死球 ÷ 換算イニング × 9 → 守備.ods V2 へ書き込み<br>
    AC2 の値（制球評価）を読み取り → AY 列へ書き込み<br>
    ※ LibreOffice がインストールされていると自動再計算されます
  </div>
</div>

<script>
let paths = { 1: '', 2: '' };

async function browse(n) {
  setStatus('<span class="sp"></span>ダイアログを開いています...');
  const filter = n === 1 ? 'xlsx' : 'ods';
  try {
    const r = await fetch('/api/browse?filter=' + filter);
    const d = await r.json();
    if (d.path) {
      paths[n] = d.path;
      document.getElementById('fp' + n).textContent = d.path;
      document.getElementById('fp' + n).className = 'file-path has';
      document.getElementById('fa' + n).className = 'file-area sel';
    }
    setStatus('');
  } catch(e) { setStatus('エラー: ' + e.message); }
  document.getElementById('btnRun').disabled = !(paths[1] && paths[2]);
}

async function run() {
  if (!paths[1] || !paths[2]) return;
  const btn = document.getElementById('btnRun');
  btn.disabled = true;
  setStatus('<span class="sp"></span>処理中...');
  showResult('','');
  try {
    const r = await fetch('/api/process', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ xlsxPath: paths[1], odsPath: paths[2] })
    });
    const d = await r.json();
    if (d.success) showResult('ok', '✓ 完了: ' + d.count + ' 行に制球を書き込みました');
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
    const filter = url.searchParams.get('filter') || 'all';
    const filterStr = filter === 'xlsx'
      ? 'Excel Files (*.xlsx)|*.xlsx'
      : filter === 'ods'
        ? 'ODS Files (*.ods)|*.ods'
        : 'All Files (*.*)|*.*';
    const fp = browseFile(filterStr);
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    return res.end(JSON.stringify({ path: fp }));
  }

  if (req.method === 'POST' && url.pathname === '/api/process') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', async () => {
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
      try {
        const { xlsxPath, odsPath } = JSON.parse(body);
        if (!xlsxPath) throw new Error('Excelファイルパスが指定されていません');
        if (!odsPath)  throw new Error('守備.odsパスが指定されていません');
        const count = await processFile(xlsxPath, odsPath);
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
