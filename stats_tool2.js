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

// ── 制球計算式 (守備.ods AC2 の等価実装) ─────────────────────────────────────
// 守備.ods AC2 の数式:
//   IFERROR(IFS(V2>=4.2, ROUND(60-(V2-4.2)/0.16),
//               V2>=1.2, ROUND(85-(V2-1.2)/0.12),
//               V2>=0,   ROUND(100-V2/0.08)), "")
// V2 = 四死球 / 換算イニング × 9  (BB9)
function calcSeikyuFromBB9(bb9) {
  if (bb9 == null || isNaN(bb9) || bb9 < 0) return '';
  if (bb9 >= 4.2) return Math.round(60 - (bb9 - 4.2) / 0.16);
  if (bb9 >= 1.2) return Math.round(85 - (bb9 - 1.2) / 0.12);
  return Math.round(100 - bb9 / 0.08);
}

// ── セルスタイル ─────────────────────────────────────────────────────────────
const PURPLE_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7030A0' } };
function purpleCell(cell, value, fs) {
  cell.value = value;
  cell.fill  = { ...PURPLE_FILL };
  cell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: fs };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
}

// 制球列: Col 51 = AY (22 主要成績 + 7×4=28 球種列 + 1)
const SEIKYU_COL = 51;

// ── メイン処理 ───────────────────────────────────────────────────────────────
async function processFile(xlsxPath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(xlsxPath);
  const ws = wb.worksheets[0];
  const fontSize = ws.getCell(1, 1).font?.size || 11;

  // ヘッダー書き込み（Row 1 / Col 51）
  purpleCell(ws.getCell(1, SEIKYU_COL), '制球', fontSize);

  const dataRows = [];
  ws.eachRow((row, rn) => {
    if (rn <= 2) return;            // Row1=大項目, Row2=小項目ヘッダーをスキップ
    const yr = row.getCell(2).value;
    if (!yr) return;
    const ipRaw = row.getCell(11).value; // K列: イニング
    if (ipRaw == null || ipRaw === '' || ipRaw === '--') return;
    const bb = Number(row.getCell(15).value) || 0; // O列: 四死球
    dataRows.push({ rn, ipRaw, bb });
  });

  let count = 0;
  for (const { rn, ipRaw, bb } of dataRows) {
    const ip = parseIP(ipRaw);
    if (!ip) continue;

    // V2 = 四死球 / 換算イニング × 9
    const bb9 = bb / ip * 9;

    // AC2 等価計算
    const seikyu = calcSeikyuFromBB9(bb9);
    if (seikyu === '') continue;

    purpleCell(ws.getCell(rn, SEIKYU_COL), seikyu, fontSize);
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
  padding:32px;width:100%;max-width:540px}
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
  padding:10px 14px;font-size:11px;color:#555;line-height:1.8;margin-top:10px}
.formula-box code{font-family:monospace;color:#7030A0;font-size:12px}
</style>
</head>
<body>
<div class="card">
  <h1>MLB投手成績 能力値追加ツール</h1>
  <div class="subtitle">投手成績.xlsx の四死球・イニングから <strong>制球</strong> を算出し AY 列へ追加します</div>

  <div class="sec">
    <div class="sec-label">追加される列</div>
    <span class="badge">AY 制球</span>
    <div class="formula-box">
      BB9 = <code>四死球(O列) ÷ 換算イニング(K列) × 9</code><br>
      制球 = 守備.ods AC2 相当<br>
      　BB9 ≥ 4.2 → <code>round(60 − (BB9−4.2) / 0.16)</code><br>
      　BB9 ≥ 1.2 → <code>round(85 − (BB9−1.2) / 0.12)</code><br>
      　BB9 ≥ 0　→ <code>round(100 − BB9 / 0.08)</code>
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
      <button class="btn-run" id="btnRun" disabled onclick="run()">▶ 制球を追加</button>
      <span class="status" id="status"></span>
    </div>
    <div class="result" id="result"></div>
  </div>

  <div class="note">
    <strong>イニング換算</strong>: X.1 → X＋1/3、X.2 → X＋2/3<br>
    <strong>例</strong>: 200.1イニング・四死球70 → BB9 = 70÷200.33×9 ≈ 3.14 → 制球 = 72
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
