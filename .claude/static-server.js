// ローカル静的配信サーバー (プレビュー用・開発専用)
const http = require('http'), fs = require('fs'), path = require('path');
const root = process.argv[2] || '.';
const port = parseInt(process.argv[3] || '3456', 10);
const types = { '.html':'text/html; charset=utf-8', '.js':'text/javascript; charset=utf-8', '.css':'text/css; charset=utf-8', '.json':'application/json; charset=utf-8', '.png':'image/png', '.jpg':'image/jpeg', '.jpeg':'image/jpeg', '.gif':'image/gif', '.svg':'image/svg+xml', '.ico':'image/x-icon', '.woff':'font/woff', '.woff2':'font/woff2' };
http.createServer((req, res) => {
  let p = decodeURIComponent(req.url.split('?')[0]);
  if (p === '/' || p === '') p = '/card_generator.html';
  const rootAbs = path.resolve(root);
  const fp = path.resolve(rootAbs, '.' + p);
  if (!fp.startsWith(rootAbs)) { res.writeHead(403); res.end('forbidden'); return; }
  fs.readFile(fp, (e, d) => {
    if (e) { res.writeHead(404); res.end('not found'); return; }
    res.writeHead(200, { 'Content-Type': types[path.extname(fp).toLowerCase()] || 'application/octet-stream', 'Cache-Control': 'no-store' });
    res.end(d);
  });
}).listen(port, () => console.log('serving ' + path.resolve(root) + ' on http://localhost:' + port));
