const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const types = { '.html':'text/html; charset=utf-8', '.js':'text/javascript; charset=utf-8', '.css':'text/css; charset=utf-8', '.json':'application/json', '.svg':'image/svg+xml', '.png':'image/png' };
const publicFiles = new Set(['index.html','app.js','cartones.js','experience.js','config.js','style.css','premium.css','manifest.json','sw.js']);
http.createServer((req, res) => {
  let relative;
  try { relative = decodeURIComponent(new URL(req.url, 'http://localhost').pathname).replace(/^\/+/, '').replace(/^RIFA-PRO\.COM\//, '') || 'index.html'; }
  catch { res.writeHead(400); return res.end(); }
  if (!publicFiles.has(relative) && !/^images\/[a-z0-9-]+\.(png|svg)$/.test(relative)) { res.writeHead(404); return res.end(); }
  fs.readFile(path.join(root, relative), (error, data) => {
    if (error) { res.writeHead(404); return res.end(); }
    res.writeHead(200, {'Content-Type':types[path.extname(relative)] || 'application/octet-stream', 'Cache-Control':'no-cache'}); res.end(data);
  });
}).listen(4173, '127.0.0.1', () => console.log('Rifa Pro: http://127.0.0.1:4173'));
