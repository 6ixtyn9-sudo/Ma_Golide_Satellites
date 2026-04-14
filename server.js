const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 5000;
const DOCS_DIR = path.join(__dirname, 'docs');

const FILES = [
  { name: 'Sheet_Setup.gs', label: 'Sheet Setup (Module 1)' },
  { name: 'Config_Ledger_Satellite.gs', label: 'Config Ledger Satellite' },
  { name: 'Contract_Enforcer.gs', label: 'Contract Enforcer' },
  { name: 'Contract_Enforcement.gs', label: 'Contract Enforcement' },
  { name: 'Data_Parser.gs', label: 'Data Parser' },
  { name: 'Game_Processor.gs', label: 'Game Processor (Module 7)' },
  { name: 'Forecaster.gs', label: 'Forecaster' },
  { name: 'Signal_Processor.gs', label: 'Signal Processor' },
  { name: 'Accumulator_Builder.gs', label: 'Accumulator Builder' },
  { name: 'Inventory_Manager.gs', label: 'Inventory Manager' },
  { name: 'Margin_Analyzer.gs', label: 'Margin Analyzer' },
];

function escapeHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function getFileSizes() {
  const sizes = {};
  for (const f of FILES) {
    try {
      const stat = fs.statSync(path.join(DOCS_DIR, f.name));
      sizes[f.name] = stat.size;
    } catch {
      sizes[f.name] = 0;
    }
  }
  return sizes;
}

function renderIndex() {
  const sizes = getFileSizes();
  const totalLines = FILES.reduce((acc, f) => {
    try {
      const content = fs.readFileSync(path.join(DOCS_DIR, f.name), 'utf8');
      return acc + content.split('\n').length;
    } catch {
      return acc;
    }
  }, 0);

  const rows = FILES.map(f => {
    const kb = (sizes[f.name] / 1024).toFixed(1);
    return `
      <tr>
        <td><a href="/view/${encodeURIComponent(f.name)}">${f.label}</a></td>
        <td class="mono">${f.name}</td>
        <td class="right">${kb} KB</td>
      </tr>`;
  }).join('');

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>MaGolide Betting System</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Segoe UI', system-ui, sans-serif; background: #0d1117; color: #c9d1d9; min-height: 100vh; }
    header { background: linear-gradient(135deg, #161b22, #1f2937); border-bottom: 1px solid #30363d; padding: 2rem; }
    header h1 { font-size: 2rem; font-weight: 700; color: #58a6ff; letter-spacing: -0.5px; }
    header p { margin-top: 0.5rem; color: #8b949e; font-size: 1rem; }
    .badge { display: inline-block; background: #388bfd26; color: #58a6ff; border: 1px solid #388bfd; border-radius: 999px; font-size: 0.75rem; padding: 2px 10px; margin-top: 0.75rem; }
    main { max-width: 900px; margin: 2.5rem auto; padding: 0 1.5rem; }
    .stats { display: flex; gap: 1.5rem; margin-bottom: 2rem; flex-wrap: wrap; }
    .stat-card { background: #161b22; border: 1px solid #30363d; border-radius: 8px; padding: 1rem 1.5rem; flex: 1; min-width: 140px; }
    .stat-card .value { font-size: 1.6rem; font-weight: 700; color: #58a6ff; }
    .stat-card .label { font-size: 0.8rem; color: #8b949e; margin-top: 2px; }
    table { width: 100%; border-collapse: collapse; background: #161b22; border: 1px solid #30363d; border-radius: 8px; overflow: hidden; }
    th { background: #1f2937; color: #8b949e; font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.05em; padding: 0.75rem 1rem; text-align: left; border-bottom: 1px solid #30363d; }
    td { padding: 0.75rem 1rem; border-bottom: 1px solid #21262d; font-size: 0.9rem; }
    tr:last-child td { border-bottom: none; }
    tr:hover td { background: #1f2937; }
    a { color: #58a6ff; text-decoration: none; font-weight: 500; }
    a:hover { text-decoration: underline; }
    .mono { font-family: monospace; color: #8b949e; font-size: 0.82rem; }
    .right { text-align: right; color: #8b949e; }
    .info-box { background: #161b22; border: 1px solid #30363d; border-left: 3px solid #58a6ff; border-radius: 4px; padding: 1rem 1.25rem; margin-bottom: 2rem; font-size: 0.875rem; color: #8b949e; line-height: 1.6; }
    .info-box strong { color: #c9d1d9; }
  </style>
</head>
<body>
  <header>
    <h1>MaGolide Betting System</h1>
    <p>Advanced sports betting prediction system — Google Apps Script modules</p>
    <span class="badge">Google Apps Script</span>
  </header>
  <main>
    <div class="info-box">
      <strong>How to use:</strong> These modules run inside Google Sheets via Google Apps Script. 
      Deploy them to a Google Sheets project using 
      <a href="https://github.com/google/clasp" target="_blank">CLASP</a> or paste them directly 
      into the Apps Script editor. Run <code>setupAllSheets()</code> first to initialize the spreadsheet.
    </div>
    <div class="stats">
      <div class="stat-card"><div class="value">${FILES.length}</div><div class="label">Script Modules</div></div>
      <div class="stat-card"><div class="value">${totalLines.toLocaleString()}</div><div class="label">Total Lines</div></div>
      <div class="stat-card"><div class="value">GAS</div><div class="label">Runtime</div></div>
    </div>
    <table>
      <thead>
        <tr>
          <th>Module</th>
          <th>File</th>
          <th class="right">Size</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  </main>
</body>
</html>`;
}

function renderFile(filename) {
  const file = FILES.find(f => f.name === filename);
  if (!file) return null;
  let content;
  try {
    content = fs.readFileSync(path.join(DOCS_DIR, filename), 'utf8');
  } catch {
    return null;
  }
  const lines = content.split('\n');
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>${file.label} — MaGolide</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Segoe UI', system-ui, sans-serif; background: #0d1117; color: #c9d1d9; min-height: 100vh; }
    header { background: #161b22; border-bottom: 1px solid #30363d; padding: 1rem 2rem; display: flex; align-items: center; gap: 1.5rem; }
    header a { color: #58a6ff; text-decoration: none; font-size: 0.9rem; }
    header a:hover { text-decoration: underline; }
    header h1 { font-size: 1.1rem; font-weight: 600; color: #c9d1d9; }
    header .meta { font-size: 0.8rem; color: #8b949e; margin-left: auto; }
    .code-wrap { max-width: 1200px; margin: 2rem auto; padding: 0 1.5rem; }
    .code-block { background: #161b22; border: 1px solid #30363d; border-radius: 8px; overflow: auto; }
    table.code-table { width: 100%; border-collapse: collapse; }
    .line-num { color: #484f58; text-align: right; padding: 0 12px; user-select: none; font-size: 0.8rem; font-family: monospace; min-width: 50px; border-right: 1px solid #21262d; }
    .line-code { padding: 0 16px; font-family: 'Fira Code', 'Consolas', monospace; font-size: 0.82rem; white-space: pre; color: #c9d1d9; }
    tr:hover .line-num, tr:hover .line-code { background: #1c2128; }
  </style>
</head>
<body>
  <header>
    <a href="/">&#8592; Back</a>
    <h1>${file.label}</h1>
    <span class="meta">${lines.length.toLocaleString()} lines &bull; ${filename}</span>
  </header>
  <div class="code-wrap">
    <div class="code-block">
      <table class="code-table">
        <tbody>
          ${lines.map((line, i) => `<tr><td class="line-num">${i + 1}</td><td class="line-code">${escapeHtml(line)}</td></tr>`).join('')}
        </tbody>
      </table>
    </div>
  </div>
</body>
</html>`;
}

const server = http.createServer((req, res) => {
  const url = req.url;

  if (url === '/' || url === '') {
    const html = renderIndex();
    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(html);
    return;
  }

  const viewMatch = url.match(/^\/view\/(.+)$/);
  if (viewMatch) {
    const filename = decodeURIComponent(viewMatch[1]);
    const html = renderFile(filename);
    if (html) {
      res.writeHead(200, { 'Content-Type': 'text/html' });
      res.end(html);
    } else {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('File not found');
    }
    return;
  }

  res.writeHead(404, { 'Content-Type': 'text/plain' });
  res.end('Not found');
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`MaGolide Betting System running at http://0.0.0.0:${PORT}`);
});
