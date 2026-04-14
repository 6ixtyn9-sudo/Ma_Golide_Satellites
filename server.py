import os
import json
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

DOCS_DIR = os.path.join(os.path.dirname(__file__), "docs")

SCRIPT_INFO = {
    "Sheet_Setup.gs": {
        "module": "Module 1",
        "title": "Sheet Setup",
        "role": "The Architect",
        "description": "One-click creation of the full Ma Golide sheet infrastructure. The Single Source of Truth for sheet names used across all modules."
    },
    "Signal_Processor.gs": {
        "module": "Module 3",
        "title": "Signal Processor",
        "role": "The Mouth",
        "description": "Standardizes messy, vertically-formatted raw data from FlashScore/Sofascore into horizontal Clean sheets."
    },
    "Data_Parser.gs": {
        "module": "Module 2",
        "title": "Data Parser",
        "role": "The Parser",
        "description": "Parses and transforms raw incoming data before processing."
    },
    "Forecaster.gs": {
        "module": "Module 5",
        "title": "Forecaster",
        "role": "The Brain (Tier 1)",
        "description": "Enriches upcoming games with historical context, rankings, and streaks for Tier 1 predictions."
    },
    "Game_Processor.gs": {
        "module": "Module 7",
        "title": "Game Processor",
        "role": "The Engine (Tier 2)",
        "description": "Production-ready prediction engine handling complex probability models, Tier 2 O/U predictions, and confidence scoring."
    },
    "Accumulator_Builder.gs": {
        "module": "Module 6",
        "title": "Accumulator Builder",
        "role": "The Stacker",
        "description": "Constructs multi-bet parlays and accumulators based on high-confidence signals."
    },
    "Inventory_Manager.gs": {
        "module": "Module 4",
        "title": "Inventory Manager",
        "role": "The Librarian",
        "description": "Manages data state and historical records across the system."
    },
    "Margin_Analyzer.gs": {
        "module": "Module 8",
        "title": "Margin Analyzer",
        "role": "The Analyst",
        "description": "Focuses on specific betting markets: Spreads, Moneylines, and Totals."
    },
    "Contract_Enforcer.gs": {
        "module": "Module 9",
        "title": "Contract Enforcer",
        "role": "The Validator",
        "description": "Validates data integrity and ensures modules adhere to the Single Source of Truth."
    },
    "Contract_Enforcement.gs": {
        "module": "Module 9b",
        "title": "Contract Enforcement",
        "role": "The Guard",
        "description": "Enforces contracts and data validation rules across the pipeline."
    },
    "Config_Ledger_Satellite.gs": {
        "module": "Config",
        "title": "Config Ledger Satellite",
        "role": "The Registry",
        "description": "Satellite configuration ledger for system-wide settings and constants."
    },
}

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MaGolide Betting System</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', system-ui, sans-serif; background: #0d1117; color: #e6edf3; min-height: 100vh; }}
  .header {{ background: linear-gradient(135deg, #1f6feb 0%, #388bfd 100%); padding: 32px 40px; }}
  .header h1 {{ font-size: 2rem; font-weight: 700; letter-spacing: -0.5px; }}
  .header p {{ color: #cce3ff; margin-top: 6px; font-size: 1rem; }}
  .badge {{ display: inline-block; background: rgba(255,255,255,0.2); border-radius: 20px; padding: 4px 12px; font-size: 0.75rem; margin-top: 10px; }}
  .container {{ max-width: 1200px; margin: 0 auto; padding: 32px 24px; }}
  .section-title {{ font-size: 0.75rem; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; color: #8b949e; margin-bottom: 16px; margin-top: 32px; }}
  .grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap: 16px; }}
  .card {{ background: #161b22; border: 1px solid #30363d; border-radius: 10px; padding: 20px; cursor: pointer; transition: all 0.2s; text-decoration: none; color: inherit; display: block; }}
  .card:hover {{ border-color: #388bfd; transform: translateY(-2px); box-shadow: 0 8px 24px rgba(56,139,253,0.15); }}
  .card-module {{ font-size: 0.7rem; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: #388bfd; margin-bottom: 6px; }}
  .card-title {{ font-size: 1.05rem; font-weight: 600; margin-bottom: 4px; }}
  .card-role {{ font-size: 0.8rem; color: #f0883e; margin-bottom: 10px; font-style: italic; }}
  .card-desc {{ font-size: 0.85rem; color: #8b949e; line-height: 1.5; }}
  .card-footer {{ margin-top: 14px; padding-top: 12px; border-top: 1px solid #21262d; display: flex; align-items: center; gap: 8px; }}
  .card-lines {{ font-size: 0.75rem; color: #8b949e; }}
  .view-btn {{ margin-left: auto; background: #21262d; border: 1px solid #30363d; color: #cdd9e5; font-size: 0.75rem; padding: 4px 10px; border-radius: 6px; cursor: pointer; }}
  .info-box {{ background: #161b22; border: 1px solid #30363d; border-radius: 10px; padding: 24px; margin-bottom: 24px; }}
  .info-box h3 {{ font-size: 0.9rem; color: #388bfd; margin-bottom: 10px; }}
  .info-box p {{ font-size: 0.85rem; color: #8b949e; line-height: 1.6; }}
  .steps {{ list-style: none; margin-top: 10px; }}
  .steps li {{ font-size: 0.85rem; color: #8b949e; padding: 4px 0; padding-left: 20px; position: relative; }}
  .steps li::before {{ content: attr(data-n); position: absolute; left: 0; color: #388bfd; font-weight: 700; }}
  .footer {{ text-align: center; padding: 40px 24px; color: #484f58; font-size: 0.8rem; border-top: 1px solid #21262d; margin-top: 40px; }}
</style>
</head>
<body>
<div class="header">
  <h1>MaGolide Betting System</h1>
  <p>Advanced sports betting prediction and analysis platform</p>
  <span class="badge">Google Apps Script &bull; {script_count} Modules</span>
</div>
<div class="container">
  <div class="info-box">
    <h3>How to use these scripts</h3>
    <p>MaGolide is a Google Apps Script project designed to run inside Google Sheets. Follow these steps to deploy it:</p>
    <ol class="steps">
      <li data-n="1.">Open a new Google Spreadsheet and go to <strong>Extensions &rarr; Apps Script</strong></li>
      <li data-n="2.">Create a new script file for each module below (use the filename as the script name)</li>
      <li data-n="3.">Paste the contents of each module into its corresponding script file</li>
      <li data-n="4.">Save the project, then run <code>setupAllSheets()</code> from Sheet_Setup.gs first</li>
      <li data-n="5.">Use the custom menu or run functions manually in the correct pipeline order</li>
    </ol>
  </div>

  <div class="section-title">Pipeline Modules</div>
  <div class="grid">
    {cards}
  </div>
</div>
<div class="footer">MaGolide Betting System &mdash; Google Apps Script Modules Viewer</div>
</body>
</html>"""

VIEWER_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title} &mdash; MaGolide</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', system-ui, sans-serif; background: #0d1117; color: #e6edf3; min-height: 100vh; display: flex; flex-direction: column; }}
  .toolbar {{ background: #161b22; border-bottom: 1px solid #30363d; padding: 12px 24px; display: flex; align-items: center; gap: 16px; }}
  .back {{ color: #388bfd; text-decoration: none; font-size: 0.9rem; display: flex; align-items: center; gap: 6px; }}
  .back:hover {{ text-decoration: underline; }}
  .toolbar-title {{ font-weight: 600; font-size: 0.95rem; }}
  .toolbar-module {{ font-size: 0.75rem; color: #388bfd; margin-left: 8px; }}
  .copy-btn {{ margin-left: auto; background: #1f6feb; color: white; border: none; padding: 6px 16px; border-radius: 6px; cursor: pointer; font-size: 0.85rem; font-weight: 600; transition: background 0.2s; }}
  .copy-btn:hover {{ background: #388bfd; }}
  .copy-btn.copied {{ background: #238636; }}
  .meta {{ background: #161b22; border-bottom: 1px solid #30363d; padding: 12px 24px; display: flex; gap: 24px; }}
  .meta-item {{ font-size: 0.8rem; color: #8b949e; }}
  .meta-item strong {{ color: #f0883e; }}
  .code-wrap {{ flex: 1; overflow: auto; }}
  pre {{ padding: 24px; font-family: 'Cascadia Code', 'Fira Code', 'Consolas', monospace; font-size: 0.82rem; line-height: 1.6; color: #c9d1d9; white-space: pre; tab-size: 2; min-height: 100%; }}
  .line-num {{ color: #484f58; user-select: none; margin-right: 16px; display: inline-block; min-width: 40px; text-align: right; }}
</style>
</head>
<body>
<div class="toolbar">
  <a href="/" class="back">&#8592; Back</a>
  <span class="toolbar-title">{title}</span>
  <span class="toolbar-module">{module}</span>
  <button class="copy-btn" onclick="copyCode()">Copy All</button>
</div>
<div class="meta">
  <div class="meta-item">Role: <strong>{role}</strong></div>
  <div class="meta-item">Lines: <strong>{lines}</strong></div>
  <div class="meta-item">File: <strong>{filename}</strong></div>
</div>
<div class="code-wrap">
<pre id="code">{code}</pre>
</div>
<script>
function copyCode() {{
  const text = {raw_json};
  navigator.clipboard.writeText(text).then(() => {{
    const btn = document.querySelector('.copy-btn');
    btn.textContent = 'Copied!';
    btn.classList.add('copied');
    setTimeout(() => {{ btn.textContent = 'Copy All'; btn.classList.remove('copied'); }}, 2000);
  }});
}}
</script>
</body>
</html>"""


def get_file_lines(filename):
    path = os.path.join(DOCS_DIR, filename)
    try:
        with open(path, 'r') as f:
            return len(f.readlines())
    except:
        return 0


def escape_html(text):
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def build_index():
    cards = []
    files = sorted(os.listdir(DOCS_DIR)) if os.path.isdir(DOCS_DIR) else []
    gs_files = [f for f in files if f.endswith('.gs')]
    
    for filename in gs_files:
        info = SCRIPT_INFO.get(filename, {
            "module": "Module",
            "title": filename.replace('.gs', '').replace('_', ' '),
            "role": "Script",
            "description": "Google Apps Script module."
        })
        lines = get_file_lines(filename)
        card = f"""
    <a class="card" href="/view?file={filename}">
      <div class="card-module">{info['module']}</div>
      <div class="card-title">{info['title']}</div>
      <div class="card-role">{info['role']}</div>
      <div class="card-desc">{info['description']}</div>
      <div class="card-footer">
        <span class="card-lines">{lines:,} lines</span>
        <span class="view-btn">View &rarr;</span>
      </div>
    </a>"""
        cards.append(card)
    
    return HTML_TEMPLATE.format(
        script_count=len(gs_files),
        cards='\n'.join(cards)
    )


def build_viewer(filename):
    path = os.path.join(DOCS_DIR, filename)
    if not os.path.isfile(path) or not filename.endswith('.gs'):
        return None
    
    info = SCRIPT_INFO.get(filename, {
        "module": "Module",
        "title": filename.replace('.gs', '').replace('_', ' '),
        "role": "Script",
        "description": ""
    })
    
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    lines = content.split('\n')
    numbered = '\n'.join(
        f'<span class="line-num">{i+1}</span>{escape_html(line)}'
        for i, line in enumerate(lines)
    )
    
    return VIEWER_TEMPLATE.format(
        title=info['title'],
        module=info['module'],
        role=info['role'],
        filename=filename,
        lines=f"{len(lines):,}",
        code=numbered,
        raw_json=json.dumps(content)
    )


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def send_html(self, content, status=200):
        encoded = content.encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.send_header('Content-Length', str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == '/' or path == '/index.html':
            self.send_html(build_index())
        elif path == '/view':
            qs = parse_qs(parsed.query)
            filename = qs.get('file', [''])[0]
            content = build_viewer(filename)
            if content:
                self.send_html(content)
            else:
                self.send_html('<h1>Not found</h1>', 404)
        else:
            self.send_html('<h1>Not found</h1>', 404)


if __name__ == '__main__':
    port = 5000
    server = HTTPServer(('0.0.0.0', port), Handler)
    print(f"MaGolide viewer running on port {port}")
    server.serve_forever()
