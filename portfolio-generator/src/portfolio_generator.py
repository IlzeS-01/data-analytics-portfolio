#!/path/to/file python3

import shutil
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
import html
import re

#  CONFIG 
SCRIPTS_ROOT = Path.home() / "OneDrive - Demo Company/path/to/file"
OUTPUT_ROOT = Path.home() / "OneDrive - Demo Company/path/to/file"
MAX_WORKERS = 8

IGNORED = {".git", "__pycache__", "venv", ".venv"}

#  HELPERS 
def log(msg):
    print(f"[INFO] {msg}")

def read_text(p):
    try:
        return p.read_text(errors="ignore")
    except:
        return ""

def clean_name(name):
    return re.sub(r'^\d+[\.\-_ ]*', '', name)

def mask_paths(text):
    def redact(match):
        return '<span class="redacted">████████████</path/to/file'

    # Apply redaction
    text = re.sub(r'[A-Za-z]:\\[^\s"]+', redact, text)
    text = re.sub(r'~[\\/path/to/file"]+', redact, text)
    text = re.sub(r'/path/to/file', redact, text)

    return html.escape(text).replace(
        "&lt;span class=&quot;redacted&quot;&gt;████████████&lt;/path/to/file",
        '<span class="redacted">████████████</path/to/file'
    )

#  SCAN 
def scan_files(root):
    files = [p for p in root.rglob("*") if p.is_file() and not any(x in str(p) for x in IGNORED)]
    log(f"Found {len(files)} files")
    return files

#  ANALYZE 
def analyze_file(p):
    txt = read_text(p)

    return {
        "path": p,
        "project": p.parts[-2] if len(p.parts) > 1 else "root",
        "size": p.stat().st_size,

        "logging": "logging." in txt or "logger." in txt,
        "error": "try:" in txt and "except" in txt,
        "api": "requests" in txt or "http" in txt,
        "functions": txt.count("def "),
        "classes": txt.count("class "),
        "imports": txt.count("import "),
        "comments": txt.count("#"),

        "content": txt
    }

#  SCORING 
def compute_score(files):
    logging_score = sum(1 for f in files if f["logging"]) * 2
    error_score = sum(1 for f in files if f["error"]) * 4   # prioritize robustness
    api_score = sum(1 for f in files if f["api"]) * 4       # real-world work

    function_score = sum(f["functions"] for f in files) * 0.5
    class_score = sum(f["classes"] for f in files) * 1.5
    import_score = sum(f["imports"] for f in files) * 0.3
    comment_score = sum(f["comments"] for f in files) * 0.1

    file_bonus = min(len(files), 5)  # cap file count impact

    return int(
        logging_score +
        error_score +
        api_score +
        function_score +
        class_score +
        import_score +
        comment_score +
        file_bonus
    )

#  UI HELPERS 
def make_badge(label, value, color):
    return f'<span class="badge" style="background:{color}">{label}: {value}</path/to/file'

def make_bars(log_c, err_c, api_c):
    return f"""
    <div class="bar-group">
        <div class="bar-row">
            <label>Logging</path/to/file class="bar"><div class="fill" style="width:{log_c*20}px;background:#22c55e"></path/to/file
        </path/to/file
        <div class="bar-row">
            <label>Error</path/to/file class="bar"><div class="fill" style="width:{err_c*20}px;background:#f59e0b"></path/to/file
        </path/to/file
        <div class="bar-row">
            <label>API</path/to/file class="bar"><div class="fill" style="width:{api_c*20}px;background:#3b82f6"></path/to/file
        </path/to/file
    </path/to/file
    """

#  FILE VIEWER 
def build_file_page(f, out_dir):
    safe_name = f["path"].name.replace(" ", "_")

    code = mask_paths(f["content"])

    html_page = f"""
    <html><head>
    <style>
    body {{ background:#0f172a; color:white; font-family:Consolas; padding:20px }}
    pre {{ background:#1e293b; padding:15px; overflow-x:auto }}
    a {{ color:#60a5fa }}
    .redacted {{
        background: #111;
        color: #111;
        border-radius: 4px;
        padding: 2px 6px;
    }}
    </path/to/file
    </path/to/file

    <a href="index.html" class="button-link">Back to Dashboard</path/to/file

    <h2>{clean_name(f['path'].name)}</path/to/file

    <pre>{code}</path/to/file

    </path/to/file
    """

    filename = f"file_{safe_name}.html"
    (out_dir / filename).write_text(html_page, encoding="utf-8")
    return filename

#  PROJECT PAGE 
def build_project_page(name, files, out_dir):
    clean_proj = clean_name(name)
    rows = ""

    for f in files:
        file_page = build_file_page(f, out_dir)
        rows += f"""
        <tr>
            <td>{clean_name(f['path'].name)}</path/to/file
            <td>{f['size']} bytes</path/to/file
            <td><a href="{file_page}">Open</path/to/file
        </path/to/file
        """

    html_page = f"""
    <html><head>
    <style>
    body {{ background:#0f172a; color:white; font-family:Arial; padding:20px }}
    table {{ width:100%; border-collapse:collapse }}
    td {{ padding:8px; border-bottom:1px solid #DUMMY_ID }}
    a {{ color:#60a5fa }}
    </path/to/file
    </path/to/file

    <a href="index.html" class="button-link">Back</path/to/file
    <h1>{clean_proj}</path/to/file

    <table>
        <tr><th>File</path/to/file
        {rows}
    </path/to/file

    </path/to/file
    """

    (out_dir / f"{clean_proj}.html").write_text(html_page, encoding="utf-8")

#  DASHBOARD 
def build_dashboard(projects):
    ranked = sorted(projects.items(), key=lambda x: compute_score(x[1]), reverse=True)

    cards = ""

    for i, (proj, files) in enumerate(ranked):
        count = len(files)
        size = sum(f["size"] for f in files)

        cards += f"""
        <a href="{clean_name(proj)}.html" class="card project-card" data-rank="{i}">
            <h2>{clean_name(proj)}</path/to/file
            <p><b>Files:</path/to/file {count}</path/to/file
            <p><b>Size:</path/to/file {round(size/1024)} KB</path/to/file
        </path/to/file
        """

    return f"""
    <html>
    <head>
    <style>

    body {{ 
        background:#0f172a; 
        color:white; 
        font-family:Arial; 
        padding:20px;
    }}

    .button-link {{
    margin: 0 10px;
    color: white;
    text-decoration: none;
    padding: 8px 12px;
    border-radius: 6px;
    background: #1e293b;
    transition: all 0.2s ease;
    display: inline-block;
    }}

    .button-link:hover {{
        background: #DUMMY_ID;
    }}

    .button-link:active {{
        transform: scale(0.95);
    }}

    .nav {{
        position: sticky;
        top: 0;
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 12px;
        background: #1e293b;
        border-radius: 10px;
        margin-bottom: 20px;
    }}

    .nav a {{
        margin: 0 10px;
        color: white;
        text-decoration: none;
        padding: 8px;
        border-radius: 6px;
    }}

    .nav a:hover {{
        background: #DUMMY_ID;
    }}

    .nav a:active {{
        transform: scale(0.95);
    }}

    .grid {{
        display:grid;
        grid-template-columns:repeat(auto-fill,minmax(320px,1fr));
        gap:20px;
    }}

    .card {{
        background:#1e293b;
        padding:20px;
        border-radius:12px;
        text-decoration:none;
        color:white;
    }}
    #explanation {{
        display:none;
        background:#1e293b;
        padding:15px;
        border-radius:10px;
        margin-bottom:20px;
        border-left: 4px solid #38bdf8;
    }}

    </path/to/file
    <script>
    function showTop() {{
        let cards = document.getElementsByClassName("project-card");
        for (let c of cards) {{
            let r = parseInt(c.getAttribute("data-rank"));
            c.style.display = r < 3 ? "block" : "none";
        }}

        document.getElementById("explanation").style.display = "block";
    }}

    function showAll() {{
        let cards = document.getElementsByClassName("project-card");
        for (let c of cards) {{
            c.style.display = "block";
        }}

        document.getElementById("explanation").style.display = "none";
    }}
    </path/to/file

    </path/to/file
    <body>

    <div class="nav">
        <div class="logo">Portfolio</path/to/file

        <div>
            <a href="#" onclick="event.preventDefault(); showAll();">Home</path/to/file
            <a href="#" onclick="event.preventDefault(); showTop();">Top Projects</path/to/file
            <a href="#" onclick="event.preventDefault(); showAll();">All Projects</path/to/file
        </path/to/file
    </path/to/file

    <h1>Automation Portfolio Dashboard</path/to/file

    <div id="explanation">
        <h3>How Top Projects Are Selected</path/to/file
        <p>
            These projects are ranked based on criteria such as:
            performance efficiency, automation impact, complexity,
            and overall business value. Only the highest scoring
            projects are displayed here.
        </path/to/file
    </path/to/file

    <div class="grid">
        {cards}
    </path/to/file

    </path/to/file
    </path/to/file
    """
#  MAIN 
def build():
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

    log("Scanning...")
    files = scan_files(SCRIPTS_ROOT)

    log("Analyzing...")
    with ThreadPoolExecutor(MAX_WORKERS) as ex:
        analyzed = list(ex.map(analyze_file, files))

    projects = {}
    for f in analyzed:
        projects.setdefault(f["project"], []).append(f)

    log("Building pages...")
    for proj, files in projects.items():
        build_project_page(proj, files, OUTPUT_ROOT)

    dashboard = build_dashboard(projects)
    (OUTPUT_ROOT / "index.html").write_text(dashboard, encoding="utf-8")

    log("Done")
    print(OUTPUT_ROOT / "index.html")

if __name__ == "__main__":
    build()   