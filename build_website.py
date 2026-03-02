"""
Build The CountUp Rankings website from Excel data.

Generates docs/ folder for GitHub Pages hosting.
Design closely matches the TSC Update Project style:
black bg, white text, Courier New monospace, Georgia serif headings.

Usage: python build_website.py
Requires: xlwings (Excel must be open with the workbook)
"""
import xlwings as xw
import os, sys, io, json, glob, random
from html import escape
from PIL import Image, ImageOps
import numpy as np

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS_DIR = os.path.join(SCRIPT_DIR, "docs")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def clean(val):
    """Clean a cell value to a trimmed string. Returns '' for None/dot/None-str."""
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s in (".", "None") else s


def fmt_pct(val):
    """Format a 0-1 float as a percentage string like '92%'."""
    if val is None or not isinstance(val, (int, float)):
        return ""
    return f"{int(round(val * 100))}%"


def fmt_score(val):
    """Format a float score to one decimal."""
    if val is None or not isinstance(val, (int, float)):
        return ""
    return f"{val:.1f}"


# ---------------------------------------------------------------------------
# Data extraction — one function per sheet
# ---------------------------------------------------------------------------

def extract_pop(ws):
    lr = ws.used_range.last_cell.row
    lc = ws.used_range.last_cell.column
    data = ws.range((1, 1), (lr, lc)).value

    items = []
    for i, row in enumerate(data):
        if i + 1 < 3:
            continue
        # Check raw value for separator row (all dots)
        raw_b = row[1]
        if raw_b is not None and str(raw_b).strip() == ".":
            break
        brand = clean(row[1])
        flavor = clean(row[2])
        if not brand or not flavor:
            continue
        # Stop at tracking rows
        if brand in ("Avg", "Adds", "Offset", "Offset Owed"):
            break
        classics = clean(row[10]) if len(row) > 10 else ""
        sugar = row[3] if isinstance(row[3], (int, float)) else None
        sugar_str = str(int(sugar)) if sugar is not None else ""
        items.append({"brand": brand, "flavor": flavor, "sugar": sugar_str, "classics": classics})

    return {
        "id": "pop", "name": "Pop",
        "columns": [("#", "rank"), ("BRAND", "text"), ("FLAVOR", "text"), ("SUGAR(g)", "num"), ("CLASSICS", "text-dim")],
        "items": items, "sub_sections": None,
    }


def extract_candy(ws):
    lr = ws.used_range.last_cell.row
    lc = ws.used_range.last_cell.column
    data = ws.range((1, 1), (lr, lc)).value

    items = []
    for i, row in enumerate(data):
        if i + 1 < 3:
            continue
        country = clean(row[1]) if len(row) > 1 else ""
        brand = clean(row[2]) if len(row) > 2 else ""
        variety = clean(row[3]) if len(row) > 3 else ""
        candy_type = clean(row[4]) if len(row) > 4 else ""
        # Skip sub-items (no country/brand), Total rows, blanks
        if not country or not brand:
            continue
        if variety.lower() == "total" or brand.lower() == "total":
            continue
        items.append({
            "country": country, "brand": brand,
            "variety": variety, "type": candy_type,
        })

    return {
        "id": "candy", "name": "Candy",
        "columns": [("#", "rank"), ("COUNTRY", "text-sm"), ("BRAND", "text"), ("VARIETY", "text"), ("TYPE", "text-sm")],
        "items": items, "sub_sections": None,
    }


def extract_chocolate(ws):
    lr = ws.used_range.last_cell.row
    lc = ws.used_range.last_cell.column
    data = ws.range((1, 1), (lr, lc)).value

    items = []
    for i, row in enumerate(data):
        if i + 1 < 4:  # headers in row 3
            continue
        # Col A (0) is blank; data starts at index 1
        brand = clean(row[1])
        name = clean(row[2])
        choc_type = clean(row[3])
        if not brand or not name:
            continue
        if brand == "Top Chocolate":
            continue
        # Normalize type
        ct = choc_type.strip().capitalize() if choc_type else ""
        if ct.startswith("Dark"):
            ct = "Dark"
        elif ct.lower().startswith("milk") or ct == "MIlk":
            ct = "Milk"
        pct = row[4] if isinstance(row[4], (int, float)) else None
        g_sugar = row[5] if isinstance(row[5], (int, float)) else None
        sugar_str = str(int(round(g_sugar))) if g_sugar is not None else ""
        items.append({
            "brand": brand, "name": name,
            "type": ct, "cacao": fmt_pct(pct), "sugar": sugar_str,
        })

    return {
        "id": "chocolate", "name": "Chocolate",
        "columns": [("#", "rank"), ("BRAND", "text"), ("NAME", "text"), ("TYPE", "text-sm"), ("CACAO", "num"), ("SUGAR(g)", "num")],
        "items": items, "sub_sections": None,
    }


def extract_sauces(ws):
    lr = ws.used_range.last_cell.row
    lc = ws.used_range.last_cell.column
    data = ws.range((1, 1), (lr, lc)).value

    section_names = {"Asian Sauces", "Other Main Sauces", "Finishing Sauces"}
    sub_sections = []
    current = None

    for i, row in enumerate(data):
        if i + 1 < 3:
            continue
        # Col A (0) is blank; data starts at index 1
        col_b = clean(row[1])  # Origin for data, None for headers
        col_c = clean(row[2])  # Name for data, section name for headers

        # Section header: col B is blank, col C has section name
        if not col_b and col_c in section_names:
            current = {"name": col_c, "items": []}
            sub_sections.append(current)
            continue

        if current is None:
            continue

        origin = col_b
        name = col_c
        brand = clean(row[3])
        if not origin or not name or not brand:
            continue

        rice = clean(row[4])
        chicken = clean(row[5])
        rice = rice if rice else "\u2014"
        chicken = chicken if chicken else "\u2014"

        current["items"].append({
            "origin": origin, "name": name, "brand": brand,
            "rice": rice, "chicken": chicken,
        })

    return {
        "id": "sauces", "name": "Sauces",
        "columns": [("#", "rank"), ("ORIGIN", "text-sm"), ("NAME", "text"), ("BRAND", "text"), ("RICE", "grade"), ("CHICKEN", "grade")],
        "items": None, "sub_sections": sub_sections,
    }


def extract_dining(ws):
    lr = ws.used_range.last_cell.row
    lc = ws.used_range.last_cell.column
    data = ws.range((1, 1), (lr, lc)).value

    items = []
    for i, row in enumerate(data):
        if i + 1 < 3:
            continue
        rank_val = row[0]
        restaurant = clean(row[1])
        item = clean(row[2])
        rating = clean(row[3])
        if not restaurant or not item:
            continue
        rank = int(rank_val) if isinstance(rank_val, (int, float)) else None
        items.append({
            "rank": rank, "restaurant": restaurant,
            "item": item, "rating": rating,
        })

    return {
        "id": "dining", "name": "Dining",
        "columns": [("RANK", "rank"), ("RESTAURANT", "text"), ("ITEM", "text"), ("RATING", "grade")],
        "items": items, "sub_sections": None,
    }


def extract_games(ws):
    lr = ws.used_range.last_cell.row
    lc = ws.used_range.last_cell.column
    data = ws.range((1, 1), (lr, lc)).value

    items = []
    seen = set()
    for i, row in enumerate(data):
        if i + 1 < 3:
            continue
        rank_val = row[1]
        game = clean(row[2])
        if not game or not isinstance(rank_val, (int, float)):
            continue
        # Deduplicate by game name (keep first occurrence = higher ranked)
        if game in seen:
            continue
        seen.add(game)

        avg = row[6] if isinstance(row[6], (int, float)) else None
        year = int(row[8]) if len(row) > 8 and isinstance(row[8], (int, float)) else None

        items.append({
            "game": game, "year": year,
            "score": fmt_score(avg),
        })

    return {
        "id": "games", "name": "Games",
        "columns": [("#", "rank"), ("GAME", "text"), ("YEAR", "num"), ("SCORE", "num")],
        "items": items, "sub_sections": None,
    }


# ---------------------------------------------------------------------------
# CSS — shared across all pages, closely matching TSC Update Project
# ---------------------------------------------------------------------------

CSS = """
    * { margin: 0; padding: 0; box-sizing: border-box; }

    body {
      font-family: 'Courier New', Courier, monospace;
      font-weight: 700;
      background: #000;
      color: #fff;
      line-height: 1.4;
      padding: 20px;
    }

    .container {
      max-width: 1000px;
      margin: 0 auto;
      border: 3px solid #fff;
      padding: 0;
    }

    /* Header — white band */
    header {
      text-align: center;
      padding: 25px 20px 15px;
      background: #fff;
      color: #000;
      position: relative;
    }

    header h1 {
      font-family: Georgia, 'Times New Roman', serif;
      font-size: 2.8em;
      font-weight: 900;
      letter-spacing: 0.08em;
      margin-bottom: 2px;
    }

    header .subtitle {
      font-family: 'Courier New', Courier, monospace;
      font-size: 0.85em;
      font-weight: 400;
      letter-spacing: 0.15em;
      text-transform: uppercase;
      color: #555;
      margin-top: 4px;
    }

    .logo {
      display: block;
      margin: 10px auto 0;
      max-width: 180px;
      height: auto;
    }

    .back-link {
      position: absolute;
      left: 16px;
      top: 50%;
      transform: translateY(-50%);
      font-family: 'Courier New', Courier, monospace;
      font-size: 0.8em;
      font-weight: 700;
      color: #000;
      text-decoration: none;
      border-bottom: 1px solid #999;
      letter-spacing: 0.05em;
    }

    .back-link:hover {
      border-bottom-color: #000;
    }

    /* Menu page */
    .menu {
      border-top: 2px solid #fff;
    }

    .menu a {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 18px 24px;
      border-bottom: 1px solid #333;
      color: #fff;
      text-decoration: none;
      font-size: 1.1em;
      letter-spacing: 0.08em;
      transition: background 0.15s;
    }

    .menu a:last-child {
      border-bottom: none;
    }

    .menu a:hover {
      background: #111;
    }

    .menu .count {
      font-size: 0.75em;
      color: #666;
      font-weight: 400;
    }

    /* Section headers — white band (for sub-sections) */
    .table-section {
      border-bottom: 2px solid #fff;
    }

    .table-section:last-child {
      border-bottom: none;
    }

    .table-header {
      background: #fff;
      padding: 10px 12px;
      text-align: center;
    }

    .table-section h2 {
      font-family: Georgia, 'Times New Roman', serif;
      font-size: 1em;
      font-weight: 700;
      letter-spacing: 0.05em;
      text-transform: uppercase;
      color: #000;
      margin: 0;
      padding: 0;
      border: none;
      display: inline;
    }

    .table-section h2 .count {
      font-weight: 400;
      font-style: italic;
      text-transform: none;
      letter-spacing: 0;
      font-size: 0.9em;
      margin-left: 6px;
    }

    /* Tables */
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 0.8em;
    }

    thead th {
      font-family: Georgia, 'Times New Roman', serif;
      text-align: left;
      padding: 4px 5px;
      font-weight: 900;
      font-size: 0.8em;
      letter-spacing: 0.05em;
      color: #fff;
      border-bottom: 1px solid #fff;
      background: #000;
      position: sticky;
      top: 0;
      z-index: 10;
    }

    tbody tr {
      border-bottom: 1px solid #333;
    }

    tbody tr:last-child {
      border-bottom: none;
    }

    td {
      padding: 3px 5px;
      vertical-align: top;
    }

    /* Column type styles */
    .col-rank {
      width: 28px;
      text-align: center;
      font-weight: 700;
    }

    .col-text {
      max-width: 140px;
    }

    .col-text-sm {
      max-width: 90px;
    }

    .col-text-dim {
      width: 1px;
      max-width: 120px;
      color: #999;
    }



    .col-num {
      width: 1px;
      text-align: center;
      font-variant-numeric: tabular-nums;
      white-space: nowrap;
    }

    .col-grade {
      width: 1px;
      text-align: center;
      font-weight: 900;
      white-space: nowrap;
    }

    .clamp {
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    thead th.col-rank,
    thead th.col-num,
    thead th.col-grade {
      text-align: center;
    }

    .sub-header td {
      background: #fff;
      padding: 10px 12px;
      text-align: center;
      border-top: 2px solid #fff;
    }

    .sub-header h2 {
      font-family: Georgia, 'Times New Roman', serif;
      font-size: 1em;
      font-weight: 700;
      letter-spacing: 0.05em;
      text-transform: uppercase;
      color: #000;
      margin: 0;
      display: inline;
    }

    .sub-header h2 .count {
      font-weight: 400;
      font-style: italic;
      text-transform: none;
      letter-spacing: 0;
      font-size: 0.9em;
      margin-left: 6px;
    }


    /* Links */
    a {
      color: #fff;
      text-decoration: none;
      border-bottom: 1px solid #555;
    }

    a:hover {
      border-bottom-color: #fff;
    }

    .table-header a,
    header a {
      color: #000;
      border-bottom: none;
    }

    /* Gallery grid (Lego page) */
    .gallery {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(160px, 1fr));
      gap: 4px;
      padding: 4px;
      background: #000;
    }

    .gallery img {
      width: 100%;
      height: auto;
      display: block;
      border: 1px solid #333;
    }

    @media (max-width: 600px) {
      .gallery {
        grid-template-columns: repeat(auto-fill, minmax(120px, 1fr));
        gap: 3px;
        padding: 3px;
      }
    }

    /* Footer */
    footer {
      text-align: center;
      padding: 14px 0;
      border-top: 2px solid #fff;
      font-size: 0.75em;
      color: #666;
      background: #000;
      letter-spacing: 0.05em;
    }

    /* Responsive */
    @media (max-width: 600px) {
      body { padding: 8px; }

      header h1 { font-size: 2em; letter-spacing: 0.04em; }

      .back-link { font-size: 0.7em; left: 10px; }

      .menu a { padding: 14px 16px; font-size: 0.95em; }

      table { font-size: 0.65em; }
      td, thead th { padding: 2px 3px; }

      .col-text { max-width: 80px; }
      .col-text-sm { max-width: 60px; }
      .col-text-dim { max-width: 70px; }
      .col-rank { width: 20px; }
    }

    @media (max-width: 380px) {
      header h1 { font-size: 1.7em; }

      .menu a { padding: 12px 12px; font-size: 0.85em; }

      table { font-size: 0.58em; }
      td, thead th { padding: 2px 2px; }

      .col-text { max-width: 65px; }
      .col-text-sm { max-width: 50px; }
      .col-text-dim { max-width: 55px; }
      .col-rank { width: 16px; }
    }
"""


# ---------------------------------------------------------------------------
# HTML generation
# ---------------------------------------------------------------------------

def page_shell(title, body_html, page_title="The CountUp"):
    """Wrap body content in a full HTML page."""
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{escape(page_title)}</title>
  <style>{CSS}
  </style>
</head>
<body>
{body_html}
</body>
</html>"""


def build_table(columns, rows):
    """Build an HTML table from column defs and row data."""
    # thead
    ths = []
    for col_name, col_type in columns:
        cls = f"col-{col_type}"
        ths.append(f'<th class="{cls}">{escape(col_name)}</th>')
    thead = f'<thead><tr>{"".join(ths)}</tr></thead>'

    # tbody
    trs = []
    for row_vals in rows:
        tds = []
        for val, (_, col_type) in zip(row_vals, columns):
            cls = f"col-{col_type}"
            cell = escape(str(val)) if val else "\u2014"
            if col_type == "text" or col_type == "text-sm" or col_type == "text-dim":
                cell = f'<div class="clamp">{cell}</div>'
            tds.append(f'<td class="{cls}">{cell}</td>')
        trs.append(f'<tr>{"".join(tds)}</tr>')
    tbody = f'<tbody>{"".join(trs)}</tbody>'

    return f'<table>{thead}{tbody}</table>'


def build_index(categories, lego_count=0):
    """Build the menu / index page."""
    menu_items = []
    for cat in categories:
        if cat["sub_sections"]:
            count = sum(len(s["items"]) for s in cat["sub_sections"])
        else:
            count = len(cat["items"])
        menu_items.append(
            f'<a href="{cat["id"]}.html">'
            f'<span>{escape(cat["name"])}</span>'
            f'<span class="count">{count}</span>'
            f'</a>'
        )
        # Insert Lego before Games
        if cat["id"] == "games" and lego_count:
            menu_items.insert(-1,
                f'<a href="lego.html"><span>Lego</span>'
                f'<span class="count">{lego_count}</span></a>'
            )

    # Look up Top Spin Collection album count
    tsc_json = os.path.join(os.path.dirname(SCRIPT_DIR), "top-spin-collection", "top400_data.json")
    try:
        with open(tsc_json, encoding="utf-8") as f:
            tsc_count = json.load(f)["total_unique"]
    except (FileNotFoundError, KeyError):
        tsc_count = "\u2014"

    body = f"""  <div class="container">
    <header>
      <h1>The CountUp</h1>
      <img src="logo.jpg" alt="The CountUp" class="logo">
    </header>

    <nav class="menu">
      <a href="https://joeldechant.github.io/nba-tap-rankings/" target="_blank"><span>NBA TED / TAP</span><span class="count">100</span></a>
      <a href="https://joeldechant.github.io/top-spin-collection/" target="_blank"><span>Top Spin</span><span class="count">{tsc_count}</span></a>
      {"".join(menu_items)}
    </nav>

    <footer>The CountUp Rankings</footer>
  </div>"""

    return page_shell("TheCountUp", body)


def items_to_rows(cat):
    """Convert a category's items list into table row tuples."""
    rows = []
    items = cat["items"]
    cols = cat["columns"]
    col_keys = [c[0] for c in cols]

    for idx, item in enumerate(items):
        row = []
        for col_name, col_type in cols:
            if col_type == "rank":
                # Use explicit rank if present (Dining), else position
                if "rank" in item and item["rank"] is not None:
                    row.append(str(item["rank"]))
                else:
                    row.append(str(idx + 1))
            elif col_name == "BRAND":
                row.append(item.get("brand", ""))
            elif col_name == "FLAVOR":
                row.append(item.get("flavor", ""))
            elif col_name in ("SUGAR", "SUGAR(g)"):
                row.append(item.get("sugar", ""))
            elif col_name == "CLASSICS":
                row.append(item.get("classics", ""))
            elif col_name == "COUNTRY":
                row.append(item.get("country", ""))
            elif col_name == "VARIETY":
                row.append(item.get("variety", ""))
            elif col_name == "TYPE":
                row.append(item.get("type", ""))
            elif col_name == "NAME":
                row.append(item.get("name", ""))
            elif col_name == "CACAO":
                row.append(item.get("cacao", ""))
            elif col_name == "ORIGIN":
                row.append(item.get("origin", ""))
            elif col_name == "RICE":
                row.append(item.get("rice", ""))
            elif col_name == "CHICKEN":
                row.append(item.get("chicken", ""))
            elif col_name == "RESTAURANT":
                row.append(item.get("restaurant", ""))
            elif col_name == "ITEM":
                row.append(item.get("item", ""))
            elif col_name == "RATING":
                row.append(item.get("rating", ""))
            elif col_name == "GAME":
                row.append(item.get("game", ""))
            elif col_name == "YEAR":
                row.append(str(item["year"]) if item.get("year") else "")
            elif col_name == "SCORE":
                row.append(item.get("score", ""))
            elif col_name == "RANK":
                row.append(str(item["rank"]) if item.get("rank") is not None else str(idx + 1))
            else:
                row.append("")
        rows.append(row)
    return rows


def build_category_page(cat):
    """Build a single category page."""
    cat_name = cat["name"]

    if cat["sub_sections"]:
        # Sub-sectioned layout (Sauces) — single table with tbody groups
        ncols = len(cat["columns"])
        # thead
        ths = []
        for col_name, col_type in cat["columns"]:
            cls = f"col-{col_type}"
            ths.append(f'<th class="{cls}">{escape(col_name)}</th>')
        thead = f'<thead><tr>{"".join(ths)}</tr></thead>'

        tbodies = []
        for sec in cat["sub_sections"]:
            header_row = (
                f'<tr class="sub-header"><td colspan="{ncols}">'
                f'<h2>{escape(sec["name"])} '
                f'<span class="count">({len(sec["items"])})</span></h2>'
                f'</td></tr>'
            )
            trs = [header_row]
            for idx, item in enumerate(sec["items"]):
                row = []
                for col_name, col_type in cat["columns"]:
                    if col_type == "rank":
                        row.append(str(idx + 1))
                    elif col_name == "ORIGIN":
                        row.append(item.get("origin", ""))
                    elif col_name == "NAME":
                        row.append(item.get("name", ""))
                    elif col_name == "BRAND":
                        row.append(item.get("brand", ""))
                    elif col_name == "RICE":
                        row.append(item.get("rice", ""))
                    elif col_name == "CHICKEN":
                        row.append(item.get("chicken", ""))
                    else:
                        row.append("")
                tds = []
                for val, (_, col_type) in zip(row, cat["columns"]):
                    cls = f"col-{col_type}"
                    cell = escape(str(val)) if val else "\u2014"
                    if col_type in ("text", "text-sm", "text-dim"):
                        cell = f'<div class="clamp">{cell}</div>'
                    tds.append(f'<td class="{cls}">{cell}</td>')
                trs.append(f'<tr>{"".join(tds)}</tr>')
            tbodies.append(f'<tbody>{"".join(trs)}</tbody>')

        content = f'<table>{thead}{"".join(tbodies)}</table>'
        total = sum(len(s["items"]) for s in cat["sub_sections"])
    else:
        # Single table layout
        rows = items_to_rows(cat)
        content = build_table(cat["columns"], rows)
        total = len(cat["items"])

    body = f"""  <div class="container">
    <header>
      <a href="index.html" class="back-link">&larr; Menu</a>
      <h1>{escape(cat_name)}</h1>
      <div class="subtitle">{total} ranked</div>
    </header>

    <main>
      {content}
    </main>

    <footer>The CountUp Rankings</footer>
  </div>"""

    return page_shell(f"{cat_name} — The CountUp", body, f"{cat_name} — The CountUp")


# ---------------------------------------------------------------------------
# Lego gallery
# ---------------------------------------------------------------------------

def process_lego_images():
    """Resize pre-cropped Lego photos from 'Lego Pics/' folder. Returns list of output filenames."""
    src_dir = os.path.join(SCRIPT_DIR, "Lego Pics")
    EXCLUDE = {"IMG_8052.jpeg"}
    src_files = sorted(glob.glob(os.path.join(src_dir, "IMG_*.jpeg")))
    src_files = [f for f in src_files if os.path.basename(f) not in EXCLUDE]
    if not src_files:
        return []

    random.shuffle(src_files)

    out_dir = os.path.join(DOCS_DIR, "lego")
    os.makedirs(out_dir, exist_ok=True)
    filenames = []

    for i, filepath in enumerate(src_files, 1):
        img = Image.open(filepath)
        img = ImageOps.exif_transpose(img)
        img.thumbnail((600, 800), Image.LANCZOS)

        fname = f"lego_{i:02d}.jpg"
        img.save(os.path.join(out_dir, fname), "JPEG", quality=85)
        filenames.append(fname)

    return filenames


def build_lego_page(filenames):
    """Build the Lego gallery page."""
    imgs = "\n".join(f'      <img src="lego/{f}" alt="Lego character">' for f in filenames)

    body = f"""  <div class="container">
    <header>
      <a href="index.html" class="back-link">&larr; Menu</a>
      <h1>Lego</h1>
      <div class="subtitle">{len(filenames)} custom characters</div>
    </header>

    <main>
      <div class="gallery">
{imgs}
      </div>
    </main>

    <footer>The CountUp Rankings</footer>
  </div>"""

    return page_shell("Lego — The CountUp", body, "Lego — The CountUp")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("Connecting to open workbook...")
    wb = xw.Book("TheCountUp Rankings.xlsx")

    print("Extracting data from sheets...")
    categories = [
        extract_games(wb.sheets["Games"]),
        extract_dining(wb.sheets["Dine"]),
        extract_pop(wb.sheets["Pop"]),
        extract_candy(wb.sheets["Candy"]),
        extract_chocolate(wb.sheets["Choc"]),
        extract_sauces(wb.sheets["Sauces"]),
    ]

    for cat in categories:
        if cat["sub_sections"]:
            n = sum(len(s["items"]) for s in cat["sub_sections"])
        else:
            n = len(cat["items"])
        print(f"  {cat['name']}: {n} items")

    os.makedirs(DOCS_DIR, exist_ok=True)

    # Process Lego photos
    print("Processing Lego photos...")
    lego_files = process_lego_images()
    if lego_files:
        print(f"  Lego: {len(lego_files)} characters")
        html = build_lego_page(lego_files)
        path = os.path.join(DOCS_DIR, "lego.html")
        with open(path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"Written: {path}")
    else:
        print("  No Lego photos found")

    # Build index page
    html = build_index(categories, lego_count=len(lego_files))
    path = os.path.join(DOCS_DIR, "index.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Written: {path}")

    # Build category pages
    for cat in categories:
        html = build_category_page(cat)
        path = os.path.join(DOCS_DIR, f"{cat['id']}.html")
        with open(path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"Written: {path}")

    print(f"\nWebsite built in {DOCS_DIR}/")
    print("Done.")


if __name__ == "__main__":
    main()
