# The CountUp — Project Instructions

## What This Is
A static ranking website generated from `TheCountUp Rankings.xlsx`. The build script (`build_website.py`) reads the Excel file via xlwings and outputs HTML files to `docs/` for GitHub Pages hosting.

## Build & Preview
- **Build:** `python build_website.py` (Excel must be open)
- **Preview:** `python -m http.server 8080 --directory docs` (launch.json configured)
- **Output:** `docs/` folder (index.html + 6 category pages + lego.html + logo.jpg + lego/ images)
- **Dependencies:** xlwings, Pillow, numpy
- **Logo:** `Thumbs Up Image.jpg` is copied to `docs/logo.jpg` — must re-copy if image changes

## Important Conventions
- The Excel file is typically open in Excel — connect with `xw.Book("TheCountUp Rankings.xlsx")`, never `xw.App(visible=False)`
- Column headers are hardcoded as ALL CAPS in Python (no CSS text-transform on thead)
- Exception: lowercase "g" in `SUGAR(g)`
- Text cells use single-line `nowrap` with ellipsis — never allow 2-line wrapping
- Numeric/grade/dim columns use `width: 1px` to shrink-to-fit — text columns get the remaining space
- Sauces sub-sections render as one `<table>` with multiple `<tbody>` groups and `.sub-header` rows (not separate tables)
- Title is "The CountUp" with mixed case — do not apply text-transform: uppercase to h1
- Homepage menu order: NBA TAP (external), Top Spin (external), Lego, Games, Dining, Pop, Candy, Chocolate, Sauces
- External links open in new tab; NBA TAP count is hardcoded (100), Top Spin count is dynamic (read from `top400_data.json`)
- External link labels are concise: "NBA TAP" and "Top Spin" (no extra suffixes)

## Lego Gallery
- Photo gallery (not Excel-driven) — source photos are `IMG_*.jpeg` in project root
- Build auto-detects figure center, crops tight, resizes to 600x800, saves to `docs/lego/`
- To add more characters: drop `IMG_*.jpeg` files in project root and rebuild

## Deploy
Target: GitHub Pages serving from `docs/` folder on `main` branch.
