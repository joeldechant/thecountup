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
- Column headers are sticky (`position: sticky; top: 0`) on all category pages (not Lego gallery)
- Sauces sub-sections render as one `<table>` with multiple `<tbody>` groups and `.sub-header` rows (not separate tables)
- Title is "The CountUp" with mixed case — do not apply text-transform: uppercase to h1
- Homepage menu order: NBA TED / TAP (external), Top Spin (external), Lego, Games, Dining, Pop, Candy, Chocolate, Sauces
- External links open in new tab; NBA TED / TAP count is hardcoded (100), Top Spin count is dynamic (read from `top400_data.json`)
- External link labels are concise: "NBA TED / TAP" and "Top Spin" (no extra suffixes)

## Lego Gallery
- Photo gallery (not Excel-driven) — source photos are pre-cropped `.jpeg`/`.jpg` files in `Lego Pics/` folder (user crops manually)
- Build uses `thumbnail((600, 800))` to limit size while preserving aspect ratio — no forced resize
- User prefers natural aspect ratios; slight gallery gaps are OK — do not force uniform dimensions
- Randomly shuffles order, saves to `docs/lego/`
- `EXCLUDE` set in `process_lego_images()` skips specific files (currently `IMG_8052.jpeg`)
- To add more characters: drop `.jpeg`/`.jpg` files in `Lego Pics/` and rebuild

## Deploy
Target: GitHub Pages serving from `docs/` folder on `main` branch.
