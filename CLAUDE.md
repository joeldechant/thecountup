# The CountUp — Project Instructions

## What This Is
A static ranking website generated from `TheCountUp Rankings.xlsx`. The build script (`build_website.py`) reads the Excel file via xlwings and outputs HTML files to `docs/` for GitHub Pages hosting.

## Build & Preview
- **Build website:** `python build_website.py` (Excel must be open)
- **Build analytics:** `python analyze_rankings.py` (Excel must be open) — generates `CountUp Analysis.xlsx`, `analytics/analytics.html`, `analytics/analytics_chart.png`
- **Preview:** `python -m http.server 8080 --directory docs` (launch.json configured)
- **Output:** `docs/` folder (index.html + 6 category pages + lego.html + logo.jpg + lego/ images)
- **Dependencies:** xlwings, Pillow, numpy, matplotlib, openpyxl
- **Logo:** `Thumbs Up Image.jpg` is copied to `docs/logo.jpg` — must re-copy if image changes

## Important Conventions
- The Excel file is typically open in Excel — connect with `xw.Book("TheCountUp Rankings.xlsx")`, never `xw.App(visible=False)`
- Column headers are hardcoded as ALL CAPS in Python (no CSS text-transform on thead)
- Exceptions: lowercase "g" in `SUGAR(g)`, lowercase in `$/4oz`, and `SUGAR<br>(g/oz)` uses a `<br>` tag for 2-line header
- Headers containing `<br>` skip HTML escaping in `build_table()`
- Text cells use single-line `nowrap` with ellipsis — never allow 2-line wrapping
- Numeric/grade/dim columns use `width: 1px` to shrink-to-fit — text columns get the remaining space
- `col-num` has extra right padding (24px) on desktop only (via `@media (min-width: 601px)`); base padding is 12px
- `col-grade` has 12px padding on both sides
- Column headers are sticky (`position: sticky; top: 0`) on all category pages (not Lego gallery)
- Sauces sub-sections render as one `<table>` with multiple `<tbody>` groups and `.sub-header` rows (not separate tables)
- Sauces has mobile-only CSS overrides (via `extra_css` in `page_shell`) — ORIGIN slightly wider, BRAND slightly narrower than their default column-type widths
- `page_shell()` accepts an optional `extra_css` parameter for page-specific styles (used by Sauces)
- Title is "The CountUp" with mixed case — do not apply text-transform: uppercase to h1
- Homepage menu order: NBA TED / TAP (external), Top Spin (external), Lego, Games, Fast Food Hack, Pop, Candy, Chocolate, Sauces
- "Dining" category renamed to "Fast Food Hack" (id still "dining", file still `dining.html`)
- External links open in new tab; NBA TED / TAP count is hardcoded (100), Top Spin count is dynamic (read from `top400_data.json`)
- Top Spin count is rounded up to even if odd (so homepage always shows an even number)
- External link labels are concise: "NBA TED / TAP" and "Top Spin" (no extra suffixes)
- Mobile back-link repositioned to subtitle level (`bottom: 14px`) to avoid overlap with long titles like "Fast Food Hack"

## Lego Gallery
- Photo gallery (not Excel-driven) — source photos are pre-cropped `.jpeg`/`.jpg` files in `Lego Pics/` folder (user crops manually)
- Build uses `ImageOps.fit((600, 900))` for 2:3 cover-fit — uniform tiles, no gaps
- Randomly shuffles order, saves to `docs/lego/`
- `EXCLUDE` set in `process_lego_images()` skips specific files (currently `IMG_8052.jpeg`, `IMG_8126.jpeg`)
- To add more characters: drop `.jpeg`/`.jpg` files in `Lego Pics/` and rebuild

## Updating Data from External Sheets
- User may drop an `Update Sheets.xlsx` file into the project folder containing updated versions of one or more worksheets (e.g. Candy, Dine, Sauces)
- The update file can contain multiple sheets — always enumerate all sheets before processing
- To apply: read each sheet's data from the update file (using `xw.App(visible=False)`), clear the corresponding sheet in the main workbook (connected via `xw.Book()`), write the new data, save, then rebuild
- After updating, run `python build_website.py` to regenerate the site

## Analytics
- `analyze_rankings.py` is a separate script from `build_website.py` — they are not linked
- Reads all 6 category sheets and produces three outputs:
  - `CountUp Analysis.xlsx` — multi-sheet Excel workbook (Overview, Geography, Price Analysis, Grades, Pop Sugar, Games by Decade)
  - `analytics/analytics.html` — standalone HTML page styled to match the site theme
  - `analytics/analytics_chart.png` — horizontal bar chart of top 15 countries/origins (Candy + Sauces)
- Analytics outputs live in `analytics/` (not `docs/`) so they never go live on the website
- Must be rerun manually after data changes; does not auto-update

## Deploy
Target: GitHub Pages serving from `docs/` folder on `main` branch.
