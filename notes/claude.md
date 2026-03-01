# Claude Session Notes

## 2026-03-01 — Fix Sauces column alignment across sub-categories

**Problem:** On the Sauces page, the RICE and CHICKEN grade columns were slightly offset to the right in the middle sub-category ("Other Main Sauces") compared to "Asian Sauces" and "Finishing Sauces".

**Root cause:** Each sub-category renders as a separate `<table>` element. Without `table-layout: fixed`, the browser auto-sizes columns based on content. The middle section had longer text (e.g., "South Africa", "Taste of Inspirations") that pushed columns rightward.

**Fix:**
1. Added `table-layout: fixed` to `.table-section table` — forces fixed column widths based on header definitions, not content
2. Added explicit `width` values alongside existing `max-width` on `.col-text`, `.col-text-sm`, `.col-text-dim` — `table-layout: fixed` ignores `max-width`, only respects `width`
3. Updated responsive breakpoints (600px, 380px) with matching `width` values

**Result:** All three sub-tables now have identical column positions (verified via getBoundingClientRect).

**Scope:** Scoped `table-layout: fixed` to `.table-section table` only (not global `table`) to avoid affecting single-table category pages.
