---
"turbosheet": patch
---

Round-trip fidelity: `readXlsx` → `writeXlsx` now keeps autoFilter criteria (filter columns, sort state), named cell styles (`cellStyleXfs` / `cellStyles` and each cell format's style link, exposed read-only as `Xf.xfId`), `sheetFormatPr` attributes (default column width, base width, zero height, ...), form controls on sheets that also have notes (VML shapes merged with the regenerated notes; Excel's `mc:AlternateContent`-wrapped `<controls>` and their ctrlProps parts), the tab position of chartsheets / dialogsheets (activeTab and sheet-scoped names follow), and the queryTable link of unchanged query tables. See "Round-trip preservation" in docs/features.md for what happens when you edit nearby.
