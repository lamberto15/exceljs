---
"turbosheet": minor
---

Excel-like polish: `new Workbook({ defaultFont })` sets the workbook's Normal font; text with line breaks gets `wrapText` automatically and unstyled hyperlink cells get Excel's hyperlink look (both as Excel does when typing); `ws.ignoredErrors` hides error-check triangles such as "number stored as text" (new `ignoredErrorsWriter` / `ignoredErrorsReader` feature); manual page breaks via `ws.rowBreaks` / `ws.colBreaks` (print setup feature; they and ignored-error ranges follow row/column edits); `setSheetStyle(ws, style)` formats the whole sheet, e.g. to unlock all cells before protecting it.
