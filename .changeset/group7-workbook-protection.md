---
"turbosheet": minor
---

Document properties: `wb.title`, `subject`, `keywords`, `description`, `category`, `contentStatus`, `company`, `manager`, written to docProps and read back. `protectSheet(ws, { password, editable, allow })` unlocks editable ranges and protects the sheet in one call. Pixel conversions for pictures and the default row height follow the workbook default font (`fontMetrics(font)`).
