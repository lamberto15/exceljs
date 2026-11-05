---
"turbosheet": minor
---

`createXlsxWriter` now defaults to `profile: 'low-memory'` (inline strings). With shared strings the table held every distinct string until `end()` and was serialized there in one blocking pass (1M x 35 rows: 527 MB peak RSS and a 304 ms stall, vs 207 MB / 42 ms inline, and 16% faster). Pass `profile: 'balanced'` or `strings: 'shared'` to keep sharedStrings.xml. `writeXlsx` and `xlsxStream` still default to `balanced`. Closes #72.
