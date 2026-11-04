---
"turbosheet": minor
---

Add formula evaluation as a separate entry point, `turbosheet/formula`: `parseFormula`, `evaluate(wb)` (dependency-ordered, circular-reference detection, optional iterative calculation, results stored into `Formula.result`), `evaluateCell`, `evaluateFormula` and `createEvaluator({ functions })` for per-instance custom functions. 87 built-in functions (math and conditional aggregation, logical, lookup incl. XLOOKUP/XMATCH, text incl. TEXT formats, dates, info) with Excel's coercion, comparison and wildcard rules. The core bundles do not change.
