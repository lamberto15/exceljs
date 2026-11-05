---
"turbosheet": patch
---

`TEXT()` (and `formatValue` / `cellText`) render fraction formats: `# ?/?` (whole part + fraction), `?/?` (improper), denominators of up to three digits (`# ??/??`, `???/???`) or a fixed one (`# ?/8`), with Excel's padding and its choice of fraction (the last continued-fraction convergent that fits: 0.14159 as `??/??` is `1/7`). Before, `TEXT(0.5,"# ?/?")` returned `" 0/1"`. 110 cases were checked against Excel for Mac. Formulas written with a partial future-function prefix (`_xlfn.SORT(`, `_xlws.FILTER(`) now get the full `_xlfn._xlws.` prefix and are recognised as dynamic-array formulas.
