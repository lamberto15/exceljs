---
"turbosheet": minor
---

Formula engine, group 15 (#104, #105, #106):

- **Dynamic arrays.** Dynamic-array formulas are evaluated in array mode and spill. `evaluate` writes the spilled values and the spill range (`Formula.array`). A blocked area gives `#SPILL!`. `A1#` spill references work (stored in files as `ANCHORARRAY(A1)`), and formulas reading spilled cells see the spilled values. New functions: `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `TRANSPOSE`, `RANDARRAY`. New error codes: `#SPILL!` and `#CALC!`; they are written to files as `#VALUE!`, which is Excel's plain value for them. `TRANSPOSE` formulas are now dynamic by default, like the other spilling functions.
- **Structured references and LET.** `Sales[Amount]`, `[@Qty]`, `Sales[[#Totals],[A]:[B]]`, `#All` / `#Headers` / `#Data` / `#This Row` are resolved against `ws.tables`, dependencies included. `LET` is supported, and writers add the `_xlpm.` prefix to LET names (Excel removes the formula without it). `SUBTOTAL` is supported (1-11 / 101-111: it skips nested subtotals and, for 101-111, hidden rows).
- **Incremental recalculation.** `track(wb)` evaluates once and keeps the dependency graph. `calc.set(sheet, address, value)` / `calc.update(cells)` then re-evaluate only the formulas downstream of the change, about 5 ms for a small change on a 100k-formula sheet versus 56 ms for a full evaluation. Volatile functions are always recalculated. Formula edits, cycles and dynamic arrays fall back to a full recalculation.
