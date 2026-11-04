---
"turbosheet": minor
---

Chart writing: `addChart(ws, chart)` / `ws.charts` writes native bar, column, line, pie, doughnut,
scatter and area charts (titles, legend, axis titles / bounds / number formats, stacking, data
labels, series colors, markers, smoothing, built-in styles) with caches of the referenced cells,
sharing the sheet's drawing with pictures and with charts kept from read files. New opt-in
`chartsWriter` for `turbosheet/lite`. Reading keeps charts raw (round-trip preservation).
