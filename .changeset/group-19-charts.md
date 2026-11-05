---
"turbosheet": minor
---

Charts, group 19 (#113, #114):

- **Reading charts.** Charts of the supported types become `ws.charts` when a file is read: 2-D bar / column / line / pie / doughnut / scatter / area, combos, secondary axes, trendlines and data labels. A read chart you don't change is written back byte for byte, Excel's formatting included; an edited chart is regenerated from the model. 3-D, bubble, stock, radar and surface charts, chart sheets and charts with literal data stay raw. New feature half: `chartsReader`. Framework: `SheetReadCtx.claim(path, model)` lets a reader own a part, so the raw reference to it is dropped and the original part is kept while the model is unchanged.
- **Combo charts and more** in `addChart`, per series:
  - `type: 'line' | 'column' | 'area'` mixes series types;
  - `axis: 'secondary'` adds a right-hand value axis (`axes.y2`);
  - `trendline` supports linear / exponential / logarithmic / polynomial / power / moving average, with the equation and R²;
  - `dataLabels` overrides the chart's labels (`false`, `position`, `numFmt`).
  - `stacked` now applies only to series of the chart's own type.
