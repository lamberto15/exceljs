// Public API of the features (re-exported from the package root and from `./lite`): model
// types, helpers and each feature's opt-in halves (`<name>Writer` / `<name>Reader`, for
// `options.features` in `./lite`; the root always runs them). The registry (registry.ts) is
// root-only and internal; of ./types.ts only the two extension-point interfaces are public.

export {
  addChart,
  type Chart,
  type ChartAxis,
  type ChartDataLabels,
  type ChartSeries,
  type ChartTrendline,
  type ChartType,
  chartsReader,
  chartsWriter,
} from './charts.js';
export {
  addNote,
  addThreadedComment,
  commentsReader,
  commentsWriter,
  type Note,
  type Person,
  type ThreadedComment,
  type ThreadedReply,
} from './comments.js';
export {
  addConditionalFormatting,
  type CfAboveAverageRule,
  type CfBlankErrorRule,
  type CfCellIsRule,
  type CfColorScaleRule,
  type CfDataBarRule,
  type CfExpressionRule,
  type CfIconSetName,
  type CfIconSetRule,
  type CfOperator,
  type CfRule,
  type CfTextRule,
  type CfTimePeriod,
  type CfTimePeriodRule,
  type CfTop10Rule,
  type CfValuesRule,
  type Cfvo,
  type CfvoType,
  type ConditionalFormatting,
  conditionalFormattingReader,
  conditionalFormattingWriter,
} from './conditional-formatting.js';
export {
  addCustomView,
  type CustomSheetView,
  type CustomView,
  customViewsReader,
  customViewsWriter,
} from './custom-views.js';
export {
  addDataValidation,
  type DataValidation,
  type DataValidationFormula,
  type DataValidationOperator,
  type DataValidationType,
  dataValidationReader,
  dataValidationWriter,
} from './data-validation.js';
export {
  addDefinedName,
  type DefinedNameOptions,
  getDefinedName,
  getPrintArea,
  getPrintTitles,
  type PrintTitles,
  removeDefinedName,
  setPrintArea,
  setPrintTitles,
} from './defined-names.js';
export { type AnchorInput, drawingEmitter } from './drawing.js';
export {
  type IgnoredError,
  ignoredErrorsReader,
  ignoredErrorsWriter,
} from './ignored-errors.js';
export {
  addImage,
  type CellPosition,
  type FitImageOptions,
  fitImage,
  type ImageAnchor,
  type ImageExtension,
  type ImagePlacement,
  type ImageSize,
  imageSize,
  imagesReader,
  imagesWriter,
  type PlaceImageOptions,
  placeImage,
  type WorkbookImage,
} from './images.js';
export {
  addPivotTable,
  type PivotFunction,
  type PivotTable,
  type PivotValue,
  pivotTablesWriter,
} from './pivot-tables.js';
export {
  type HeaderFooter,
  type PageMargins,
  type PageSetup,
  printSetupReader,
  printSetupWriter,
} from './print-setup.js';
export {
  hashPassword,
  type SheetProtection,
  sheetProtectionReader,
  sheetProtectionWriter,
} from './sheet-protection.js';
export {
  addTableColumn,
  removeTableColumns,
  removeTableRows,
  renameTableColumn,
  setTableHeaderRow,
  setTableTotalsRow,
} from './table-edit.js';
export {
  type AddTableColumn,
  type AddTableOptions,
  addTable,
  addTableRows,
  getTable,
  type Table,
  type TableColumn,
  type TableStyle,
  type TotalsRowFunction,
  tableRows,
  tablesReader,
  tablesWriter,
} from './tables.js';
export { themeWriter } from './theme.js';
export type { FeatureReader, FeatureWriter } from './types.js';
