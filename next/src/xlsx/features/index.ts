// Public API of the features (re-exported from the package root): model types and helpers.
// Each feature adds its public exports here. The registry (registry.ts) and the extension-point
// types (./types.ts) are internal.

export { addNote, type Note } from './comments.js';
export {
  addDataValidation,
  type DataValidation,
  type DataValidationFormula,
  type DataValidationOperator,
  type DataValidationType,
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
export {
  addImage,
  type CellPosition,
  type ImageAnchor,
  type ImageExtension,
  type ImagePlacement,
  type ImageSize,
  type PlaceImageOptions,
  placeImage,
  type WorkbookImage,
} from './images.js';
export type { HeaderFooter, PageMargins, PageSetup } from './print-setup.js';
export { hashPassword, type SheetProtection } from './sheet-protection.js';
export {
  type AddTableOptions,
  addTable,
  type Table,
  type TableColumn,
  type TableStyle,
  type TotalsRowFunction,
} from './tables.js';
