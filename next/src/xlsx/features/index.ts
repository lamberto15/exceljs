// Public API of the features (re-exported from the package root): model types and helpers.
// Each feature adds its public exports here. The registry (registry.ts) and the extension-point
// types (./types.ts) are internal.

export type { HeaderFooter, PageMargins, PageSetup } from './print-setup.js';
export { hashPassword, type SheetProtection } from './sheet-protection.js';
