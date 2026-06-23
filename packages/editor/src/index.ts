/**
 * @docen/editor — assembly layer that wires the Fluent UI layer to @docen/docx
 * into turnkey editor super-components. The UI layer lives in ./ui (formerly
 * the standalone @docen/ui package, now inlined to keep i18n side effects in
 * the same bundle and remove the cross-package tree-shaking hazard).
 *
 * @module
 */

// Super-components (register their custom elements on import)
export { default as DocenDocument } from "./document";
export { default as DocenPresentation } from "./presentation";
export { default as DocenWorkbook } from "./workbook";

// Re-export the UI bootstrap so demos/consumers import everything from one
// entry — matches importing ./src/index.ts directly.
export { applyTheme, registerComponents } from "./ui";
