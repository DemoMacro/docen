/**
 * @docen/editor — assembly layer that bundles the Fluent UI shell with
 * @docen/docx into turnkey editor elements. The UI layer lives in ./ui
 * (inlined so i18n side effects stay in one bundle, with no cross-package
 * tree-shaking hazard).
 *
 * @module
 */

// Editor elements (register their custom elements on import)
export { default as DocenDocument } from "./document";
export { default as DocenPresentation } from "./presentation";
export { default as DocenWorkbook } from "./workbook";

// Re-export the UI bootstrap so demos/consumers import everything from one
// entry — matches importing ./src/index.ts directly.
export {
  applyTheme,
  availableLanguages,
  notifyLocaleChange,
  observeLang,
  registerComponents,
  registerLocalization,
  registerTranslation,
  t,
} from "./ui";

// Public types for add-in authors and host consumers.
export type { TaskPaneId, VisibilityMode } from "./document";
export type { AddinHost } from "./ui/addin/host";
export type { DocenHost, DocenAddin, RibbonTab } from "./ui/addin/types";
export type { AdditionalLanguage, LanguageOption, LocalizationInfo } from "./ui";
