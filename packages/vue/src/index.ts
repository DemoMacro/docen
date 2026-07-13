// @docen/vue — Vue 3 adapter for the <docen-document> web component.
//
// The editor UI is a framework-neutral fast-element web component in
// @docen/editor; this package wraps it as a typed Vue component — v-model for
// content, a `v-slot="{ editor }"` scope, and a template-ref expose — so a Vue
// app gets props/emits and can reach the underlying Tiptap editor. Only
// @docen/editor is a runtime dependency — @docen/docx is reached through it
// (chain), and its converters are framework-neutral, so import them directly
// from "@docen/docx" rather than through this adapter.
export { DocenDocument } from "./DocenDocument";
// Cancelable host-action event types — consumers type @save/@open/@print
// handlers with these and call e.preventDefault() to take over the host's
// built-in save/open/print dialog. See DocenDocument emits.
export type {
  DocenSaveEvent,
  DocenSaveAsEvent,
  DocenOpenEvent,
  DocenNewEvent,
  DocenPrintEvent,
} from "./DocenDocument";

// Re-export the web-component bootstrap so a Vue app imports everything from one entry.
export { applyTheme, registerComponents } from "@docen/editor";

// Re-export the i18n API so a Vue app registers locales from the same entry —
// mirrors the @docen/editor public surface (registerTranslation merges into
// the built-in en/zh-CN tables; availableLanguages drives the Options dropdown
// and the status-bar language cycle).
export { availableLanguages, registerLocalization, registerTranslation, t } from "@docen/editor";
export type { AdditionalLanguage, LanguageOption, LocalizationInfo } from "@docen/editor";

// Add-in type for the template-ref addAddin() method (runtime registration of
// command/pane/ribbon add-ins that can't cross the `addins` attribute boundary).
export type { DocenAddin } from "@docen/editor";
