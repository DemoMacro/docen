import type { AdditionalLanguage } from "../localize";

/**
 * English (US) — the default and fallback locale. Component-internal strings
 * only; business strings come from the editor package.
 */
export const en: AdditionalLanguage = {
  languageTag: "en",
  $name: "English",
  $dir: "ltr",
  translations: {
    "taskPane.close": "Close panel",
    "outline.empty": "No outline entries",
    "properties.empty": "No properties",
    "nav.search": "Search",
    "nav.headings": "Headings",
    "nav.pages": "Pages",
    "nav.results": "Results",
  },
};
