/**
 * Lightweight i18n for UI component-internal strings — close button
 * titles, empty states, ARIA labels. Business strings (ribbon commands,
 * property labels, outline titles) are translated by the editor package and
 * passed in as plain strings; this module never touches them.
 *
 * Modelled on Shoelace's localize, with types aligned to the Office.js
 * unified manifest (`localizationInfo`): register translations, resolve the
 * locale from the nearest `<docen-workspace lang>` (falling back to
 * `<html lang>`), and re-render on language change. Vanilla HTMLElement — no
 * framework dep.
 *
 * Office.js parallel: manifest `localizationInfo` (container: `defaultLanguageTag`
 * + `additionalLanguages[]`) references an external JSON file per language;
 * docen inlines the same key→value table as `AdditionalLanguage.translations`.
 * `Office.context.displayLanguage` is exposed on the host as
 * {@link DocenHost.displayLanguage} (read-only).
 *
 * Why not i18next / @formatjs: the built-in string surface is tiny (a handful
 * of labels), so a registry + `<html lang>` observer is enough. The editor
 * package is free to use any i18n solution for business strings.
 */

/** A single language's translation entry — Office.js manifest
 *  `additionalLanguages[]` element, with `file` inlined as `translations`
 *  (Office.js references an external JSON file; docen inlines the same
 *  key→value table). `$name`/`$dir` are docen extensions (Office.js file
 *  content is pure key→value). */
export interface AdditionalLanguage {
  /** BCP-47 language tag, e.g. "zh-CN" (Office.js `languageTag`). */
  readonly languageTag: string;
  /** The translation table (Office.js `file` content, inlined). */
  readonly translations: Readonly<Record<string, string>>;
  /** Display name (informational), e.g. "中文（简体）". */
  readonly $name?: string;
  /** Text direction. Defaults to "ltr". */
  readonly $dir?: "ltr" | "rtl";
}

/** A localization manifest — Office.js `localizationInfo`. `defaultLanguageTag`
 *  is the fallback locale (localeChain's last link); `additionalLanguages` are
 *  the per-language tables. Office.js references external files; docen inlines
 *  the tables as {@link AdditionalLanguage.translations}.
 *
 *  docen extension: `additionalLanguages` may include an entry for the default
 *  language itself. Office.js keeps default-language strings in the manifest
 *  file, but a docen addin is a plain JS object with no equivalent, so it
 *  contributes its default-locale keys through `additionalLanguages` like any
 *  other language. */
export interface LocalizationInfo {
  /** Fallback locale used when neither `<docen-workspace lang>` nor `<html lang>`
   *  yields a hit, and as localeChain's terminal fallback (Office.js
   *  `defaultLanguageTag`). */
  readonly defaultLanguageTag: string;
  readonly additionalLanguages: readonly AdditionalLanguage[];
}

// Seed the built-in component locales (nav/properties/outline empty states…)
// inline at module load. `t()` callers import this module directly, so the seed
// must live here; inlining (vs importing ./locales/*) avoids a dev-server
// module-instance issue where the imported locale object resolved empty.
// Translations and metadata ($name/$dir) are split into two maps so the lookup
// path stays a plain Record<string,string> (no $-prefixed keys mixed in).
const translations = new Map<string, Record<string, string>>([
  [
    "en",
    {
      "taskPane.close": "Close panel",
      "outline.empty": "No outline entries",
      "properties.empty": "No properties",
      "nav.search": "Search",
      "nav.headings": "Headings",
      "nav.pages": "Pages",
      "nav.results": "Results",
      "findReplace.title": "Find and Replace",
      "findReplace.find": "Find what:",
      "findReplace.replaceWith": "Replace with:",
      "findReplace.matchCase": "Match case",
      "findReplace.wholeWord": "Whole word",
      "findReplace.findNext": "Find Next",
      "findReplace.replace": "Replace",
      "findReplace.replaceAll": "Replace All",
      "findReplace.cancel": "Cancel",
    },
  ],
  [
    "zh-CN",
    {
      "taskPane.close": "关闭面板",
      "outline.empty": "暂无大纲条目",
      "properties.empty": "暂无属性",
      "nav.search": "搜索",
      "nav.headings": "标题",
      "nav.pages": "页面",
      "nav.results": "结果",
      "findReplace.title": "查找和替换",
      "findReplace.find": "查找内容：",
      "findReplace.replaceWith": "替换为：",
      "findReplace.matchCase": "区分大小写",
      "findReplace.wholeWord": "全字匹配",
      "findReplace.findNext": "查找下一个",
      "findReplace.replace": "替换",
      "findReplace.replaceAll": "全部替换",
      "findReplace.cancel": "取消",
    },
  ],
]);
const metadata = new Map<string, { readonly $name?: string; readonly $dir?: "ltr" | "rtl" }>([
  ["en", { $name: "English", $dir: "ltr" }],
  ["zh-CN", { $name: "中文（简体）", $dir: "ltr" }],
]);
/** The fallback locale (localeChain's terminal link). The built-in default is
 *  "en"; `registerLocalization` updates it from `LocalizationInfo.defaultLanguageTag`. */
let defaultLanguageTag = "en";

const listeners = new Set<() => void>();
let htmlObserver: MutationObserver | null = null;

/**
 * Register a single language's translation table, **merging** into any
 * existing table for the same `languageTag`. Component-internal defaults
 * (nav/properties/outline empty states) are seeded inline in this module; the
 * editor package registers business strings (ribbon/header/pane) under the
 * same tags. A flat assignment would let one source clobber the other and drop
 * its keys — merging keeps both. Later registrations win on key conflicts.
 */
export function registerTranslation(entry: AdditionalLanguage): void {
  const tag = entry.languageTag;
  const existing = translations.get(tag);
  translations.set(tag, { ...existing, ...entry.translations });
  const metaExisting = metadata.get(tag);
  metadata.set(tag, {
    $name: entry.$name ?? metaExisting?.$name,
    $dir: entry.$dir ?? metaExisting?.$dir,
  });
  notifyLocaleChange();
}

/**
 * Register a localization manifest (Office.js `localizationInfo`): adopt its
 * `defaultLanguageTag` as the fallback locale, then register every additional
 * language. Addins pass their `localizationInfo` here; the host calls this on
 * `addAddin`. Equivalent to N `registerTranslation` calls plus the fallback update.
 */
export function registerLocalization(info: LocalizationInfo): void {
  if (info.defaultLanguageTag) defaultLanguageTag = info.defaultLanguageTag;
  info.additionalLanguages.forEach(registerTranslation);
}

/** A registered language surfaced as a picker option — `languageTag` plus the
 *  optional display name. Languages with no registered `$name` fall back to
 *  their tag in the UI. */
export interface LanguageOption {
  readonly languageTag: string;
  readonly $name?: string;
}

/**
 * List every registered language for the options language picker. Built-in
 * seeds (`en` / `zh-CN`) and any addin-registered tags appear; a new language
 * lands here automatically once {@link registerTranslation} (or an addin's
 * `localizationInfo`) registers it — that is the supported way to add a locale.
 * The default language sorts first, the rest alphabetically by tag.
 */
export function availableLanguages(): readonly LanguageOption[] {
  return [...translations.keys()]
    .map((languageTag) => ({ languageTag, $name: metadata.get(languageTag)?.$name }))
    .sort((a, b) => {
      if (a.languageTag === defaultLanguageTag) return -1;
      if (b.languageTag === defaultLanguageTag) return 1;
      return a.languageTag.localeCompare(b.languageTag);
    });
}

/**
 * Resolve the effective locale for an element. The element's own `lang`
 * attribute wins (`<docen-document lang>` / `<docen-workspace lang>` /
 * `<html lang>`) — `closest()` can't reach the workspace from the host
 * element or inside its shadow root, so the element's own `lang` is the
 * bridge from `<docen-document lang>` down to the i18n lookup. Otherwise
 * the nearest ancestor `<docen-workspace lang>` (for elements inside a
 * workspace's light-DOM subtree). If the element lives directly in a shadow
 * root whose host carries `lang` (e.g. an options dialog placed in
 * `<docen-document>`'s shadow rather than slotted into the workspace), the
 * shadow host's `lang` is read — `closest()` stops at the shadow boundary.
 * Finally `<html lang>`, then {@link defaultLanguageTag} (manifest
 * `defaultLanguageTag`, "en" by default).
 *
 * Scoped to the known workspace host (not any `[lang]` ancestor) so a
 * workspace can override the page locale — this sidesteps Shoelace's
 * "lang must be on the component itself or <html>" limitation.
 */
export function resolveLang(el: Element | null = document.documentElement): string {
  const ownLang = el?.getAttribute("lang");
  if (ownLang) return ownLang;
  const host = el?.closest("docen-workspace");
  if (host?.hasAttribute("lang")) return host.getAttribute("lang")!;
  const root = el?.getRootNode();
  if (root instanceof ShadowRoot) {
    const hostLang = (root.host as Element | null)?.getAttribute("lang");
    if (hostLang) return hostLang;
  }
  return document.documentElement.lang || defaultLanguageTag;
}

/** Resolve text direction for an element's locale ("ltr" by default). */
export function resolveDir(el: Element | null = document.documentElement): "ltr" | "rtl" {
  const lang = resolveLang(el);
  for (const code of localeChain(lang)) {
    const dir = metadata.get(code)?.$dir;
    if (dir) return dir;
  }
  return "ltr";
}

/**
 * Translate a key for an element's locale, falling back through the locale
 * chain (e.g. "zh-CN" → "zh" → `defaultLanguageTag`). Returns the key itself
 * if missing.
 */
export function t(key: string, el?: Element | null): string {
  const lang = resolveLang(el);
  for (const code of localeChain(lang)) {
    const value = translations.get(code)?.[key];
    if (value != null) return value;
  }
  return key;
}

/** BCP-47 fallback chain: exact → language family → `defaultLanguageTag`. */
function localeChain(lang: string): string[] {
  const chain = [lang];
  const dash = lang.indexOf("-");
  if (dash > 0) chain.push(lang.slice(0, dash));
  if (!chain.includes(defaultLanguageTag)) chain.push(defaultLanguageTag);
  return chain;
}

/**
 * Subscribe to locale or translation changes (e.g. to re-render a component).
 * Triggers on `<html lang>` mutation, on `registerTranslation`, and on
 * {@link notifyLocaleChange} (which a host that drives locale through its own
 * `lang` attribute calls after forwarding it to the workspace).
 * Returns an unsubscribe function.
 */
export function observeLang(listener: () => void): () => void {
  listeners.add(listener);
  ensureHtmlObserver();
  return () => {
    listeners.delete(listener);
  };
}

function ensureHtmlObserver(): void {
  if (htmlObserver || typeof MutationObserver === "undefined") return;
  htmlObserver = new MutationObserver(notifyLocaleChange);
  htmlObserver.observe(document.documentElement, {
    attributes: true,
    attributeFilter: ["lang"],
  });
}

/**
 * Fire every {@link observeLang} listener. Call this after a non-`<html lang>`
 * locale source changes — e.g. `<docen-document>` forwarding its `lang`
 * attribute to the internal `<docen-workspace>` — so observers re-render
 * even though the `MutationObserver` on `documentElement.lang` didn't fire.
 * Also called internally by `registerTranslation` and the html-lang observer.
 */
export function notifyLocaleChange(): void {
  for (const listener of listeners) listener();
}
