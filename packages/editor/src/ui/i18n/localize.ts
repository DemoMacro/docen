/**
 * Lightweight i18n for @docen/ui component-internal strings — close button
 * titles, empty states, ARIA labels. Business strings (ribbon commands,
 * property labels, outline titles) are translated by the editor package and
 * passed in as plain strings; this module never touches them.
 *
 * Modelled on Shoelace's localize: register translations, resolve the locale
 * from the nearest `<docen-workspace lang>` (falling back to `<html lang>`),
 * and re-render on language change. Vanilla HTMLElement — no framework dep.
 *
 * Why not i18next / @formatjs: the built-in string surface is tiny (a handful
 * of labels), so a registry + `<html lang>` observer is enough. The editor
 * package is free to use any i18n solution for business strings.
 */

/** A single locale's translation table. `$code` is the BCP-47 tag. */
export interface DocenTranslation {
  /** BCP-47 language tag, e.g. "zh-CN". */
  readonly $code: string;
  /** Display name (informational), e.g. "中文（简体）". */
  readonly $name?: string;
  /** Text direction. Defaults to "ltr". */
  readonly $dir?: "ltr" | "rtl";
  readonly [key: string]: string | undefined;
}

// Seed the built-in component locales (nav/properties/outline empty states…)
// inline at module load. `t()` callers import this module directly, so the seed
// must live here; inlining (vs importing ./locales/*) avoids a dev-server
// module-instance issue where the imported locale object resolved empty.
const translations = new Map<string, DocenTranslation>([
  [
    "en",
    {
      $code: "en",
      $name: "English",
      $dir: "ltr",
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
      $code: "zh-CN",
      $name: "中文（简体）",
      $dir: "ltr",
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
const listeners = new Set<() => void>();
let htmlObserver: MutationObserver | null = null;

/**
 * Register a translation table, **merging** into any existing table for the
 * same `$code`. Component-internal defaults (nav/properties/outline empty
 * states) are seeded inline in this module; the editor package registers
 * business strings (ribbon/header/pane) under the same codes. A flat set()
 * would let one source clobber the other and drop its keys — merging keeps
 * both. Later registrations win on key conflicts.
 */
export function registerTranslation(translation: DocenTranslation): void {
  const code = translation.$code;
  const existing = translations.get(code);
  translations.set(code, existing ? { ...existing, ...translation } : translation);
  notify();
}

/**
 * Resolve the effective locale for an element: the nearest
 * `<docen-workspace lang>` wins, otherwise `<html lang>`, otherwise "en".
 * Scoped to the known workspace host (not any `[lang]` ancestor) so a
 * workspace can override the page locale — this sidesteps Shoelace's
 * "lang must be on the component itself or <html>" limitation.
 */
export function resolveLang(el: Element | null = document.documentElement): string {
  const host = el?.closest("docen-workspace");
  return host?.getAttribute("lang") || document.documentElement.lang || "en";
}

/** Resolve text direction for an element's locale ("ltr" by default). */
export function resolveDir(el: Element | null = document.documentElement): "ltr" | "rtl" {
  const lang = resolveLang(el);
  for (const code of localeChain(lang)) {
    const dir = translations.get(code)?.$dir;
    if (dir) return dir;
  }
  return "ltr";
}

/**
 * Translate a key for an element's locale, falling back through the locale
 * chain (e.g. "zh-CN" → "zh" → "en"). Returns the key itself if missing.
 */
export function t(key: string, el?: Element | null): string {
  const lang = resolveLang(el);
  for (const code of localeChain(lang)) {
    const value = translations.get(code)?.[key];
    if (value != null) return value;
  }
  return key;
}

/** BCP-47 fallback chain: exact → language family → "en". */
function localeChain(lang: string): string[] {
  const chain = [lang];
  const dash = lang.indexOf("-");
  if (dash > 0) chain.push(lang.slice(0, dash));
  if (!chain.includes("en")) chain.push("en");
  return chain;
}

/**
 * Subscribe to locale or translation changes (e.g. to re-render a component).
 * Triggers on `<html lang>` mutation and on `registerTranslation`.
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
  htmlObserver = new MutationObserver(notify);
  htmlObserver.observe(document.documentElement, {
    attributes: true,
    attributeFilter: ["lang"],
  });
}

function notify(): void {
  for (const listener of listeners) listener();
}
