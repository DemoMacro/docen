/**
 * Office-style layout tokens as CSS custom properties.
 *
 * Command styling (button/toggle/menu colors, hover/pressed) is delegated to
 * Fluent UI's theme via `applyTheme()`; only container/layout metrics and the
 * Ribbon's white-panel / gray-tab-strip / gray-status palette live here.
 */
export const OFFICE_TOKENS_CSS = `
:where(:root) {
  --docen-ribbon-panel-height: 92px;
  --docen-color-bg: #ffffff;
  --docen-color-tab-bg: #f0f0f0;
  --docen-color-status-bg: #f0f0f0;
  --docen-color-text: #444444;
  --docen-color-text-muted: #666666;
  --docen-color-divider: #e2e2e2;
  --docen-color-canvas: #f0f0f0;
  --docen-color-page: #ffffff;
  --docen-page-width: 210mm;
  --docen-page-min-height: 297mm;
  --docen-page-gap: 24px;
  --docen-font-size-ribbon: 12px;
  --docen-font-size-group-label: 10px;
}` as const;

const STYLE_ID = "docen-office-tokens";

/** Stamp Office tokens on `:root` (default) or a scoped element. Idempotent. */
export function injectOfficeTokens(scope: Document | HTMLElement = document): void {
  const root = scope === document ? document.documentElement : scope;
  if (root.querySelector(`style[data-docen="${STYLE_ID}"]`)) return;
  const style = document.createElement("style");
  style.setAttribute("data-docen", STYLE_ID);
  style.textContent = OFFICE_TOKENS_CSS;
  root.append(style);
}
