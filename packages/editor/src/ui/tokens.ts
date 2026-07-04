/**
 * Office-style layout tokens as CSS custom properties.
 *
 * Color surfaces reuse Fluent UI's design tokens (`--colorNeutralBackground1`,
 * `--colorNeutralForeground2`, …) so the chrome follows the active light/dark
 * theme set via {@link applyTheme} automatically — docen does not redefine the
 * palette. Only Office-specific layout metrics and the page-paper / editing-mark
 * inks stay docen-defined (no Fluent equivalent). Fallbacks preserve the
 * pre-Fluent palette until `applyTheme` injects the tokens.
 *
 * Fluent UI web-components expose 470+ tokens as global CSS custom properties
 * (verified: colorNeutralBackground1=#ffffff light / #292929 dark); the values
 * resolve per-theme, so `var(--colorNeutralBackground1)` is all a docen
 * component needs to be theme-aware.
 */
export const OFFICE_TOKENS_CSS = `
:where(:root) {
  /* Office layout metrics — no Fluent equivalent. */
  --docen-ribbon-panel-height: 92px;
  --docen-page-width: 210mm;
  --docen-page-min-height: 297mm;
  --docen-page-gap: 24px;
  --docen-font-size-ribbon: 12px;
  --docen-font-size-group-label: 10px;

  /* Color surfaces — reuse Fluent UI tokens (auto light/dark via applyTheme). */
  --docen-color-bg: var(--colorNeutralBackground1, #ffffff);
  --docen-color-canvas: var(--colorNeutralBackground4, #f0f0f0);
  --docen-color-tab-bg: var(--colorNeutralBackground4, #f0f0f0);
  --docen-color-status-bg: var(--colorNeutralBackground4, #f0f0f0);
  --docen-color-hover: var(--colorNeutralBackground1Hover, #f5f5f5);
  --docen-color-subtle-background-hover: var(--colorSubtleBackgroundHover, #f5f5f5);
  --docen-color-text: var(--colorNeutralForeground2, #444444);
  --docen-color-text-1: var(--colorNeutralForeground1, #242424);
  --docen-color-text-3: var(--colorNeutralForeground3, #616161);
  --docen-color-text-muted: var(--colorNeutralForeground3, #666666);
  --docen-color-divider: var(--colorNeutralStroke2, #e2e2e2);
  --docen-color-stroke-1: var(--colorNeutralStroke1, #c7c7c7);
  --docen-color-brand: var(--colorBrandBackground, #0078d4);
  --docen-color-accent: var(--colorBrandBackground, #0f6cbd);

  /* Page paper + editing marks are document annotations, not chrome — they
     stay a fixed ink tone regardless of the UI theme (Word keeps a white page
     in dark mode by default). */
  --docen-color-page: #ffffff;
  --docen-color-crop: #adadad;
  --docen-color-marks: #767676;
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
