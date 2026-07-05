import {
  createHighContrastTheme,
  teamsDarkTheme,
  teamsDarkV21Theme,
  teamsHighContrastTheme,
  teamsLightTheme,
  teamsLightV21Theme,
  webDarkTheme,
  webLightTheme,
} from "@fluentui/tokens";
import type { Theme } from "@fluentui/tokens";
import { setTheme } from "@fluentui/web-components";

/** A Fluent theme, or a zero-arg factory that builds one (createHighContrastTheme). */
export type ThemeDefinition = Theme | (() => Theme);

/** Theme registry keyed by themeId. Seeded with the 8 built-in @fluentui/tokens
 *  presets (light / dark / high-contrast + Teams variants); extended at runtime
 *  via {@link registerTheme} — iterate with `builtinThemes.keys()` (no wrapper).
 *  Insertion order = Options-dropdown order. Consumers add brand themes via
 *  createLightTheme/createDarkTheme(brand) (the Office.js fluentThemeData model:
 *  a Fluent Theme object keyed by id). */
export const builtinThemes = new Map<string, ThemeDefinition>([
  ["light", webLightTheme],
  ["dark", webDarkTheme],
  ["high-contrast", createHighContrastTheme],
  ["teams-light", teamsLightTheme],
  ["teams-dark", teamsDarkTheme],
  ["teams-high-contrast", teamsHighContrastTheme],
  ["teams-light-v21", teamsLightV21Theme],
  ["teams-dark-v21", teamsDarkV21Theme],
]);

/** Register a custom theme under `themeId` so `<docen-document theme="…">`
 *  resolves to it. Mirrors Office.js `fluentThemeData` — a Fluent Theme object
 *  keyed by id — and pairs with `@fluentui/tokens` brand factories
 *  (createLightTheme/createDarkTheme). A factory is accepted so parameterized
 *  themes (e.g. createHighContrastTheme) construct on demand. */
export function registerTheme(themeId: string, theme: ThemeDefinition): void {
  builtinThemes.set(themeId, theme);
}

/** Map a `theme` attribute value to a registered themeId, defaulting to light. */
export function resolveTheme(value: string | null | undefined): string {
  return value && builtinThemes.has(value) ? value : "light";
}

/**
 * Apply a registered theme by themeId. Wraps `@fluentui/web-components`
 * `setTheme`; `node` scopes the theme (defaults to the whole document).
 */
export function applyTheme(themeId: string, node?: Document | HTMLElement): void {
  const def = builtinThemes.get(themeId) ?? builtinThemes.get("light")!;
  setTheme(typeof def === "function" ? def() : def, node);
}
