import { webDarkTheme, webLightTheme } from "@fluentui/tokens";
import { setTheme } from "@fluentui/web-components";

export type ThemeMode = "light" | "dark";

/**
 * Apply the Fluent UI color theme. Wraps `@fluentui/web-components` `setTheme`
 * with an Office-style light/dark preset; `node` scopes the theme (defaults to
 * the whole document).
 */
export function applyTheme(mode: ThemeMode, node?: Document | HTMLElement): void {
  setTheme(mode === "dark" ? webDarkTheme : webLightTheme, node);
}
