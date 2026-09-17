// The eight-color swatch row shared by the workspace color palettes (font
// underline color, border color, watermark color). Entries are [hex, i18n
// key] under the `fontDialog.color*` namespace; each palette prepends its own
// automatic/no-color entry in its own value space.
export const SWATCH_COLORS: ReadonlyArray<readonly [string, string]> = [
  ["000000", "colorBlack"],
  ["800000", "colorDarkRed"],
  ["008000", "colorGreen"],
  ["000080", "colorDarkBlue"],
  ["FF0000", "colorRed"],
  ["FF00FF", "colorMagenta"],
  ["FFFF00", "colorYellow"],
  ["00FFFF", "colorCyan"],
];
