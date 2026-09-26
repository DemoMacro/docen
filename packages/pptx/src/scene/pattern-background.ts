/** OOXML preset pattern fills → a page-sized SVG image. The projection keeps
 *  the painter's single image-background contract while preserving the tiled
 *  geometry and both pattern colors. */

import type { FillOptions } from "@office-open/core/drawing";

type PatternFill = Extract<FillOptions, { type: "pattern" }>;

const TILE = 8;

/** A repeatable SVG fragment in the common 8×8 pattern tile. */
function tileOf(pattern: PatternFill["pattern"]): string {
  const dots = (points: readonly [number, number][]) =>
    points.map(([x, y]) => `<rect x="${x}" y="${y}" width="1" height="1"/>`).join("");
  const line = (d: string, width = 1) => `<path d="${d}" stroke-width="${width}"/>`;
  const percent = (fraction: number) => {
    const count = Math.max(1, Math.round(TILE * TILE * fraction));
    return dots(
      Array.from(
        { length: count },
        (_, index) =>
          [(index * 5) % TILE, Math.floor((index * 5) / TILE) % TILE] as [number, number],
      ),
    );
  };

  switch (pattern) {
    case "percent5":
      return percent(0.05);
    case "percent10":
      return percent(0.1);
    case "percent20":
      return percent(0.2);
    case "percent25":
      return percent(0.25);
    case "percent30":
      return percent(0.3);
    case "percent40":
      return percent(0.4);
    case "percent50":
      return percent(0.5);
    case "percent60":
      return percent(0.6);
    case "percent70":
      return percent(0.7);
    case "percent75":
      return percent(0.75);
    case "percent80":
      return percent(0.8);
    case "percent90":
      return percent(0.9);
    case "horizontal":
      return line("M0 .5H8");
    case "vertical":
      return line("M.5 0V8");
    case "lightHorizontal":
      return line("M0 .25H8", 0.5);
    case "lightVertical":
      return line("M.25 0V8", 0.5);
    case "darkHorizontal":
      return line("M0 .5H8 M0 4.5H8", 2);
    case "darkVertical":
      return line("M.5 0V8 M4.5 0V8", 2);
    case "narrowHorizontal":
      return line("M0 .5H8 M0 2.5H8 M0 4.5H8 M0 6.5H8");
    case "narrowVertical":
      return line("M.5 0V8 M2.5 0V8 M4.5 0V8 M6.5 0V8");
    case "dashedHorizontal":
      return line("M0 .5H4 M4 4.5H8");
    case "dashedVertical":
      return line("M.5 0V4 M4.5 4V8");
    case "cross":
      return line("M0 .5H8 M.5 0V8");
    case "downDiagonal":
      return line("M-2 -2L10 10 M2 -2L14 10 M-2 2L10 14", 0.5);
    case "upDiagonal":
      return line("M-2 10L10 -2 M2 14L14 2 M-2 2L2 -2", 0.5);
    case "lightDownDiagonal":
      return line("M0 0L8 8 M-2 2L6 10 M2 -2L10 6", 0.5);
    case "lightUpDiagonal":
      return line("M0 8L8 0 M-2 6L6 -2 M2 10L10 2", 0.5);
    case "darkDownDiagonal":
      return line("M0 0L8 8 M0 -4L12 8 M-4 0L8 12", 2);
    case "darkUpDiagonal":
      return line("M0 8L8 0 M-4 8L8 -4 M0 12L12 0", 2);
    case "wideDownDiagonal":
      return line("M0 2L6 8 M2 0L8 6 M0 6L2 8 M6 0L8 2 M4 -2L12 6 M-2 4L6 12", 0.5);
    case "wideUpDiagonal":
      return line("M0 6L6 0 M2 8L8 2 M0 2L2 0 M6 8L8 6 M-2 4L6 -4 M4 12L12 4", 0.5);
    case "dashedDownDiagonal":
      return line("M0 0L4 4 M4 0L8 4 M0 4L4 8 M4 4L8 8", 0.5);
    case "dashedUpDiagonal":
      return line("M0 4L4 0 M4 8L8 4 M0 8L4 4 M4 4L8 0", 0.5);
    case "diagonalCross":
      return line("M0 0L8 8 M8 0L0 8");
    case "smallChecker":
      return '<rect width="2" height="2"/><rect x="2" y="2" width="2" height="2"/><rect x="4" y="0" width="2" height="2"/><rect x="6" y="2" width="2" height="2"/>';
    case "largeChecker":
      return '<rect width="4" height="4"/><rect x="4" y="4" width="4" height="4"/>';
    case "smallGrid":
      return line("M0 .5H8 M.5 0V8 M0 4.5H8 M4.5 0V8", 0.5);
    case "largeGrid":
      return line("M0 .5H8 M.5 0V8", 2);
    case "dotGrid":
      return dots([
        [1, 1],
        [5, 1],
        [3, 3],
        [7, 3],
        [1, 5],
        [5, 5],
        [3, 7],
        [7, 7],
      ]);
    case "smallConfetti":
      return line("M1 1L2 2 M5 1L6 2 M1 5L2 6 M5 5L6 6 M3 3L4 4 M7 7L8 8", 0.5);
    case "largeConfetti":
      return line("M0 0L2 2 M4 1L6 3 M1 4L3 6 M5 5L7 7 M2 2L4 4", 1);
    case "horizontalBrick":
      return `${line("M0 .5H8 M0 4.5H8")} ${line("M.5 0V4 M4.5 4V8")}`;
    case "diagonalBrick":
      return `${line("M0 2L4 6 M4 -2L8 2 M0 6L2 8 M6 4L8 6")} ${line("M.5 0V2 M4.5 4V6")}`;
    case "solidDiamond":
      return '<path d="M4 0L8 4L4 8L0 4Z"/>';
    case "openDiamond":
      return '<path d="M4 .75L7.25 4L4 7.25L.75 4Z" fill="none" stroke-width="1"/>';
    case "dottedDiamond":
      return dots([
        [4, 0],
        [0, 4],
        [8, 4],
        [4, 8],
        [2, 2],
        [6, 2],
        [2, 6],
        [6, 6],
      ]);
    case "plaid":
      return line("M0 .5H8 M0 4.5H8 M.5 0V8 M4.5 0V8", 0.5);
    case "sphere":
      return '<circle cx="4" cy="4" r="3"/><path d="M2 2A4 4 0 0 1 6 2" fill="none" stroke="#FFFFFF" stroke-width=".5"/>';
    case "weave":
      return line("M0 2C2 0 6 0 8 2 M0 6C2 4 6 4 8 6 M2 0V8 M6 0V8", 0.5);
    case "divot":
      return line("M1 3Q4 0 7 3 M1 7Q4 4 7 7", 0.75);
    case "shingle":
      return line("M0 5Q2 1 4 5 M4 5Q6 1 8 5 M2 1V5 M6 1V5", 0.5);
    case "wave":
      return line("M0 4Q2 1 4 4T8 4", 1);
    case "trellis":
      return line("M0 0L8 8 M8 0L0 8 M0 4L4 0 M4 8L8 4", 0.5);
    case "zigZag":
      return line("M0 2L2 0L4 2L6 0L8 2 M0 6L2 4L4 6L6 4L8 6", 0.75);
  }
}

/** The pattern tile repeated over the slide, flattened to an SVG data URL. */
export function patternBackgroundOf(
  fill: PatternFill,
  widthPx: number,
  heightPx: number,
): string | undefined {
  const foreground = colorOf(fill.foregroundColor) ?? "000000";
  const background = colorOf(fill.backgroundColor);
  const tile = tileOf(fill.pattern);
  const svg = [
    '<svg xmlns="http://www.w3.org/2000/svg"',
    `width="${widthPx}" height="${heightPx}" viewBox="0 0 ${widthPx} ${heightPx}">`,
    "<defs>",
    `<pattern id="pattern" width="${TILE}" height="${TILE}" patternUnits="userSpaceOnUse">`,
    background ? `<rect width="${TILE}" height="${TILE}" fill="#${background}"/>` : "",
    `<g fill="#${foreground}" stroke="#${foreground}">${tile}</g>`,
    "</pattern>",
    "</defs>",
    '<rect width="100%" height="100%" fill="url(#pattern)"/>',
    "</svg>",
  ].join("");
  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
}

function colorOf(value: unknown): string | undefined {
  if (typeof value === "string") return value.replace("#", "").toUpperCase();
  if (!value || typeof value !== "object") return undefined;
  const solid = value as { type?: string; color?: unknown };
  if (solid.type !== "solid") return undefined;
  if (typeof solid.color === "string") return solid.color.replace("#", "").toUpperCase();
  if (solid.color && typeof solid.color === "object" && "value" in solid.color) {
    const value = (solid.color as { value?: unknown }).value;
    return typeof value === "string" ? value.replace("#", "").toUpperCase() : undefined;
  }
  return undefined;
}
