import type {
  BorderOptions,
  BordersOptions,
  IndentAttributesProperties,
  ShadingAttributesProperties,
  SpacingProperties,
} from "@office-open/docx";

// ── CSS color helpers ──

/** Common CSS named colors → hex */
const CSS_COLORS: Record<string, string> = {
  black: "#000000",
  white: "#FFFFFF",
  red: "#FF0000",
  green: "#008000",
  blue: "#0000FF",
  yellow: "#FFFF00",
  cyan: "#00FFFF",
  magenta: "#FF00FF",
  gray: "#808080",
  grey: "#808080",
  orange: "#FFA500",
  purple: "#800080",
  pink: "#FFC0CB",
  brown: "#A52A2A",
  lime: "#00FF00",
  navy: "#000080",
  teal: "#008080",
  silver: "#C0C0C0",
  maroon: "#800000",
  olive: "#808000",
  aqua: "#00FFFF",
  fuchsia: "#FF00FF",
  indigo: "#4B0082",
  violet: "#EE82EE",
  coral: "#FF7F50",
  gold: "#FFD700",
  salmon: "#FA8072",
  tomato: "#FF6347",
};

/** Normalize a CSS color value to hex (e.g., "red" → "#FF0000", "#ff0000" → "#FF0000").
 *  Accepts a string (CSS name/hex or bare OOXML hex), or an OOXML ColorOptions
 *  object ({ val, themeColor, themeTint, themeShade }) as emitted for theme
 *  colors — the object form resolves to its val fallback (full theme-color
 *  resolution needs theme.xml, not yet wired). */
export function normalizeColorToHex(color: unknown): string | undefined {
  if (!color) return undefined;
  if (typeof color === "object") {
    const { val } = color as { val?: unknown };
    return val ? normalizeColorToHex(val) : undefined;
  }
  if (typeof color !== "string") return undefined;
  // OOXML "auto" has no CSS equivalent — skip (leave color unset).
  if (color === "auto") return undefined;
  if (color.startsWith("#"))
    return color.length === 4
      ? `#${color[1]}${color[1]}${color[2]}${color[2]}${color[3]}${color[3]}`.toUpperCase()
      : color.toUpperCase();
  // OOXML stores bare hex without "#" (e.g., "FF0000") — add the prefix.
  if (/^[0-9A-Fa-f]{6}$/.test(color)) return `#${color.toUpperCase()}`;
  if (/^[0-9A-Fa-f]{3}$/.test(color))
    return `#${color[0]}${color[0]}${color[1]}${color[1]}${color[2]}${color[2]}`.toUpperCase();
  const hex = CSS_COLORS[color.toLowerCase()];
  return hex ?? undefined;
}

/** Resolve a font value (string or OOXML rFonts { ascii, eastAsia, hAnsi, cs }) to a CSS family name. */
export function resolveFontName(font: unknown): string | null {
  if (!font) return null;
  if (typeof font === "string") return font;
  if (typeof font === "object") {
    const f = font as { ascii?: string; hAnsi?: string; eastAsia?: string };
    return f.ascii || f.hAnsi || f.eastAsia || null;
  }
  return null;
}

// ── Unit conversion helpers ──
// office-open stores native values (twips, points); CSS conversion lives here.

/** CSS value (e.g., "18pt") → twip number. 1 pt = 20 twips, 1 px = 15 twips (96 DPI). */
export function cssToTwip(value: string | undefined): number | undefined {
  if (!value) return undefined;
  const match = value.match(/^([\d.]+)(pt|px|em|cm|in)?$/);
  if (!match) return undefined;
  const num = parseFloat(match[1]);
  const unit = match[2] ?? "pt";
  switch (unit) {
    case "pt":
      return Math.round(num * 20);
    case "px":
      return Math.round(num * 15);
    case "in":
      return Math.round(num * 1440);
    case "cm":
      return Math.round(num * 567);
    case "em":
      return Math.round(num * 240);
    default:
      return Math.round(num * 20);
  }
}

/** Twip value → CSS string (e.g., 360 → "18pt"). */
export function twipToCss(value: number | string | undefined): string | null {
  if (value == null) return null;
  if (typeof value === "string") return value;
  return `${value / 20}pt`;
}

// ── Alignment mapping ──
// OOXML AlignmentType has no "justify" — both-sides value is "both".

const ALIGNMENT_TO_CSS: Record<string, string> = {
  left: "left",
  center: "center",
  right: "right",
  start: "start",
  end: "end",
  both: "justify",
  distribute: "justify",
};

const CSS_TO_ALIGNMENT: Record<string, string> = {
  left: "left",
  center: "center",
  right: "right",
  start: "start",
  end: "end",
  justify: "both",
};

/** OOXML alignment → CSS text-align. */
export function alignmentToCss(alignment: string | null | undefined): string | null {
  if (!alignment) return null;
  return ALIGNMENT_TO_CSS[alignment] ?? null;
}

/** CSS text-align → OOXML alignment. */
export function alignmentFromCss(css: string | null | undefined): string | null {
  if (!css) return null;
  return CSS_TO_ALIGNMENT[css] ?? null;
}

// ── Shading mapping ──

/** Shading.fill → CSS background-color hex. */
export function shadingToCss(
  shading: ShadingAttributesProperties | null | undefined,
): string | null {
  if (!shading?.fill) return null;
  return normalizeColorToHex(shading.fill) ?? null;
}

/** CSS background-color → ShadingAttributesProperties (fill normalized to hex). */
export function shadingFromCss(css: string | null | undefined): ShadingAttributesProperties | null {
  const hex = normalizeColorToHex(css ?? undefined);
  return hex ? { fill: hex, type: "clear" } : null;
}

// ── Line spacing mapping ──
// OOXML spacing.line: 240 = single, 360 = 1.5.
// lineRule "exact"/"atLeast" → fixed pt; "auto"/undefined → multiple of 240.

export function lineSpacingToCss(spacing: SpacingProperties | null | undefined): string | null {
  if (!spacing?.line) return null;
  const rule = spacing.lineRule;
  if (rule === "exact" || rule === "exactly" || rule === "atLeast") {
    return `${spacing.line / 20}pt`;
  }
  const multiple = Number((spacing.line / 240).toFixed(2));
  return String(multiple);
}

// ── Font size mapping ──
// office-open size is in POINTS (new convention).

export function sizeToCss(size: number | null | undefined): string | null {
  if (size == null) return null;
  return `${size}pt`;
}

export function sizeFromCss(css: string | null | undefined): number | null {
  if (!css) return null;
  const m = css.match(/^([\d.]+)(pt|px)?$/);
  if (!m) return null;
  const num = parseFloat(m[1]);
  const unit = m[2] ?? "pt";
  // 1px = 0.75pt at 96 DPI (DOCX stores points, not pixels)
  return unit === "px" ? num * 0.75 : num;
}

// ── Character spacing mapping ──
// OOXML characterSpacing is in twips (1/20 pt).

export function characterSpacingToCss(spacing: number | null | undefined): string | null {
  if (spacing == null) return null;
  return `${spacing / 20}pt`;
}

export function characterSpacingFromCss(css: string | null | undefined): number | null {
  if (!css) return null;
  const m = css.match(/^(-?[\d.]+)pt$/);
  return m ? Math.round(parseFloat(m[1]) * 20) : null;
}

// ── Border rendering ──

/** Render a BorderOptions to CSS string. OOXML border.size is in eighths of a point. */
export function renderBorderCSS(border: BorderOptions): string | null {
  if (!border || border.style === "none") return null;
  const size = border.size != null ? `${border.size / 8}pt` : "1pt";
  const styleMap: Record<string, string> = {
    single: "solid",
    dashed: "dashed",
    dotted: "dotted",
    double: "double",
    dotDash: "dashed",
    dotDotDash: "dotted",
  };
  const cssStyle = styleMap[border.style || "single"] || "solid";
  // OOXML color "auto" has no CSS equivalent and bare hex needs a "#" prefix —
  // normalize to hex, or omit the color entirely (CSS defaults to currentColor).
  const hex = border.color && border.color !== "auto" ? normalizeColorToHex(border.color) : null;
  return hex ? `${cssStyle} ${size} ${hex}` : `${cssStyle} ${size}`;
}

// ── Style rendering (consume nested office-open attrs) ──

interface ParagraphStyleShape {
  alignment?: string | null;
  indent?: IndentAttributesProperties | null;
  spacing?: SpacingProperties | null;
  shading?: ShadingAttributesProperties | null;
  border?: BordersOptions | null;
}

/**
 * Compute all paragraph-level CSS styles from nested attrs.
 * Shared by Paragraph and Heading extensions for node-level renderHTML.
 * Attrs store office-open native values; mappers here convert to CSS.
 */
export function renderParagraphStyles(attrs: Record<string, unknown>): string[] {
  const a = attrs as ParagraphStyleShape;
  const styles: string[] = [];

  const align = alignmentToCss(a.alignment);
  if (align) styles.push(`text-align:${align}`);

  if (a.indent) {
    const left = twipToCss(a.indent.left);
    if (left) styles.push(`margin-left:${left}`);
    const right = twipToCss(a.indent.right);
    if (right) styles.push(`margin-right:${right}`);
    if (a.indent.firstLine != null) {
      const fl = twipToCss(a.indent.firstLine);
      if (fl) styles.push(`text-indent:${fl}`);
    } else if (a.indent.hanging != null) {
      const h = twipToCss(a.indent.hanging);
      if (h) styles.push(`text-indent:-${h}`);
    } else if (a.indent.firstLineChars != null) {
      styles.push(`text-indent:${a.indent.firstLineChars / 100}em`);
    }
  }

  if (a.spacing) {
    const before = twipToCss(a.spacing.before);
    if (before) styles.push(`margin-top:${before}`);
    const after = twipToCss(a.spacing.after);
    if (after) styles.push(`margin-bottom:${after}`);
    const lh = lineSpacingToCss(a.spacing);
    if (lh) styles.push(`line-height:${lh}`);
  }

  const bg = shadingToCss(a.shading);
  if (bg) styles.push(`background-color:${bg}`);

  if (a.border) {
    const sides: Array<[string, BorderOptions | undefined]> = [
      ["top", a.border.top],
      ["bottom", a.border.bottom],
      ["left", a.border.left],
      ["right", a.border.right],
    ];
    for (const [side, b] of sides) {
      const css = b ? renderBorderCSS(b) : null;
      if (css) styles.push(`border-${side}:${css}`);
    }
  }

  return styles;
}

interface CellStyleShape {
  noWrap?: boolean | null;
  shading?: ShadingAttributesProperties | null;
  verticalAlign?: string | null;
}

/**
 * Compute table cell CSS styles from nested attrs.
 * Shared by TableCell and TableHeader extensions.
 */
export function renderTableCellStyles(attrs: Record<string, unknown>): string[] {
  const a = attrs as CellStyleShape;
  const styles: string[] = [];

  if (a.noWrap) styles.push("white-space:nowrap");

  const bg = shadingToCss(a.shading);
  if (bg) styles.push(`background-color:${bg}`);

  if (a.verticalAlign) styles.push(`vertical-align:${a.verticalAlign}`);

  return styles;
}

// ── Element parsers (CSS → office-open native, for parseHTML) ──
// Shared by Paragraph and Heading: each attr's parseHTML calls one of these.

/** Parse text-align → OOXML alignment. */
export function alignmentFromElement(el: HTMLElement): string | null {
  return alignmentFromCss(el.style.textAlign || null);
}

/** Parse margin-left/right + text-indent → OOXML indent (twips). */
export function indentFromElement(el: HTMLElement): IndentAttributesProperties | null {
  const indent: IndentAttributesProperties = {};
  const left = cssToTwip(el.style.marginLeft);
  if (left) indent.left = left;
  const right = cssToTwip(el.style.marginRight);
  if (right) indent.right = right;
  const ti = el.style.textIndent;
  if (ti) {
    if (ti.startsWith("-")) {
      const h = cssToTwip(ti.slice(1));
      if (h) indent.hanging = h;
    } else {
      const f = cssToTwip(ti);
      if (f) indent.firstLine = f;
    }
  }
  return Object.keys(indent).length > 0 ? indent : null;
}

/** Parse margin-top/bottom + line-height → OOXML spacing (twips). */
export function spacingFromElement(el: HTMLElement): SpacingProperties | null {
  const spacing: SpacingProperties = {};
  const before = cssToTwip(el.style.marginTop);
  if (before) spacing.before = before;
  const after = cssToTwip(el.style.marginBottom);
  if (after) spacing.after = after;
  const lh = el.style.lineHeight;
  if (lh) {
    const m = lh.match(/^([\d.]+)(pt|px)?$/);
    if (m) {
      const num = parseFloat(m[1]);
      if (m[2]) {
        // absolute (pt/px) → exact line spacing in twips
        spacing.line = Math.round(num * (m[2] === "px" ? 15 : 20));
        spacing.lineRule = "exact";
      } else {
        // bare number → multiple of 240
        spacing.line = Math.round(num * 240);
        spacing.lineRule = "auto";
      }
    }
  }
  return Object.keys(spacing).length > 0 ? spacing : null;
}

/** Parse border-* → OOXML BordersOptions. */
export function bordersFromElement(el: HTMLElement): BordersOptions | null {
  const borders: BordersOptions = {};
  const sides: Array<[keyof BordersOptions, string]> = [
    ["top", el.style.borderTop],
    ["bottom", el.style.borderBottom],
    ["left", el.style.borderLeft],
    ["right", el.style.borderRight],
  ];
  for (const [side, css] of sides) {
    if (!css || css === "initial" || css === "none") continue;
    const m = css.match(/^(none|solid|dashed|dotted|double)\s+([\d.]+pt)\s+(.+)$/);
    if (!m) continue;
    const styleMap: Record<string, BorderOptions["style"]> = {
      solid: "single",
      dashed: "dashed",
      dotted: "dotted",
      double: "double",
    };
    borders[side] = {
      style: styleMap[m[1]] ?? "single",
      size: Math.round(parseFloat(m[2]) * 8),
      color: m[3],
    };
  }
  return Object.keys(borders).length > 0 ? borders : null;
}

/** Parse background-color → OOXML shading. */
export function shadingFromElement(el: HTMLElement): ShadingAttributesProperties | null {
  return shadingFromCss(el.style.backgroundColor || null);
}
