import { sectionMarginDefaults, sectionPageSizeDefaults } from "@office-open/docx";
import type {
  BorderOptions,
  BordersOptions,
  IndentAttributesProperties,
  ShadingAttributesProperties,
  SpacingProperties,
  TableCellOptions,
  TableFloatOptions,
  TableWidthProperties,
} from "@office-open/docx";

// ── Tiptap attr factory ──

/** Factory for a Tiptap attr that carries an office-open native value: never
 *  parsed from HTML nor rendered to it, defaulting to null (ProseMirror stores
 *  every declared attr). Shared by every extension carrying OOXML attrs
 *  (paragraph/heading/table/table-cell/…). */
export const attrNative = () => ({ default: null, parseHTML: () => null, rendered: false });

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
 *  object ({ val, themeColor, themeTint, themeShade }) — the object form
 *  resolves to its val (the RGB fallback Word stores alongside themeColor) for
 *  CSS rendering. The themeColor/tint/shade are preserved verbatim in the attrs
 *  and round-trip back to the DOCX (see text-style/paragraph parseDocx), so
 *  theme semantics survive even though only val is rendered here. A pure theme
 *  reference with no val (rare — Word usually stores both) would need theme.xml
 *  to resolve and is left unset. */
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

/** Resolve a font value to a CSS font-family list with an eastAsia fallback.
 *  Mirrors Word's per-Unicode-range font selection (MS-OE376): Basic Latin uses
 *  the ascii font, CJK Ideographs use eastAsia. CSS can't split a run by range,
 *  so list ascii first (it carries Basic Latin glyphs) then eastAsia — the
 *  browser falls back to eastAsia for CJK chars the ascii font lacks. Without
 *  this, CJK text renders in the ascii font (e.g. Times) instead of the
 *  document's CJK font (e.g. SimSun). Use this for font-family rendering;
 *  resolveFontName (single name) stays for UI/parse. */
export function resolveFontFamilyCss(font: unknown): string | null {
  if (!font) return null;
  if (typeof font === "string") return font;
  if (typeof font === "object") {
    const f = font as { ascii?: string; hAnsi?: string; eastAsia?: string };
    const ascii = f.ascii || f.hAnsi;
    const ea = f.eastAsia;
    if (ascii && ea && ascii !== ea) return `"${ascii}","${ea}"`;
    return ascii || ea || null;
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
// lineRule "exact"/"atLeast" → fixed pt; "auto"/undefined → a multiple of single.

export function lineSpacingToCss(
  spacing: SpacingProperties | null | undefined,
  snapToGrid: boolean | null = null,
): string | null {
  if (!spacing?.line) return null;
  const rule = spacing.lineRule;
  if (rule === "exact" || rule === "exactly" || rule === "atLeast") {
    return `${Number(spacing.line) / 20}pt`;
  }
  // lineRule "auto": `line` is 240ths of the font's single-line height. The
  // single-line metric is the font's `line-height: normal` (ascent + descent +
  // line-gap) — exactly Word's single line. CSS calc() cannot reference
  // `normal`, so the editor measures it per font and sets --docen-font-metric
  // (a ratio; 1.2 fallback for static HTML export).
  // Per ECMA-376 (§17.3.1.87 snapToGrid — "align to document grid"): when
  // snapToGrid is on (default under a docGrid), every line snaps to the grid —
  // line height = the font's natural metric + the grid pitch, and the spacing
  // multiple is ABSORBED (a 1.5×/double line renders at single-grid + pitch, not
  // multiple×natural). snapToGrid===false (e.g. header/footer) applies the
  // multiple. Verified via a Word-generated PDF: single 微软雅黑-Bold 12pt ≈34pt =
  // natural + pitch 17pt; a double-spaced 16pt CJK line ≈38pt (single-grid scale).
  // `1em` resolves against the paragraph's INHERITED font-size (the container
  // default, e.g. 14pt), under-counting paragraphs whose runs are larger — a
  // 42pt heading rendered at the line-height of 14pt. A line box is as tall as
  // its tallest run, so the editor injects --docen-line-base (= the paragraph's
  // max run size) per paragraph; `1em` fallback covers static HTML export and
  // any paragraph missing the decoration.
  if (snapToGrid === false) {
    const multiple = Number((Number(spacing.line) / 240).toFixed(2));
    return `calc(var(--docen-font-metric, 1.2) * ${multiple} * var(--docen-line-base, 1em))`;
  }
  return `calc(var(--docen-font-metric, 1.2) * var(--docen-line-base, 1em) + var(--docen-line-pitch, 0pt))`;
}

// ── Section geometry → CSS ──
// OOXML section properties (CT_SectPr): page size/margin + document grid.
// The grid is section-level (each section carries its own linePitch), so the
// line-pitch must be injected per-section, not as a single document-root var.

/** twips → mm (1 in = 1440 tw = 25.4 mm), 2dp string. */
export function twipsToMm(twips: number): string {
  return `${((twips / 1440) * 25.4).toFixed(2)}mm`;
}

/** Resolve a section's printable page dimensions (twips), honoring orientation.
 *  A landscape section commonly stores portrait dims (w<h) with
 *  `orientation: "landscape"` — swap width/height so width is the larger edge.
 *  Falls back to the engine's default page size (@office-open/docx
 *  `sectionPageSizeDefaults` = A4) when the size is absent or non-numeric — the
 *  engine's `stringifySectionPropertiesXml` fills an empty sectPr the same way,
 *  so edit-time geometry matches render/measure/generate/export. */
export function resolvePageSize(size: unknown): { width: number; height: number } {
  const fallback = { width: sectionPageSizeDefaults.WIDTH, height: sectionPageSizeDefaults.HEIGHT };
  if (!size || typeof size !== "object") return fallback;
  const s = size as { width?: unknown; height?: unknown; orientation?: unknown };
  const w = typeof s.width === "number" ? s.width : undefined;
  const h = typeof s.height === "number" ? s.height : undefined;
  if (w == null || h == null) return fallback;
  return s.orientation === "landscape" && w < h ? { width: h, height: w } : { width: w, height: h };
}

/** Page margin (twips) → CSS padding. Sides absent or non-numeric fall back to
 *  the engine's default margins (@office-open/docx `sectionMarginDefaults`:
 *  top/bottom 1440tw, left/right 1800tw — MS Office zh-CN "Normal"), matching
 *  the page-size fallback above and the engine's empty-sectPr behavior. Margins
 *  are left as-is: office-open returns them already rotated for a landscape
 *  section. */
export function sectionMarginCss(margin: unknown): string {
  const def = sectionMarginDefaults;
  const m = (margin && typeof margin === "object" ? margin : {}) as {
    top?: unknown;
    right?: unknown;
    bottom?: unknown;
    left?: unknown;
  };
  const num = (v: unknown, d: number): number => (typeof v === "number" ? v : d);
  const sides = [
    num(m.top, def.TOP),
    num(m.right, def.RIGHT),
    num(m.bottom, def.BOTTOM),
    num(m.left, def.LEFT),
  ];
  return `padding:${sides.map(twipsToMm).join(" ")}`;
}

/** Document grid linePitch (twips) → container CSS line-height.
 *  Per ECMA-376 snapToGrid, the grid linePitch is ADDED to each line (single-
 *  line paragraphs inherit it). So the container line-height (inherited by
 *  paragraphs without their own spacing) is the font's `normal` metric + the
 *  grid pitch, set only when the grid type enables line-snapping (not
 *  "default"). The --docen-line-pitch var carries the pitch so per-paragraph
 *  lineSpacingToCss can add the same pitch; --docen-font-metric is injected per
 *  paragraph by the editor; 1.2 is the fallback for static HTML export. */
export function sectionLinePitchCss(grid: unknown): string[] {
  if (!grid || typeof grid !== "object") return [];
  const g = grid as { linePitch?: unknown; type?: unknown };
  // OOXML: docGrid @type omitted or "default" = NO grid (lines do not snap to
  // @linePitch). @type is absent on many Western docs that still carry a
  // @linePitch; treating absent as snapping added 18pt to every line and
  // inflated pagination ~60%.
  if (!g.type || g.type === "default" || typeof g.linePitch !== "number" || !g.linePitch) return [];
  const pitch = `${(g.linePitch / 20).toFixed(2)}pt`;
  // Container line-height for paragraphs WITHOUT their own spacing: the font's
  // `normal` metric × 1 (single) + the grid pitch.
  return [
    `line-height:calc(var(--docen-font-metric, 1.2) * 1em + ${pitch})`,
    `--docen-line-pitch:${pitch}`,
  ];
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
  // OOXML "nil" and a missing style both mean "no border". Only "none" was
  // guarded before, so "nil" fell through to the styleMap default ("solid") and
  // painted a ghost border on every paragraph carrying an empty <w:pBdr>.
  if (!border || !border.style || border.style === "none" || border.style === "nil") {
    return null;
  }
  const size = border.size != null ? `${border.size / 8}pt` : "1pt";
  const styleMap: Record<string, string> = {
    single: "solid",
    dashed: "dashed",
    dotted: "dotted",
    double: "double",
    dotDash: "dashed",
    dotDotDash: "dotted",
    dashSmallGap: "dashed",
  };
  const cssStyle = styleMap[border.style || "single"] || "solid";
  // OOXML color "auto" has no CSS equivalent and bare hex needs a "#" prefix —
  // normalize to hex, or omit the color entirely (CSS defaults to currentColor).
  const hex = border.color && border.color !== "auto" ? normalizeColorToHex(border.color) : null;
  return hex ? `${cssStyle} ${size} ${hex}` : `${cssStyle} ${size}`;
}

// ── Floating (drawing anchor) rendering ──

// EMU conversions: 1px = 9525 EMU, 1pt = 12700 EMU.
const EMU_PER_PX = 9525;
const EMU_PER_PT = 12700;

interface FloatingLike {
  behindDocument?: boolean | null;
  zIndex?: number | null;
  wrap?: { type?: number | null; side?: string | null } | null;
  horizontalPosition?: {
    offset?: number | null;
    align?: string | null;
    relative?: string | null;
  } | null;
  verticalPosition?: {
    offset?: number | null;
    align?: string | null;
    relative?: string | null;
  } | null;
  margins?: {
    top?: number | null;
    bottom?: number | null;
    left?: number | null;
    right?: number | null;
  } | null;
}

/**
 * Render a drawing's Floating anchor (wp:anchor) to CSS — shared by Image and
 * WpgGroup so anchored drawings render consistently:
 *  - wrapNone (type 0): position:absolute at the EMU offset (text does not flow
 *    around it — "in front of text" / "behind text"); z-index lifts it above or
 *    below the text layer per behindDocument.
 *  - topAndBottom (3): float + clear:both.
 *  - square/tight/through (1/2/4): with an hPos.offset, float:right pulled to
 *    the offset via margin-right so text wraps beside it with no overlap;
 *    otherwise float on the wrap side (tight/through add shape-outside, src).
 * Wrap margins (wp:distT/B/L/R) are EMU, rendered as pt.
 */
export function floatingToStyles(floating: unknown, src?: string, width?: number): string[] {
  const f = floating as FloatingLike | null | undefined;
  if (!f) return [];
  const styles: string[] = [];
  styles.push(`z-index:${f.behindDocument ? -1 : (f.zIndex ?? 1)}`);

  const wrapType = f.wrap?.type ?? 0;
  const hOff = f.horizontalPosition?.offset;
  const vOff = f.verticalPosition?.offset;

  if (wrapType === 0) {
    // wrapNone: in front of/behind text — text does not flow around it. Pin to
    // the page box (.docen-page is position:relative) at the EMU offset, or at
    // the alignment when there is no offset: align center → left:50% +
    // translateX so a page/margin-anchored drawing centers instead of collapsing
    // to its inline origin. z-index lifts it above/below the text layer.
    styles.push("position:absolute");
    if (hOff != null) {
      styles.push(`left:${(hOff / EMU_PER_PX).toFixed(1)}px`);
    } else {
      const hAlign = f.horizontalPosition?.align;
      if (hAlign === "center") styles.push("left:50%", "transform:translateX(-50%)");
      else if (hAlign === "right") styles.push("right:0");
    }
    if (vOff != null) styles.push(`top:${(vOff / EMU_PER_PX).toFixed(1)}px`);
  } else if (wrapType !== 3 && hOff != null && width != null) {
    // square/tight/through with an explicit offset: the image must float so text
    // wraps around it — position:absolute would sit on top of the text. The
    // offset is the image's left edge; float:right + margin-right:
    // calc(100% - offset - width) pulls the float box left from the right edge
    // until its left edge lands on the offset, so the image keeps its hPos.offset
    // AND text wraps beside it with no overlap. vPos.offset → margin-top.
    const offPx = (hOff / EMU_PER_PX).toFixed(1);
    styles.push("float:right", `margin-right:calc(100% - ${offPx}px - ${width}px)`);
    if (vOff != null) styles.push(`margin-top:${(vOff / EMU_PER_PX).toFixed(1)}px`);
    // Wrap margins (wp:distL/R/T/B, EMU → pt): the minimum gap between the image
    // and the wrapping text (OOXML default 0.5"/457200 EMU when unset). Text
    // wraps on the left of a float:right image, so distL widens the float's
    // margin-box the text wraps against (text right edge = image left − distL)
    // — without it the text sits flush against the image edge. margin-left
    // extends the wrap box, not the border-box (still pinned to the offset by
    // margin-right), so the image keeps its position.
    const m = f.margins;
    if (m?.left) styles.push(`margin-left:${(m.left / EMU_PER_PT).toFixed(1)}pt`);
  } else if (wrapType === 3) {
    // topAndBottom: float + clear so text sits only above/below the image.
    styles.push("float:left", "clear:both");
  } else {
    // square/tight/through without an offset: float on the wrap side; tight/
    // through add shape-outside so text follows the image contour (needs src).
    const side = f.wrap?.side;
    const floatSide = side === "right" || side === "outside" ? "right" : "left";
    styles.push(`float:${floatSide}`);
    if ((wrapType === 2 || wrapType === 4) && src) {
      styles.push(`shape-outside:url(${src})`);
    }
  }

  // Wrap margins (wp:distT/B/L/R, EMU → pt) set the gap between the image and
  // the wrapping text — only for float modes whose position isn't already
  // encoded in margin-right (the offset-float above). wrapNone is absolute.
  const offsetFloat = wrapType !== 0 && wrapType !== 3 && hOff != null && width != null;
  if (!offsetFloat && wrapType !== 0) {
    const m = f.margins;
    if (m) {
      if (m.top) styles.push(`margin-top:${(m.top / EMU_PER_PT).toFixed(1)}pt`);
      if (m.bottom) styles.push(`margin-bottom:${(m.bottom / EMU_PER_PT).toFixed(1)}pt`);
      if (m.left) styles.push(`margin-left:${(m.left / EMU_PER_PT).toFixed(1)}pt`);
      if (m.right) styles.push(`margin-right:${(m.right / EMU_PER_PT).toFixed(1)}pt`);
    }
  }
  return styles;
}

/**
 * Where a wrapNone floating drawing's absolute top/left should resolve from.
 *
 * verticalPosition.relative (OOXML vRelativeFrom) defaults to "paragraph"
 * (also "line"); "page"/"margin"/"column" scope to the page box. The drawing is
 * position:absolute, so its offsetParent decides the origin: a
 * paragraph-anchored drawing must sit inside a position:relative <p> (the editor
 * CSS adds that via p:has([data-float-anchor])), otherwise top/left measure
 * from the page top and the drawing floats over the heading/body instead of
 * its own blank paragraph (verified on sample anchored drawing groups). Only
 * relevant for wrapNone (type 0); the float-based wraps stay inline.
 */
export function floatAnchorScope(floating: unknown): "paragraph" | "page" {
  const f = floating as FloatingLike | null | undefined;
  const vRel = f?.verticalPosition?.relative;
  if (vRel === "page" || vRel === "margin" || vRel === "column") return "page";
  return "paragraph";
}

// ── Floating table (w:tblpPr) rendering ──

/**
 * Render a table's float anchor (w:tblpPr → TableFloatOptions) to CSS.
 *
 * Unlike a drawing's Floating (wp:anchor, EMU offsets, wrap types), a floating
 * table carries no wrap type — Word's "text wrapping" around a table is plain
 * CSS float with margins. Twips (not EMU) throughout: tblpX/Y and the fromText
 * gaps are dxa.
 *
 * Two render modes:
 *  - text-anchored wrap (horizontalAnchor=text) → CSS float + margins, so body
 *    text flows beside the table. tblpX → margin on the float side; fromText →
 *    the opposite side + top/bottom (mirrors floatingToStyles so the offset and
 *    the wrap gap never compete for the same margin edge).
 *  - page/margin anchor → position:absolute pinned to the page box, floating at
 *    the offset/alignment like Word's page-anchored table. .docen-page is
 *    position:relative and its padding box is the physical page (no border), so
 *    top:0/left:0 matches OOXML's page origin; symmetric page padding makes
 *    left:50% the content-box center too, so a margin anchor centers alike.
 *  - text-anchored center/inside/outside have no CSS float equivalent (and
 *    inside/outside need odd/even pages the editor lacks) → degraded to [].
 *    attrs still round-trip byte-faithful via renderDocx/parseDocx passthrough.
 *  - overlap (neverOverlap) has no CSS float equivalent; ignored.
 */
export function tableFloatToCss(float: unknown): string[] {
  const f = float as TableFloatOptions | null | undefined;
  if (!f) return [];

  const hAnchorPage = f.horizontalAnchor === "page" || f.horizontalAnchor === "margin";
  const vAnchorPage = f.verticalAnchor === "page" || f.verticalAnchor === "margin";

  // page/margin anchor → position:absolute within the page box (see JSDoc). The
  // table detaches from the text flow and floats at the offset/alignment.
  if (hAnchorPage || vAnchorPage) {
    const styles: string[] = ["position:absolute"];
    if (vAnchorPage && f.absoluteVerticalPosition != null) {
      const top = twipToCss(f.absoluteVerticalPosition);
      if (top) styles.push(`top:${top}`);
    }
    if (hAnchorPage) {
      const side = f.relativeHorizontalPosition;
      if (f.absoluteHorizontalPosition != null) {
        const left = twipToCss(f.absoluteHorizontalPosition);
        if (left) styles.push(`left:${left}`);
      } else if (side === "center") {
        styles.push("left:50%", "transform:translateX(-50%)");
      } else if (side === "right") {
        styles.push("right:0");
      } else {
        styles.push("left:0"); // left/inside/outside → best-effort left edge
      }
    }
    return styles;
  }

  // text-anchored wrap — center/inside/outside have no CSS float equivalent
  // (inside/outside also need odd/even pages the editor lacks).
  const side = f.relativeHorizontalPosition;
  if (side === "center" || side === "inside" || side === "outside") return [];

  const floatRight = side === "right";
  const styles: string[] = [floatRight ? "float:right" : "float:left"];

  // Absolute horizontal offset (tblpX, twips) — on the float side. Usually only
  // set for left floats (right floats use tblpXSpec=right with no tblpX).
  if (f.absoluteHorizontalPosition != null) {
    const off = twipToCss(f.absoluteHorizontalPosition);
    if (off) styles.push(floatRight ? `margin-right:${off}` : `margin-left:${off}`);
  }

  // fromText gaps (twips) — opposite the float side + top/bottom, so they never
  // clash with the offset margin above.
  const gapSide = floatRight ? f.leftFromText : f.rightFromText;
  if (gapSide != null) {
    const m = twipToCss(gapSide);
    if (m) styles.push(floatRight ? `margin-left:${m}` : `margin-right:${m}`);
  }
  if (f.topFromText != null) {
    const m = twipToCss(f.topFromText);
    if (m) styles.push(`margin-top:${m}`);
  }
  if (f.bottomFromText != null) {
    const m = twipToCss(f.bottomFromText);
    if (m) styles.push(`margin-bottom:${m}`);
  }

  return styles;
}

// ── Style rendering (consume nested office-open attrs) ──

interface ParagraphStyleShape {
  alignment?: string | null;
  indent?: IndentAttributesProperties | null;
  spacing?: SpacingProperties | null;
  shading?: ShadingAttributesProperties | null;
  border?: BordersOptions | null;
  /** snapToGrid (w:snapToGrid): preserved for OOXML round-trip fidelity. Word
   *  snaps the baseline grid to linePitch (governs lines-per-page) but does NOT
   *  add linePitch to rendered line height, so this flag no longer affects
   *  lineSpacingToCss. Defaults to true (omitted = use grid); header/footer
   *  styles set it false. */
  snapToGrid?: boolean | null;
  /** Paragraph-mark (¶) run properties (pPr/rPr). Per OOXML (ECMA-376) these
   *  format the ¶ glyph only — never applied to run text (a large ¶ marker
   *  must not inflate body runs). Only `size` is rendered, as the
   *  paragraph's line-height: the ¶ glyph is a physical character whose
   *  font-size sets the paragraph's (esp. empty) line height in Word. */
  run?: { size?: number | null } | null;
}

/**
 * Compute all paragraph-level CSS styles from nested attrs.
 * Shared by Paragraph and Heading extensions for node-level renderHTML.
 * Attrs store office-open native values; mappers here convert to CSS.
 */
export function renderParagraphStyles(
  attrs: Record<string, unknown>,
  opts?: { empty?: boolean },
): string[] {
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

  // Paragraph-mark (¶) glyph font-size → line-height. Per OOXML (ECMA-376) the
  // ¶ glyph is a physical character whose font-size sets the paragraph's line
  // height — this is why an empty paragraph still occupies a line in Word (its
  // sole content is the ¶ glyph). Only `size` is rendered: the ¶ glyph is
  // invisible, so font/color/bold have no visual effect and must NOT become the
  // paragraph's font-size (that would leak onto every run).
  // Placed BEFORE spacing so an explicit spacing/line rule overrides it (Word:
  // an explicit line rule wins over the ¶-glyph single-line height).
  const markLineHeight = a.run?.size != null ? sizeToCss(a.run.size) : null;
  if (markLineHeight) styles.push(`line-height:${markLineHeight}`);

  if (a.spacing) {
    const before = twipToCss(a.spacing.before);
    if (before) styles.push(`margin-top:${before}`);
    const after = twipToCss(a.spacing.after);
    if (after) styles.push(`margin-bottom:${after}`);
    // An empty paragraph (¶ glyph only) does NOT receive the document-grid
    // pitch — Word renders the ¶ at the font's natural metric (verified vs a
    // Word PDF: an empty single line between paragraphs measures ~natural, not
    // natural+pitch). snapToGrid=false applies the spacing multiple against the
    // natural metric, mirroring measure.ts emptyLineHeight (edit == render).
    const lh = lineSpacingToCss(a.spacing, opts?.empty ? false : a.snapToGrid);
    if (lh) styles.push(`line-height:${lh}`);
  }

  const bg = shadingToCss(a.shading);
  // A fill flips the default ink against it (Word "auto"): declare color on
  // the fill's own element so descendants inherit the inverted value here,
  // not a value pre-computed at the page root (var resolves at the declaring
  // element, so a page-level contrast-color(var) wouldn't follow this fill).
  if (bg) styles.push(`background-color:${bg}`, `color:contrast-color(${bg})`);

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

  // (pPr/rPr run props other than size — font/color/bold — are NOT rendered:
  // the ¶ glyph is invisible, and emitting them would leak onto the paragraph's
  // runs. Only size is emitted, as line-height above. attrs.run is still
  // carried verbatim for lossless DOCX round-trip; renderDocx emits opts.run.)

  return styles;
}

interface RunStyleShape {
  bold?: boolean | null;
  italic?: boolean | null;
  underline?: unknown;
  strike?: boolean | null;
  doubleStrike?: boolean | null;
  color?: unknown;
  size?: number | null;
  font?: unknown;
  smallCaps?: boolean | null;
  allCaps?: boolean | null;
  characterSpacing?: number | null;
  highlight?: unknown;
}

/**
 * Compute run-level CSS (font/size/color/weight/…) from office-open run attrs.
 * Shared by text-style marks and the styles→CSS generator (stylesToCss).
 */
export function renderRunStyles(attrs: Record<string, unknown>): string[] {
  const a = attrs as RunStyleShape;
  const styles: string[] = [];

  if (a.bold) styles.push("font-weight:bold");
  if (a.italic) styles.push("font-style:italic");
  if (a.smallCaps) styles.push("font-variant:small-caps");
  if (a.allCaps) styles.push("text-transform:uppercase");

  const deco: string[] = [];
  if (a.underline) deco.push("underline");
  if (a.strike || a.doubleStrike) deco.push("line-through");
  if (deco.length) styles.push(`text-decoration:${deco.join(" ")}`);

  const font = resolveFontFamilyCss(a.font);
  if (font) styles.push(`font-family:${font}`);
  const size = sizeToCss(a.size);
  if (size) styles.push(`font-size:${size}`);
  // OOXML "auto"/unset = a readable ink against the surrounding fill. Both
  // emit no inline color so the text inherits the page default
  // (color: contrast-color(var(--docen-ink-bg)) on .docen-page), flipping to
  // white on dark fills like Word. An explicit hex overrides it.
  const color = a.color === "auto" ? undefined : normalizeColorToHex(a.color);
  if (color) styles.push(`color:${color}`);
  const spacing = characterSpacingToCss(a.characterSpacing);
  if (spacing) styles.push(`letter-spacing:${spacing}`);
  if (a.highlight) {
    const hl = normalizeColorToHex(typeof a.highlight === "string" ? a.highlight : null);
    // A highlight is a run-local background, so it also becomes the ink-bg.
    // A highlight is a run-local background, so it also flips the ink.
    if (hl) styles.push(`background-color:${hl}`, `color:contrast-color(${hl})`);
  }

  return styles;
}

interface CellStyleShape {
  noWrap?: boolean | null;
  shading?: ShadingAttributesProperties | null;
  verticalAlign?: string | null;
  borders?: BordersOptions | null;
  // w:tcMar (or the inherited w:tblCellMar pushed onto the cell by resolveTable)
  // — office-open TableCellMarginOptions: top/left/bottom/right each a
  // TableWidthProperties { size (twips), type }.
  margins?: TableCellOptions["margins"];
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
  // A fill flips the default ink against it (Word "auto") for every run in
  // the cell — descendants inherit the inverted color declared here.
  if (bg) styles.push(`background-color:${bg}`, `color:contrast-color(${bg})`);

  if (a.verticalAlign) styles.push(`vertical-align:${a.verticalAlign}`);

  // Cell insets (w:tcMar; resolveTable pushes the table's tblCellMar default
  // onto cells that lack their own tcMar): each side → padding (twips→pt).
  // Without this a row renders at its trHeight floor with zero insets — ~8pt
  // shorter than Word, which reserves tblCellMar top/bottom around every cell.
  if (a.margins) {
    const m = a.margins;
    const side = (s: TableWidthProperties | null | undefined): string =>
      s && typeof s.size === "number" ? (twipToCss(s.size) ?? "0pt") : "0pt";
    styles.push(`padding:${side(m.top)} ${side(m.right)} ${side(m.bottom)} ${side(m.left)}`);
  }

  // Cell borders (w:tcBorders): each present side → border-<side>. A
  // "nil"/"none" side emits "none" so it overrides the default Table-Grid
  // border (Word leaves those cell edges open via w:val="nil").
  if (a.borders) {
    const sides = ["top", "right", "bottom", "left"] as const;
    for (const side of sides) {
      const b = (a.borders as Record<string, BorderOptions | undefined>)[side];
      if (!b) continue;
      styles.push(`border-${side}:${renderBorderCSS(b) ?? "none"}`);
    }
  }

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
