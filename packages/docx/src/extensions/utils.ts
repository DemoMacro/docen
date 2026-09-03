import { sectionPageSizeDefaults } from "@office-open/docx";
import type {
  BorderOptions,
  BordersOptions,
  IndentProperties,
  ParagraphPropertiesOptionsBase,
  ShadingProperties,
  SpacingProperties,
  TableCellOptions,
} from "@office-open/docx";

// ── Tiptap attr factory ──

/** Factory for a Tiptap attr that carries an office-open native value: never
 *  parsed from HTML nor rendered to it, defaulting to null (ProseMirror stores
 *  every declared attr). Shared by every extension carrying OOXML attrs
 *  (paragraph/table/table-cell/…). */
export const attrNative = () => ({ default: null, parseHTML: () => null, rendered: false });

// ── Shared paragraph attr factory ──
//
// In OOXML a heading IS a paragraph — a <w:p> carrying <w:pStyle val="Heading1">
// in its pPr; sectPr (when present) lives in that same <w:p>'s pPr. office-open
// models this faithfully: `heading` is a field on ParagraphPropertiesOptionsBase,
// not a separate type (see @office-open/docx parts/paragraph/properties.ts).
// The single `paragraph` node mirrors that model directly: heading/style/bullet/
// numbering/thematicBreak are attrs like every other paragraph property, and
// consumers derive the heading level via detectHeadingLevel (paragraph.ts).

/** Editor-only attrs marking a textblock as a section's last paragraph (its pPr
 *  holds the OOXML sectPr). DocxManager peels them off to close a section in
 *  compile; they must NOT leak into ParagraphOptions. Typed as plain strings
 *  for .has(string) call sites; the keys stay pinned to the attr table via
 *  ParagraphAttrKey below. */
const SECTION_CLOSE_KEYS = ["sectionProperties", "sectionHeaders", "sectionFooters"] as const;
export const SECTION_ATTR_KEYS = new Set<string>(SECTION_CLOSE_KEYS);

// ── Paragraph attr mirror contract ──
//
// The attr table is a hand-written mirror of ParagraphPropertiesOptionsBase.
// Types keep the mirror from drifting: `satisfies Record<ParagraphAttrKey, …>`
// requires EVERY office-open paragraph property to have an attr (a new
// office-open field without one fails the build) and rejects attr keys that
// exist nowhere in office-open (typos, retired fields).

/** Attr keys layered on top of the office-open mirror: the section-close
 *  markers (SECTION_ATTR_KEYS). Every other office-open paragraph property —
 *  including heading/style/bullet/numbering/thematicBreak, once owned by the
 *  deleted heading/list nodes — is mirrored verbatim. */
type EditorParagraphAttrKey = (typeof SECTION_CLOSE_KEYS)[number];

/** The full attr key set the paragraph node declares. */
type ParagraphAttrKey = EditorParagraphAttrKey | keyof ParagraphPropertiesOptionsBase;

/** The attr spec shape every mirror table uses (attrNative plus HTML-paste
 *  parseHTML keys). Shared by the paragraph/table-family/run satisfies guards. */
export type DocxAttrSpec = {
  default: unknown;
  parseHTML?: (el: HTMLElement) => unknown;
  rendered?: boolean;
};

/** Shared office-open paragraph attrs — ParagraphPropertiesOptionsBase mirror.
 *  The paragraph node declares the whole table (a heading is a paragraph in
 *  OOXML — its HeadingLevel pStyle rides on the `heading` attr). The satisfies
 *  guard pins the key set to ParagraphAttrKey — see the mirror contract above. */
export function docxParagraphAttrs() {
  return {
    // pStyle reference (e.g. "Heading1", "Normal"). parseHTML reads it back
    // from class="docx-style-{style}" on pasted HTML.
    style: {
      default: null,
      parseHTML: (el: HTMLElement) => {
        const m = (el.getAttribute("class") || "").match(/(?:^|\s)docx-style-(\S+)/);
        return m ? m[1] : null;
      },
    },
    // HeadingLevel pStyle literal ("Heading1".."Heading9"/"Title") — what
    // office-open lifts a Heading pStyle into. `style` wins on stringify
    // (office-open: style ?? heading), matching Word's single-pStyle rule.
    heading: attrNative(),
    // Bullet list marker {level} — the flat list model (R2).
    bullet: attrNative(),
    // Numbered list reference {reference, instance, level} — the flat list
    // model (R2).
    numbering: attrNative(),
    // <w:thematicBreak/> marker (a paragraph reduced to a horizontal rule).
    thematicBreak: attrNative(),
    // Nested office-open objects (parsed from HTML where CSS exists).
    alignment: {
      default: null,
      rendered: false,
      parseHTML: (el: HTMLElement) => alignmentFromElement(el),
    },
    indent: {
      default: null,
      rendered: false,
      parseHTML: (el: HTMLElement) => indentFromElement(el),
    },
    spacing: {
      default: null,
      rendered: false,
      parseHTML: (el: HTMLElement) => spacingFromElement(el),
    },
    shading: {
      default: null,
      rendered: false,
      parseHTML: (el: HTMLElement) => shadingFromElement(el),
    },
    border: {
      default: null,
      rendered: false,
      parseHTML: (el: HTMLElement) => bordersFromElement(el),
    },
    frame: attrNative(),
    // Paragraph-mark (¶) run properties (pPr/rPr): format ONLY the ¶ glyph.
    run: attrNative(),
    // Section properties carried on a section's LAST paragraph (OOXML sectPr
    // lives in that paragraph's pPr). A heading can be that paragraph.
    sectionProperties: attrNative(),
    sectionHeaders: attrNative(),
    sectionFooters: attrNative(),
    // Scalar OOXML paragraph properties (stored verbatim; no CSS equivalent).
    keepNext: attrNative(),
    keepLines: attrNative(),
    pageBreakBefore: attrNative(),
    widowControl: attrNative(),
    contextualSpacing: attrNative(),
    bidirectional: attrNative(),
    outlineLevel: attrNative(),
    textDirection: attrNative(),
    textAlignment: attrNative(),
    suppressLineNumbers: attrNative(),
    wordWrap: attrNative(),
    overflowPunctuation: attrNative(),
    autoSpaceEastAsianText: attrNative(),
    suppressOverlap: attrNative(),
    suppressAutoHyphens: attrNative(),
    adjustRightInd: attrNative(),
    snapToGrid: attrNative(),
    mirrorIndents: attrNative(),
    kinsoku: attrNative(),
    topLinePunct: attrNative(),
    autoSpaceDE: attrNative(),
    textboxTightWrap: attrNative(),
    rightTabStop: attrNative(),
    leftTabStop: attrNative(),
    divId: attrNative(),
    tabStops: attrNative(),
    cnfStyle: attrNative(),
    // Round-trip marker for a bare <w:pPr/> (element presence is fidelity).
    emptyProperties: attrNative(),
    // Mirror contract: every office-open paragraph property + editor key
    // declared, nothing else (see ParagraphAttrKey).
  } satisfies Record<ParagraphAttrKey, DocxAttrSpec>;
}

/** TableCellOptions keys owned elsewhere: `rowSpan` is office-open's
 *  authoring sugar (verticalMerge is the round-tripped truth), `children` is
 *  rebuilt by DocxManager, and `text`/`cellProperties` are sdt-cell-only
 *  markers — not mirrored as office-open-native attrs. */
type CellKeyElsewhere = "rowSpan" | "children" | "text" | "cellProperties";

/** The full attr key set the tableCell node declares. */
type TableCellAttrKey = Exclude<keyof TableCellOptions, CellKeyElsewhere>;

/** Shared office-open table-cell attrs — TableCellPropertiesOptions mirror.
 *  The satisfies guard pins the key set to TableCellAttrKey — same mirror
 *  contract as docxParagraphAttrs. */
export function docxTableCellAttrs() {
  return {
    // Nested office-open objects (parsed from HTML where CSS exists)
    shading: {
      default: null,
      rendered: false,
      parseHTML: (el: HTMLElement) => shadingFromElement(el),
    },
    borders: {
      default: null,
      rendered: false,
      parseHTML: (el: HTMLElement) => bordersFromElement(el),
    },
    verticalAlign: {
      default: null,
      rendered: false,
      parseHTML: (el: HTMLElement) => el.style.verticalAlign || null,
    },

    // Scalar OOXML cell properties (stored verbatim; no CSS equivalent)
    textDirection: attrNative(),
    width: attrNative(),
    columnSpan: attrNative(),
    margins: attrNative(),
    noWrap: {
      default: null,
      rendered: false,
      parseHTML: (el: HTMLElement) => (el.style.whiteSpace === "nowrap" ? true : null),
    },
    verticalMerge: attrNative(),
    horizontalMerge: attrNative(),
    fitText: attrNative(),
    hideMark: attrNative(),
    headers: attrNative(),
    cnfStyle: attrNative(),
    // Row/cell-level track-change markers (w:ins/w:del on the tcPr/trPr side).
    insertion: attrNative(),
    deletion: attrNative(),
    cellMerge: attrNative(),
    revision: attrNative(),
    // Mirror contract: every office-open cell property declared, nothing else.
  } satisfies Record<TableCellAttrKey, DocxAttrSpec>;
}

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
  // element.style serializes colors as rgb()/rgba() — the form pasted HTML
  // actually arrives in. Fully transparent carries no visible color.
  const rgb = /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*(?:,\s*([\d.]+)\s*)?\)$/.exec(
    color,
  );
  if (rgb) {
    if (rgb[4] != null && Number(rgb[4]) === 0) return undefined;
    const byte = (n: string) => Number(n).toString(16).padStart(2, "0").toUpperCase();
    return `#${byte(rgb[1])}${byte(rgb[2])}${byte(rgb[3])}`;
  }
  const hex = CSS_COLORS[color.toLowerCase()];
  return hex ?? undefined;
}

// Office default theme1.xml font tokens → concrete font names. office-open does
// not parse the theme part, so a run's rFonts *Theme attributes survive as the
// literal tokens "minorHAnsi"/"minorEastAsia"/…; map them to the Office default
// theme fonts here (Calibri/等线 on a stock zh-CN template). Documents carrying
// a custom theme fall back to these defaults — rare, and still better than the
// browser's Segoe UI fallback. The CJK eastAsia font varies by the document's
// eastAsia language; zh-CN (等线) is the default, covering the common case.
const THEME_LATIN_FONTS: Record<string, string> = {
  minorHAnsi: "Calibri",
  majorHAnsi: "Calibri Light",
  minorBidi: "Arial",
  majorBidi: "Arial",
};
const THEME_EAST_ASIA_FONTS: Record<string, Record<string, string>> = {
  "zh-CN": { minorEastAsia: "等线", majorEastAsia: "等线 Light" },
  "zh-TW": { minorEastAsia: "新細明體", majorEastAsia: "微軟正黑體" },
  "ja-JP": { minorEastAsia: "游ゴシック", majorEastAsia: "游ゴシック Light" },
  "ko-KR": { minorEastAsia: "맑은 고딕", majorEastAsia: "맑은 고딕" },
};

/** Resolve a run's rFonts — literal ascii/eastAsia/hAnsi OR *Theme tokens — to
 *  the concrete { ascii, eastAsia } pair used by both font-family renderers.
 *  *Theme tokens map to the Office default theme fonts (office-open leaves them
 *  unresolved). `eastAsiaLang` picks the CJK font for *eastAsiaTheme (zh-CN
 *  default). Returns null only when the value is empty/non-font. */
export function resolveRFonts(
  font: unknown,
  eastAsiaLang = "zh-CN",
): { ascii: string | null; eastAsia: string | null } | null {
  if (!font) return null;
  if (typeof font === "string") return { ascii: font, eastAsia: null };
  if (typeof font !== "object") return null;
  const f = font as {
    ascii?: string;
    hAnsi?: string;
    eastAsia?: string;
    asciiTheme?: string;
    hAnsiTheme?: string;
    eastAsiaTheme?: string;
  };
  const eaMap = THEME_EAST_ASIA_FONTS[eastAsiaLang] ?? THEME_EAST_ASIA_FONTS["zh-CN"];
  const ascii =
    f.ascii ?? THEME_LATIN_FONTS[f.asciiTheme ?? ""] ?? THEME_LATIN_FONTS[f.hAnsiTheme ?? ""];
  const hAnsi = f.hAnsi ?? THEME_LATIN_FONTS[f.hAnsiTheme ?? ""] ?? ascii;
  const eastAsia = f.eastAsia ?? eaMap[f.eastAsiaTheme ?? ""];
  return { ascii: ascii ?? hAnsi ?? null, eastAsia: eastAsia ?? null };
}

/** Resolve a font value (string or OOXML rFonts, incl. *Theme tokens) to a
 *  single CSS family name (ascii/hAnsi/eastAsia). */
export function resolveFontName(font: unknown, eastAsiaLang = "zh-CN"): string | null {
  const r = resolveRFonts(font, eastAsiaLang);
  return r?.ascii ?? r?.eastAsia ?? null;
}

// ── Unit conversion helpers ──
// office-open stores native values (twips, points); CSS values convert here.

/** CSS value (e.g., "18pt") → twip number. 1 pt = 20 twips, 1 px = 15 twips (96 DPI). */
export function cssToTwip(value: string | undefined): number | undefined {
  if (!value) return undefined;
  const match = value.match(/^([\d.]+)(pt|px|em|cm|in)?$/);
  if (!match) return undefined;
  const num = parseFloat(match[1]);
  const unit = match[2] ?? "pt";
  // The regex leaves exactly these units; pt (or no unit) is the fallback.
  if (unit === "px") return Math.round(num * 15);
  if (unit === "in") return Math.round(num * 1440);
  if (unit === "cm") return Math.round(num * 567);
  if (unit === "em") return Math.round(num * 240);
  return Math.round(num * 20);
}

// ── Alignment mapping ──
// OOXML AlignmentType has no "justify" — both-sides value is "both".

const CSS_TO_ALIGNMENT: Record<string, string> = {
  left: "left",
  center: "center",
  right: "right",
  start: "start",
  end: "end",
  justify: "both",
};

/** CSS text-align → OOXML alignment. */
export function alignmentFromCss(css: string | null | undefined): string | null {
  if (!css) return null;
  return CSS_TO_ALIGNMENT[css] ?? null;
}

// ── Shading mapping ──

/** CSS background-color → ShadingProperties (fill normalized to hex). */
export function shadingFromCss(css: string | null | undefined): ShadingProperties | null {
  const hex = normalizeColorToHex(css ?? undefined);
  // w:fill is ST_HexColor — bare six-digit hex, no "#" prefix.
  return hex ? { fill: hex.replace(/^#/, ""), type: "clear" } : null;
}

// ── Section geometry ──
// OOXML section properties (CT_SectPr): page size/margin + document grid.

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

// ── Font size mapping ──
// office-open size is in POINTS (new convention).

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

export function characterSpacingFromCss(css: string | null | undefined): number | null {
  if (!css) return null;
  const m = css.match(/^(-?[\d.]+)pt$/);
  return m ? Math.round(parseFloat(m[1]) * 20) : null;
}

// ── Element parsers (CSS → office-open native, for parseHTML) ──
// Shared by Paragraph and Heading: each attr's parseHTML calls one of these.

/** Parse text-align → OOXML alignment. */
export function alignmentFromElement(el: HTMLElement): string | null {
  return alignmentFromCss(el.style.textAlign || null);
}

/** Parse margin-left/right + text-indent → OOXML indent (twips). */
export function indentFromElement(el: HTMLElement): IndentProperties | null {
  const indent: IndentProperties = {};
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

/** Parse border-* → OOXML BordersOptions. Chrome serializes the shorthand as
 *  "width style color" ("1px solid rgb(0, 0, 0)") — the form real pastes
 *  arrive in; hand-written HTML often lists "style width color" (CSS lets the
 *  components appear in any order), so both orders parse. The width accepts
 *  px and pt, and the color rides the same hex normalization as shading,
 *  since a raw rgb() string in w:color violates ST_HexColor. */
export function bordersFromElement(el: HTMLElement): BordersOptions | null {
  const borders: BordersOptions = {};
  const styleMap: Record<string, BorderOptions["style"]> = {
    solid: "single",
    dashed: "dashed",
    dotted: "dotted",
    double: "double",
  };
  const sides: Array<[keyof BordersOptions, string]> = [
    ["top", el.style.borderTop],
    ["bottom", el.style.borderBottom],
    ["left", el.style.borderLeft],
    ["right", el.style.borderRight],
  ];
  for (const [side, css] of sides) {
    if (!css || css === "initial" || css === "none") continue;
    const m =
      /^([\d.]+)(pt|px)\s+(solid|dashed|dotted|double)\s+(.+)$/.exec(css) ??
      /^(solid|dashed|dotted|double)\s+([\d.]+)(pt|px)\s+(.+)$/.exec(css);
    if (!m) continue;
    // The second form lists style first — normalize to [width, unit, style].
    const [width, unit, style] = Number.isNaN(parseFloat(m[1]!))
      ? [m[2], m[3], m[1]]
      : [m[1], m[2], m[3]];
    const color = normalizeColorToHex(m[4]);
    if (!color) continue;
    borders[side] = {
      style: styleMap[style] ?? "single",
      // Eighth-points (w:sz): 1pt = 8, 1px = 6 at 96 DPI.
      size: Math.round(parseFloat(width) * (unit === "px" ? 6 : 8)),
      color: color.replace(/^#/, ""),
    };
  }
  return Object.keys(borders).length > 0 ? borders : null;
}

/** Parse background-color → OOXML shading. */
export function shadingFromElement(el: HTMLElement): ShadingProperties | null {
  return shadingFromCss(el.style.backgroundColor || null);
}

// ── Clipboard slice round-trip ──

/** The custom MIME type carrying a docen selection as ProseMirror slice JSON —
 *  the lossless clipboard lane between docen editors (the plain-text lane is
 *  for every other consumer). */
export const DOCEN_CLIP_MIME = "application/x-docen-docx";

/** Serialize the selection as a PM slice JSON payload (marks, node attrs, and
 *  open depths intact) — null for an empty selection. */
export function selectionSlicePayload(state: {
  selection: { from: number; to: number };
  doc: {
    slice(
      from: number,
      to: number,
      leafNodes?: boolean,
    ): { openStart: number; openEnd: number; content: { toJSON(): unknown } };
  };
}): string | null {
  const { from, to } = state.selection;
  if (from === to) return null;
  const slice = state.doc.slice(from, to, true);
  return JSON.stringify({
    openStart: slice.openStart,
    openEnd: slice.openEnd,
    content: slice.content.toJSON(),
  });
}
