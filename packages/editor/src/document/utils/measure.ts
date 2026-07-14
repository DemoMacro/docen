import {
  prepareRichInline,
  layoutNextRichInlineLineRange,
  materializeRichInlineLineRange,
  type RichInlineItem,
  type RichInlineCursor,
} from "@chenglou/pretext/rich-inline";
import {
  defaultParagraphStyleId,
  indexParagraphStyles,
  mergeStyleChain,
  type StylesOptions,
} from "@docen/docx";
import type { SpacingProperties, TableWidthProperties } from "@office-open/docx";
import type { Node as PmNode } from "@tiptap/pm/model";

import { clearFontMetricCache, fontNormalRatio } from "./font-metric";

/**
 * Deterministic block measurement via Pretext (canvas measureText + pure
 * arithmetic), replacing DOM `getBoundingClientRect`/`offsetHeight` for the
 * paginator.
 *
 * Why not DOM: a block's rendered height wobbles sub-pixel between re-flow
 * passes — CJK `line-height:normal` resolves slightly differently across layout
 * passes (font hinting), so a block on the page-height boundary flips pages
 * each pass and page-count oscillates. Pretext measures with the browser's own
 * font engine (canvas measureText) but outside the layout process: same input
 * → same output, every pass. The remaining risk is *drift* between Pretext's
 * measurement and the browser's actual layout, controlled by matching
 * font/lineHeight/letterSpacing inputs to the CSS the renderer emits and by
 * clearing the cache once `document.fonts.ready` settles (see callers).
 *
 * See CLAUDE.md → Pagination Architecture and CONTRIBUTING.md → Pagination
 * Conventions.
 */

type PreparedRichInline = ReturnType<typeof prepareRichInline>;

// prepare() is the expensive one-time pass (segment + canvas-measure each run);
// measureRichInlineStats() is pure arithmetic over cached widths. Cache the
// prepared handle by (text + font + letterSpacing) so re-flows and resizes only
// pay the cheap path.
const preparedCache = new Map<string, PreparedRichInline>();
const PREPARED_CACHE_LIMIT = 4000;

/** Drop the prepare cache (call after `document.fonts.ready` — fonts loaded
 *  after caching change canvas widths, so old measurements drift). */
export function clearMeasureCache(): void {
  preparedCache.clear();
  clearFontMetricCache();
}

function getPrepared(items: RichInlineItem[]): PreparedRichInline {
  const key = items
    .map((it) => `${it.text}\u0000${it.font}\u0000${it.letterSpacing ?? ""}`)
    .join("");
  const cached = preparedCache.get(key);
  if (cached) return cached;
  const prepared = prepareRichInline(items);
  if (preparedCache.size >= PREPARED_CACHE_LIMIT) preparedCache.clear();
  preparedCache.set(key, prepared);
  return prepared;
}

// ── unit constants ──
// 1in = 25.4mm = 72pt = 96px → 1pt = 4/3 px. OOXML twip = 1/20 pt.
const PT_TO_PX = 4 / 3;
const TWIP_TO_PX = 4 / 3 / 20;

const DEFAULT_SIZE_PT = 12;
const DEFAULT_FAMILY = "serif";

// ── run-style extraction from node attrs + marks ──
// Field names mirror `packages/docx/src/extensions/utils.ts` `renderRunStyles`
// (font/size/bold/italic/characterSpacing) so the measured font matches the CSS
// the renderer emits.

interface RunStyle {
  size: number | null; // points
  font: unknown; // string | { ascii, eastAsia, hAnsi } | null
  bold: boolean;
  italic: boolean;
  characterSpacing: number | null; // twips
}

/** OOXML font (string or rFonts {ascii/hAnsi/eastAsia}) → a CSS family name. */
function resolveFamily(font: unknown): string | null {
  if (!font) return null;
  if (typeof font === "string") return font;
  if (typeof font === "object") {
    const f = font as { ascii?: string; hAnsi?: string; eastAsia?: string };
    return f.ascii || f.hAnsi || f.eastAsia || null;
  }
  return null;
}

type MarkLike = { type: { name: string }; attrs: Record<string, unknown> };

function markStyleOf(marks: readonly MarkLike[] | undefined): Partial<RunStyle> {
  const out: Partial<RunStyle> = {};
  if (!marks) return out;
  for (const m of marks) {
    if (m.type.name === "textStyle") {
      const a = m.attrs;
      if (a.size != null) out.size = a.size as number;
      if (a.font != null) out.font = a.font;
      if (a.characterSpacing != null) out.characterSpacing = a.characterSpacing as number;
    } else if (m.type.name === "bold") {
      out.bold = true;
    } else if (m.type.name === "italic") {
      out.italic = true;
    }
  }
  return out;
}

/** Document style table shape (doc.attrs.styles). Measure must mirror the
 *  renderer's cascade: a paragraph whose direct attrs are empty still renders
 *  with its style's spacing/font (e.g. a "Table" style's line=360, or a doc
 *  default of 宋体), so reading only `attrs` under-measures it. */
interface StyleTable {
  paragraphStyles?: Array<{
    id?: string;
    default?: boolean;
    basedOn?: string;
    paragraph?: {
      spacing?: SpacingProperties | null;
      indent?: IndentAttrs | null;
    };
    run?: { font?: unknown; size?: number | null; bold?: boolean; italic?: boolean };
  }>;
  default?: {
    document?: {
      paragraph?: {
        spacing?: SpacingProperties | null;
        indent?: IndentAttrs | null;
      };
      run?: { font?: unknown; size?: number | null; bold?: boolean; italic?: boolean };
    };
  };
}

function styleTableOf(styles: unknown): StyleTable | null {
  return styles && typeof styles === "object" ? (styles as StyleTable) : null;
}

/** The merged {run, paragraph} for a paragraph's style — reusing the RENDERER's
 *  `mergeStyleChain` (deep-merge the style's `basedOn` chain, root first) so
 *  pagination measures the SAME effective properties stylesToCss emits. A
 *  paragraph with NO explicit styleId resolves to the doc's default paragraph
 *  style (`w:default="1", usually Normal) — the same target as the renderer's
 *  `.docx-default` class. docDefaults is NOT in this chain (the renderer emits
 *  it as a separate base rule on p/h1-6); each resolver falls back to it last.
 *
 *  Cached per (styles, styleId): a document's styles model is stable for its
 *  lifetime, and resolveSpacing + resolveIndentAttrs + defaultRunOf all need
 *  the chain for ONE paragraph — without the cache each re-flow re-walked the
 *  basedOn chain 3× per paragraph. */
const NO_STYLE_KEY = "<no-style>";
const styleChainCache = new WeakMap<
  object,
  Map<string, { run: Record<string, unknown>; paragraph: Record<string, unknown> }>
>();
function styleChainOf(
  styles: unknown,
  styleId: string | null | undefined,
): { run: Record<string, unknown>; paragraph: Record<string, unknown> } | null {
  if (!styles || typeof styles !== "object") return null;
  const key = styleId || NO_STYLE_KEY;
  let perStyles = styleChainCache.get(styles as object);
  if (!perStyles) {
    perStyles = new Map();
    styleChainCache.set(styles as object, perStyles);
  }
  const cached = perStyles.get(key);
  if (cached) return cached;
  const byId = indexParagraphStyles(styles as StylesOptions);
  const id = styleId || defaultParagraphStyleId(styles as StylesOptions);
  const chain = id ? mergeStyleChain(byId, id) : { run: {}, paragraph: {} };
  perStyles.set(key, chain);
  return chain;
}

/** A paragraph's effective spacing: direct attr, else its style's, else the
 *  document default. Without the style/default fallback, paragraphs whose
 *  line-spacing lives in their style (e.g. line=360 on a "Table" style → 1.5×
 *  grid pitch = 31.2px) measured at the bare grid pitch (20.8px), so table rows
 *  measured ~1.5× too short and the paginator never split them. */
export function resolveSpacing(node: PmNode, styles: unknown): SpacingProperties | null {
  const direct = (node.attrs as { spacing?: SpacingProperties | null }).spacing;
  if (direct && direct.line != null) return direct;
  // Style chain via the renderer's mergeStyleChain (direct style + its basedOn
  // ancestors, e.g. Heading1 → Normal) — same source as stylesToCss, so measure
  // == render. A paragraph with no styleId resolves to the doc's default style
  // (Normal). Previously this read only the single direct style, missing a
  // basedOn ancestor's spacing.line (e.g. a heading whose line spacing lives on
  // Normal) → measured too tall/short vs the rendered page.
  const styleId = (node.attrs as { styleId?: string | null }).styleId;
  const sp = styleChainOf(styles, styleId)?.paragraph?.spacing as SpacingProperties | undefined;
  if (sp && sp.line != null) return sp;
  // docDefaults is the base layer under every named style (the renderer emits it
  // on p/h1-6); mergeStyleChain covers only the named-style chain, so fall back
  // to docDefaults last (matches stylesToCss's base-rule layer order).
  const t = styleTableOf(styles);
  const docSp = t?.default?.document?.paragraph?.spacing;
  if (docSp && docSp.line != null) return docSp;
  return null;
}

/** A paragraph's spacing.before/after in px (OOXML twips → px), resolved through
 *  the same direct → style → document-default cascade as line spacing. The
 *  renderer emits these as the paragraph's margin-top/margin-bottom (paragraph
 *  decoration), so measure must count them or every paragraph under-measures by
 *  its before+after. The cascade matters: a table-cell paragraph with empty
 *  direct attrs still carries its table style's spacing (e.g. before/after=60tw
 *  → 4px each), which is exactly what inflates each rendered cell by ~8px vs a
 *  line-only measurement — so a large table measured short of one page and
 *  overflowed the fixed page box instead of splitting, off by tens of pages. */
export function paragraphSpacingMargins(
  node: PmNode,
  styles: unknown,
): { beforePx: number; afterPx: number } {
  const direct = (node.attrs as { spacing?: SpacingProperties | null }).spacing;
  let before: number | null = direct?.before != null ? Number(direct.before) : null;
  let after: number | null = direct?.after != null ? Number(direct.after) : null;
  if (before == null || after == null) {
    // Style chain via the renderer's mergeStyleChain (direct style + basedOn
    // ancestors; a styleId-less paragraph → default style). A cell paragraph
    // with style "TableHeaderText" (no spacing) inherits "TableText" (spacing
    // 60tw = 4px each) — the renderer resolves the same chain, so measure must
    // too, or every table row under-measures by its paragraph's before+after.
    const styleId = (node.attrs as { styleId?: string | null }).styleId;
    const sp = styleChainOf(styles, styleId)?.paragraph?.spacing as SpacingProperties | undefined;
    if (before == null && sp?.before != null) before = Number(sp.before);
    if (after == null && sp?.after != null) after = Number(sp.after);
    // docDefaults is the base layer under every named style (renderer emits it
    // on p/h1-6); mergeStyleChain covers only the named-style chain.
    const t = styleTableOf(styles);
    const docSp = t?.default?.document?.paragraph?.spacing;
    if (before == null && docSp?.before != null) before = Number(docSp.before);
    if (after == null && docSp?.after != null) after = Number(docSp.after);
  }
  return {
    beforePx: typeof before === "number" ? before * TWIP_TO_PX : 0,
    afterPx: typeof after === "number" ? after * TWIP_TO_PX : 0,
  };
}

/** Paragraph default run properties (pPr/rPr) → run style baseline. Falls back
 *  to the paragraph's style and the document default run when attrs are absent,
 *  so the measured font matches the rendered one — a CJK doc default of 宋体
 *  measures wider than the generic "serif" fallback, which under-counted wrapped
 *  lines. */
function defaultRunOf(node: PmNode, styles?: unknown): Partial<RunStyle> {
  // Style chain via the renderer's mergeStyleChain (direct style + basedOn
  // ancestors; a styleId-less paragraph → default style Normal). Same source as
  // stylesToCss so the measured font matches the rendered one — a CJK doc
  // default of 宋体 measures wider than the generic "serif" fallback.
  const styleId = (node.attrs as { styleId?: string | null }).styleId;
  const run = styleChainOf(styles, styleId)?.run as
    | {
        size?: number | null;
        font?: unknown;
        bold?: boolean;
        italic?: boolean;
      }
    | undefined;
  const t = styleTableOf(styles);
  const defRun = t?.default?.document?.run;
  return {
    // ¶ glyph rPr (attrs.run) styles ONLY the paragraph-mark glyph, NOT run text
    // (ECMA-376: w:pPr/w:rPr applies only to the ¶ glyph). A run therefore
    // inherits font/size/bold/italic from its paragraph STYLE / doc default —
    // NEVER attrs.run — so attrs.run is not read here. Empty-paragraph struts
    // still read attrs.run.size directly via emptyLineHeight (the ¶ glyph is
    // their sole content).
    size: run?.size ?? defRun?.size ?? null,
    font: run?.font ?? defRun?.font ?? null,
    bold: run?.bold ?? defRun?.bold ?? false,
    italic: run?.italic ?? defRun?.italic ?? false,
    characterSpacing: null,
  };
}

interface FontSpec {
  size: number; // points
  font: unknown;
  bold: boolean;
  italic: boolean;
  // Whether the run's text is CJK. docGrid type=lines is a CJK grid — Word
  // snaps line height by the CJK font's single, so the line metric is
  // CJK-dominant (a Latin run alongside CJK must not inflate the grid row).
  isCjk: boolean;
}

function buildFont(spec: FontSpec): string {
  const parts: string[] = [];
  if (spec.italic) parts.push("italic");
  if (spec.bold) parts.push("bold");
  parts.push(`${spec.size}pt`);
  parts.push(resolveFamily(spec.font) ?? DEFAULT_FAMILY);
  return parts.join(" ");
}

// ── line-height: the font's `normal` metric (DOM-measured, incl. line-gap) ──
// getComputedStyle().lineHeight returns "normal" (not a number), and CSS calc()
// cannot reference it. fontNormalRatio measures it via a hidden DOM probe
// (canvas measureText's fontBoundingBox omits the font's line-gap, drifting
// below the rendered `normal`). The ratio is a fixed property of the font
// (size-independent), so the paginator stays deterministic across re-flows —
// only a one-shot DOM probe per (family, bold, italic) on first use.
/** A run's `normal` line-height in px = ratio × size, mirroring the renderer's
 *  `var(--docen-font-metric) × 1em`. */
function normalPxOf(spec: FontSpec): number {
  const sizePx = spec.size * PT_TO_PX;
  return (
    fontNormalRatio({
      family: resolveFamily(spec.font) ?? DEFAULT_FAMILY,
      bold: spec.bold,
      italic: spec.italic,
    }) * sizePx
  );
}

/**
 * Resolve a paragraph's line-height in px (OOXML model, ECMA-376):
 *  1. `spacing.line` (exact/atLeast/auto) applies to ALL paragraphs, INCLUDING
 *     table cells — ECMA-376 docGrid exempts only its own linePitch snap from
 *     table cells (via adjustLineHeightInTable §2.15.3.1), never the paragraph's
 *     spacing.line. exact/atLeast -> fixed twips->px; auto -> multiple (line/240)
 *     × single-line height.
 *  2. The auto "single-line height" = docGrid linePitch when a grid is defined
 *     (docGrid §2.6.2.4: linePitch "defines the pitch for each line ... such
 *     that the desired number of single spaced lines ... fits"), else the font's
 *     natural metric. Verified vs Word: a 1.5× line on a CJK cell with a grid
 *     renders at 1.5×linePitch, not 1.5×natural.
 *  3. No spacing.line + snapToGrid on (pitch > 0):
 *     - table cell -> MAX(natural, pitch) (cell exempts pitch snap, but the row
 *       never goes below a grid row; trHeight floors separately)
 *     - CJK-dominant body -> CEIL(natural / pitch) * pitch (snap UP to a whole
 *       grid row; docGrid type=lines is a CJK grid, CJK chars align to it)
 *     - Latin-dominant body -> MAX(natural, pitch) (Latin chars don't snap)
 *  4. No spacing.line + snapToGrid off / no grid -> natural.
 * `normalPx` + `hasCjk` come from paragraphNormalPx (CJK-dominant metric +
 * CJK-presence flag). Mirrors the renderer's lineSpacingToCss (edit == render). */
export function resolveLineHeight({
  spacing,
  normalPx,
  linePitchPx = 0,
  snapToGrid = null,
  inTable = false,
  hasCjk = false,
}: {
  spacing: SpacingProperties | null | undefined;
  normalPx: number;
  linePitchPx?: number;
  snapToGrid?: boolean | null;
  inTable?: boolean;
  hasCjk?: boolean;
}): number {
  // spacing.line applies to every paragraph incl. table cells (docGrid exempts
  // only its own pitch snap from cells, never spacing.line).
  if (spacing?.line) {
    const rule = spacing.lineRule;
    if (rule === "exact" || rule === "exactly" || rule === "atLeast") {
      return Number(spacing.line) * TWIP_TO_PX;
    }
    // auto: line is 240ths of a SINGLE LINE height — the docGrid linePitch when
    // a grid is defined, else the font natural. Verified vs Word.
    const singleLinePx = linePitchPx > 0 ? linePitchPx : normalPx;
    return (Number(spacing.line) / 240) * singleLinePx;
  }
  // No spacing.line: snap to the document grid.
  const pitch = snapToGrid === false ? 0 : linePitchPx;
  if (pitch > 0) {
    if (inTable) return Math.max(normalPx, pitch);
    if (hasCjk) return Math.ceil(normalPx / pitch) * pitch;
    return Math.max(normalPx, pitch);
  }
  // snapToGrid off / no grid: natural.
  return normalPx;
}

// ── paragraph measurement ──

/** Collect the paragraph's TEXT-run font specs (for the DOM font-metric probe).
 *  Hard breaks / images carry no font metric, so only text runs are included;
 *  an empty paragraph returns the default-run spec (its strut). Shared with the
 *  per-paragraph --docen-font-metric decoration so measure and render agree on
 *  the paragraph's dominant font. */
/** OOXML chooses a run's font by the Unicode range of its text (CJK → eastAsia,
 *  else ascii); `hint` only disambiguates borderline chars (mainly the ¶ glyph).
 *  A run font carrying ONLY a hint (e.g. `{hint:"eastAsia"}` with no concrete
 *  name) inherits the paragraph/doc default font — merge it onto the default as
 *  a base, never let the hint-only object replace it. Returns the family the
 *  renderer applies to this run's text, so the measured line-box metric matches
 *  the rendered one: a CJK run on a 宋体-default doc measures at 宋体's ~1.14, not
 *  generic serif's ~1.44 (which inflated every CJK table row ~3pt vs Word). */
const CJK_RANGE = /[ᄀ-ᇿ⺀-鿿ꥠ-꥿가-힯豈-﫿぀-ヿ＀-￯]/;
function resolveRunFont(runFont: unknown, defFont: unknown, text: string): unknown {
  if (typeof runFont === "string") return runFont;
  const base = defFont && typeof defFont === "object" ? (defFont as Record<string, unknown>) : {};
  const over = runFont && typeof runFont === "object" ? (runFont as Record<string, unknown>) : {};
  const f: Record<string, unknown> = { ...base, ...over };
  const fam = (k: string): string | null => (typeof f[k] === "string" ? (f[k] as string) : null);
  if ((text && CJK_RANGE.test(text)) || f.hint === "eastAsia") {
    return fam("eastAsia") ?? fam("ascii") ?? fam("hAnsi");
  }
  return fam("ascii") ?? fam("hAnsi") ?? fam("eastAsia");
}

export function collectRunSpecs(para: PmNode, def: Partial<RunStyle>): FontSpec[] {
  const specs: FontSpec[] = [];
  const fallback = (): FontSpec => ({
    size: def.size ?? DEFAULT_SIZE_PT,
    font: def.font,
    bold: def.bold ?? false,
    italic: def.italic ?? false,
    isCjk: false,
  });
  para.forEach((child) => {
    if (child.isText) {
      const ms = markStyleOf(child.marks as readonly MarkLike[] | undefined);
      const text = child.text ?? "";
      specs.push({
        size: ms.size ?? def.size ?? DEFAULT_SIZE_PT,
        font: resolveRunFont(ms.font, def.font, text),
        bold: ms.bold ?? def.bold ?? false,
        italic: ms.italic ?? def.italic ?? false,
        isCjk: CJK_RANGE.test(text),
      });
    }
  });
  if (specs.length === 0) specs.push(fallback());
  return specs;
}

/** A paragraph's single-line metric for snapToGrid: the MAX `normal` across its
 *  CJK runs (docGrid type=lines is a CJK grid — Word snaps by the CJK font's
 *  single; a Calibri run alongside CJK must not inflate the grid row), plus a
 *  `hasCjk` flag (CJK lines ceil to a whole pitch multiple; Latin lines don't).
 *  Falls back to the all-run max for pure-Latin paragraphs (hasCjk=false). */
function paragraphNormalPx(para: PmNode, def: Partial<RunStyle>): { px: number; hasCjk: boolean } {
  const specs = collectRunSpecs(para, def);
  let cjkMax = 0;
  let allMax = 0;
  for (const s of specs) {
    const px = normalPxOf(s);
    allMax = Math.max(allMax, px);
    if (s.isCjk) cjkMax = Math.max(cjkMax, px);
  }
  return { px: cjkMax > 0 ? cjkMax : allMax, hasCjk: cjkMax > 0 };
}

/** Whether the paragraph has any CJK text run. docGrid type=lines snaps CJK
 *  chars to the grid (line height ceil to a whole pitch multiple) but leaves
 *  Latin chars on their natural metric; resolveLineHeight / the font-metric
 *  decoration use this to pick ceil vs max. */
export function paragraphHasCjk(node: PmNode, styles: unknown): boolean {
  const def = defaultRunOf(node, styles);
  return collectRunSpecs(node, def).some((s) => s.isCjk);
}

/** Max `normal` RATIO across the paragraph's CJK runs (size-independent) — for
 *  the per-paragraph --docen-font-metric decoration. CJK-dominant (mirrors
 *  paragraphNormalPx): docGrid type=lines snaps by the CJK font, so a Latin run
 *  (e.g. Calibri) alongside CJK must not inflate the metric. Falls back to the
 *  all-run max for pure-Latin paragraphs. */
export function paragraphMaxRatio(node: PmNode, styles: unknown): number {
  const def = defaultRunOf(node, styles);
  const specs = collectRunSpecs(node, def);
  let cjkMax = 0;
  let allMax = 0;
  for (const s of specs) {
    const r = fontNormalRatio({
      family: resolveFamily(s.font) ?? DEFAULT_FAMILY,
      bold: s.bold,
      italic: s.italic,
    });
    allMax = Math.max(allMax, r);
    if (s.isCjk) cjkMax = Math.max(cjkMax, r);
  }
  const max = cjkMax > 0 ? cjkMax : allMax;
  return max > 0 ? max : 1.2;
}

/** Max font SIZE (pt) across the paragraph's runs — feeds the per-paragraph
 *  --docen-line-base decoration so lineSpacingToCss's
 *  `calc(metric × multiple × line-base + pitch)` resolves against the line
 *  box's tallest font, NOT the paragraph's inherited container font-size (which
 *  under-counts large-font runs: a 42pt heading must scale at 42pt, not 14pt). */
export function paragraphMaxSizePt(node: PmNode, styles: unknown): number {
  const def = defaultRunOf(node, styles);
  const specs = collectRunSpecs(node, def);
  // The line box scales at the paragraph's tallest ACTUAL run size, not the
  // inherited default. Flooring at def.size let a 10.5pt table cell render at
  // Normal's 12pt (over-tall rows). Empty paragraphs (no text runs) keep the
  // default-run strut.
  let max: number | null = null;
  for (const s of specs) {
    if (s.size != null) max = max === null ? s.size : Math.max(max, s.size);
  }
  return max ?? def.size ?? DEFAULT_SIZE_PT;
}

/** snapToGrid (w:snapToGrid) on the paragraph: null = OOXML default (true when a
 *  document grid is defined); explicit false drops the grid pitch. */
function readSnapToGrid(node: PmNode): boolean | null {
  return (node.attrs as { snapToGrid?: boolean | null }).snapToGrid ?? null;
}

function collectInlineItems(para: PmNode, def: Partial<RunStyle>): RichInlineItem[] {
  const items: RichInlineItem[] = [];
  const fallbackSpec = (): FontSpec => ({
    size: def.size ?? DEFAULT_SIZE_PT,
    font: def.font,
    bold: def.bold ?? false,
    italic: def.italic ?? false,
    isCjk: false,
  });
  para.forEach((child) => {
    if (child.isText) {
      const ms = markStyleOf(child.marks as readonly MarkLike[] | undefined);
      const text = child.text ?? "";
      const spec: FontSpec = {
        size: ms.size ?? def.size ?? DEFAULT_SIZE_PT,
        font: resolveRunFont(ms.font, def.font, text),
        bold: ms.bold ?? def.bold ?? false,
        italic: ms.italic ?? def.italic ?? false,
        isCjk: CJK_RANGE.test(text),
      };
      const characterSpacing = ms.characterSpacing ?? def.characterSpacing ?? null;
      items.push({
        text: child.text ?? "",
        font: buildFont(spec),
        letterSpacing: characterSpacing != null ? characterSpacing * TWIP_TO_PX : undefined,
      });
    } else if (child.type.name === "hardBreak") {
      items.push({ text: "\n", font: buildFont(fallbackSpec()) });
    }
    // pageBreak and other inline atoms carry no text height — skip.
  });
  if (items.length === 0) {
    // Empty paragraph still occupies one line-height in the flow.
    items.push({ text: "", font: buildFont(fallbackSpec()) });
  }
  return items;
}

/** True if the paragraph has no inline content (no text, hardBreak, or image
 *  children) — its sole content is the ¶ glyph. Such a paragraph's height is
 *  the paragraph-mark line height (see emptyLineHeight), not a text line. */
function isEmptyTextblock(node: PmNode): boolean {
  for (let i = 0; i < node.childCount; i++) {
    const child = node.child(i);
    if (child.isText || child.type.name === "hardBreak" || child.type.name === "image") {
      return false;
    }
  }
  return true;
}

/** Height of a paragraph's strut line — the line-box minimum when there is no
 *  text (an empty paragraph, or an image row shorter than a text line). Mirrors
 *  renderParagraphStyles: spacing.line wins; else the paragraph-mark run size
 *  (pPr/rPr.sz) renders as `line-height:${size}pt` — an ABSOLUTE value, so
 *  markSize × PT_TO_PX (the ¶ glyph's line-height is an absolute pt, not a font
 *  metric); else the font's `normal` metric — the empty ¶ line is the font's
 *  pure natural metric (no grid pitch; verified vs Word). */
function emptyLineHeight({
  node,
  styles,
  linePitchPx,
  normalPx,
}: {
  node: PmNode;
  styles: unknown;
  linePitchPx: number | undefined;
  normalPx: number;
}): number {
  const spacing = resolveSpacing(node, styles);
  // An empty paragraph (¶ glyph only) does NOT receive the grid pitch — Word
  // renders the ¶ at the font's natural metric (verified: an empty single line
  // between paragraphs measures ~natural, not natural+pitch).
  if (spacing?.line) {
    return resolveLineHeight({ spacing, normalPx, linePitchPx, snapToGrid: false });
  }
  const markSize = (node.attrs as { run?: { size?: number | null } | null }).run?.size ?? null;
  if (markSize != null) return markSize * PT_TO_PX;
  // No spacing, no ¶ glyph size: single line = the font's `normal` metric (no
  // grid pitch for the empty ¶ line).
  return resolveLineHeight({ spacing: null, normalPx, linePitchPx, snapToGrid: false });
}

// NOTE: browsers apply CJK kinsoku (line-start/end prohibition) by default — an
// opening bracket at line end or a closing punct at line start moves to the
// next/previous line, adding a line vs Pretext's pure width breaking. Pretext
// exposes no kinsoku, and a deterministic post-pass was never wired up here, so
// a CJK-punctuation paragraph can measure ~1 line short. This is a known
// fidelity limit, not a correctness bug — the paginator converges on the
// measured count.

/** A square/tight floating image (CSS float:left/right) occupies flow space
 *  expressed as a vertical band relative to the page content top: text whose
 *  line overlaps the band wraps beside it (available width = full − zone.width).
 *  Used by layoutLineOffsets so a paragraph measured under a float credits each
 *  line only the width the renderer gives it — otherwise it under-counts wrapped
 *  lines and the page overflows after reflow. wrapNone (position:absolute) and
 *  page/margin anchors are excluded: they sit outside the text flow. */
export interface FloatZone {
  width: number; // image width + L/R wrap margins (px)
  top: number; // page-content-relative top Y
  bottom: number; // top + image height
}

// 1 px = 9525 EMU (914400 EMU/inch ÷ 96 px/inch). Drawing offsets/margins are EMU.
const EMU_PER_PX = 9525;

/** Usable text width for a line occupying [y, y2]: full width minus the widest
 *  float zone overlapping it (CSS float — text wraps beside the image). */
function widthAtY(
  y: number,
  y2: number,
  fullWidth: number,
  zones: readonly FloatZone[] | undefined,
): number {
  if (!zones || zones.length === 0) return fullWidth;
  let reduce = 0;
  for (const z of zones) {
    if (z.bottom > y && z.top < y2) reduce = Math.max(reduce, z.width);
  }
  return Math.max(0, fullWidth - reduce);
}

/** Float zones a paragraph contributes (its square/tight/topAndBottom floating
 *  images, wp:anchor with float:left/right). wrapNone (type 0, position:absolute)
 *  and page/margin/column anchors are skipped — they don't ride the paragraph's
 *  text flow. Each zone's top = `startY` + the image's vertical offset (margin-top,
 *  EMU→px); width = image width + L/R wrap margins (the float's wrap box). Used by
 *  measureFlatItems to track active floats across a page so later paragraphs the
 *  image overhangs measure with the reduced width (text-wrap fidelity). */
export function paragraphFloatZonesOf(node: PmNode, startY: number): FloatZone[] {
  const zones: FloatZone[] = [];
  node.forEach((child) => {
    if (child.type.name !== "image") return;
    const fl = (child.attrs as { floating?: unknown }).floating as
      | {
          wrap?: { type?: number | null } | null;
          verticalPosition?: { relative?: string | null; offset?: number | null } | null;
          margins?: { left?: number | null; right?: number | null } | null;
        }
      | null
      | undefined;
    if (!fl) return;
    if ((fl.wrap?.type ?? 0) === 0) return; // wrapNone: absolute, outside flow
    const vRel = fl.verticalPosition?.relative;
    if (vRel === "page" || vRel === "margin" || vRel === "column") return; // page-box anchor
    const w = (child.attrs as { width?: number }).width;
    const h = (child.attrs as { height?: number }).height;
    if (typeof w !== "number" || typeof h !== "number" || w <= 0 || h <= 0) return;
    const emuPx = (e: number | null | undefined): number =>
      typeof e === "number" && e > 0 ? e / EMU_PER_PX : 0;
    const width = w + emuPx(fl.margins?.left) + emuPx(fl.margins?.right);
    const top = startY + emuPx(fl.verticalPosition?.offset);
    zones.push({ width, top, bottom: top + h });
  });
  return zones;
}

/** Lay out a prepared paragraph line by line: the FIRST line at
 *  `width − firstLinePx` and later lines at `width` — mirroring CSS text-indent
 *  (only the first line is indented). Returns the line count and the PM content
 *  offset at the END of each line (a legal split point). Pretext's walk/stats
 *  take a single width for all lines and can't express a narrower first line,
 *  which under-counted the first line's capacity; layoutNextRichInlineLineRange
 *  lets us shrink only line 1. Offset mapping: each Pretext item is 1:1 with a
 *  PM inline child (text: 1 char = 1 offset; hardBreak: "\n" = 1 offset), so
 *  accumulating fragment text lengths == accumulating PM content offsets. */
function layoutLineOffsets(
  prepared: PreparedRichInline,
  width: number,
  firstLinePx: number,
  zones?: readonly FloatZone[],
  startY?: number,
  lh?: number,
): { lineCount: number; lineBreakOffsets: number[] } {
  const lineBreakOffsets: number[] = [];
  let pmOff = 0;
  let lineCount = 0;
  let cursor: RichInlineCursor | undefined;
  let first = true;
  // When float zones are present, each line's usable width shrinks by any zone
  // overlapping it (CSS float text-wrap), measured line by line via Pretext's
  // per-line width. Without this a paragraph beside a float under-counts lines
  // and the page overflows after reflow.
  const haveZones =
    !!zones && zones.length > 0 && startY != null && typeof lh === "number" && lh > 0;
  while (true) {
    const baseW = haveZones
      ? widthAtY(startY! + lineCount * lh!, startY! + (lineCount + 1) * lh!, width, zones)
      : width;
    const mw = first ? Math.max(0, baseW - firstLinePx) : baseW;
    const range = layoutNextRichInlineLineRange(prepared, mw, cursor);
    if (!range) break;
    const line = materializeRichInlineLineRange(prepared, range);
    let chars = 0;
    for (const frag of line.fragments) chars += frag.text.length;
    const endCursor: RichInlineCursor = range.end;
    cursor = endCursor;
    pmOff += chars;
    lineBreakOffsets.push(pmOff);
    first = false;
    lineCount++;
  }
  if (lineCount === 0) lineCount = 1; // empty paragraph still occupies one line
  return { lineCount, lineBreakOffsets };
}

/** Measure a textblock (paragraph/heading) height in px at `width`. Line count
 *  mirrors the renderer: first line at width−firstLineIndent, later lines at
 *  width (text-indent shrinks only line 1). `linePitchPx` applies the section
 *  document grid (see resolveLineHeight). */
export function measureParagraphHeight(node: PmNode, width: number, ctx?: MeasureContext): number {
  if (!node.isTextblock || width <= 0) return 0;
  const def = defaultRunOf(node, ctx?.styles);
  const styles = ctx?.styles;
  const linePitchPx = ctx?.linePitchPx;
  const snapToGrid = readSnapToGrid(node);
  const inTable = ctx?.inTable ?? false;
  // Max `normal` metric across the paragraph's runs (Word: a line box is as
  // tall as its tallest font).
  const { px: normalPx, hasCjk } = paragraphNormalPx(node, def);
  const items = collectInlineItems(node, def);
  const prepared = getPrepared(items);
  const firstLinePx = resolveFirstLineIndentPx(node, styles);
  const strut = emptyLineHeight({ node, styles, linePitchPx, normalPx });
  // Empty paragraph (no text/image children): its sole content is the ¶ glyph,
  // whose height is the paragraph-mark line height (strut) — NOT a text line.
  if (isEmptyTextblock(node)) return strut;
  // Image-only paragraph: lay out images by content width (mirrors renderHTML
  // — each inline image occupies its display width × height, clamped to the
  // content width like the renderer's `img { max-width: 100%; height: auto }`;
  // multiple images share a line until their combined width exceeds the content
  // width, then wrap). Each row's height is its tallest image but never below
  // the strut (a row of small images still occupies a full text line — the
  // line-box minimum from font-size/line-height). Total = sum of per-row
  // heights. Previously each image was credited its own line, over-counting
  // side-by-side images by multiples — the paginator packed them on one page
  // and clipped the rest.
  if (hasInlineImage(node)) {
    const layout = layoutImageLines(node, width, strut);
    if (layout) return layout.lineHeights.reduce((s, h) => s + h, 0);
    // Mixed text+image paragraph: conservative upper bound — text line height
    // plus each image at least the strut (mixed paragraphs are not split; they
    // move whole — over-estimate rather than overflow). Precise wrapping needs
    // a unified text+image line breaker (Pretext's flow holds no image atom);
    // out of scope here.
    const lh = resolveLineHeight({
      spacing: resolveSpacing(node, styles),
      normalPx,
      linePitchPx,
      snapToGrid,
      inTable,
      hasCjk,
    });
    let imgH = 0;
    node.forEach((child) => {
      if (child.type.name === "image") {
        const h = (child.attrs as { height?: number }).height;
        imgH += typeof h === "number" && h > 0 ? Math.max(h, strut) : strut;
      }
    });
    const lineCount = layoutLineOffsets(prepared, width, firstLinePx).lineCount;
    return Math.max(1, lineCount) * lh + imgH;
  }
  const lh = resolveLineHeight({
    spacing: resolveSpacing(node, styles),
    normalPx,
    linePitchPx,
    snapToGrid,
    inTable,
    hasCjk,
  });
  const lcInfo = layoutLineOffsets(prepared, width, firstLinePx, ctx?.floatZones, ctx?.startY, lh);
  return Math.max(1, lcInfo.lineCount) * lh;
}

/** True if the paragraph contains an inline image child. Such paragraphs
 *  cannot be split mid-paragraph — the image is not in Pretext's inline flow
 *  (collectInlineItems skips it), so per-line offsets would drift — so the
 *  paginator moves them whole. */
/** True if the paragraph contains a FLOATING image (wp:anchor, attrs.floating).
 *  Such paragraphs are NOT split: a floating image anchors to the paragraph and
 *  floats outside the text flow (behind/around text), so splitting the paragraph
 *  would strand the anchor and fan out empty head fragments. The paginator moves
 *  the whole paragraph instead. (Inline wp:inline images — no attrs.floating —
 *  ARE splittable at row boundaries, see hasInlineImage/layoutImageLines.) */
function hasFloatingImage(para: PmNode): boolean {
  let found = false;
  para.forEach((child) => {
    if (child.type.name === "image" && (child.attrs as { floating?: unknown }).floating)
      found = true;
  });
  return found;
}

/** True if the paragraph contains an INLINE image (wp:inline, no attrs.floating). */
function hasInlineImage(para: PmNode): boolean {
  let found = false;
  para.forEach((child) => {
    if (child.type.name === "image" && !(child.attrs as { floating?: unknown }).floating)
      found = true;
  });
  return found;
}

/** True if the paragraph carries a pageBreak atom. Such paragraphs are NOT
 *  split mid-paragraph: the break atom must stay with the paragraph so it lands
 *  at the END of its page (Word — a page/section break never opens a new page).
 *  Splitting would push the atom into the tail (next page), so the paginator
 *  moves the whole paragraph and forcesPageBreakAfter closes the page after it. */
function hasPageBreakAtom(para: PmNode): boolean {
  let found = false;
  para.forEach((child) => {
    if (child.type.name === "pageBreak") found = true;
  });
  return found;
}

/** Per-line break offsets for a textblock: `lineBreakOffsets[i]` is the PM
 *  content offset at the END of line i+1 (a legal split point — everything
 *  before it is the first i+1 lines). The first line wraps at width−firstLine
 *  indent, later lines at the full width, matching the renderer. Deterministic
 *  via Pretext (canvas measureText). Returns null for non-textblocks, zero
 *  width, or paragraphs with inline images (not splittable). */
export function measureParagraphLines(
  node: PmNode,
  width: number,
  ctx?: MeasureContext,
): {
  lineHeight: number;
  lineCount: number;
  lineBreakOffsets: number[];
  block?: boolean;
  lineHeights?: number[];
} | null {
  if (!node.isTextblock || width <= 0) return null;
  // A floating-image paragraph is not split — its anchor must stay with the
  // paragraph. Splitting it produced empty head fragments that fanned out one
  // per page and the doc oscillated forever between split/merged states.
  if (hasFloatingImage(node)) return null;
  // Image-only paragraph: lay out images by content width and split at row
  // boundaries — the PM offset at the end of each row is a legal break (Word
  // breaks an inline-image paragraph across pages between rows, never through
  // an image). `block: true` flags image rows so trySplitBlock skips
  // widow/orphan (an image can sit alone on a page); `lineHeights` lets it
  // accumulate each row's real height (image rows are not uniform). A mixed
  // text+image paragraph returns null (not split — conservative).
  if (hasInlineImage(node)) {
    const def = defaultRunOf(node, ctx?.styles);
    const strut = emptyLineHeight({
      node,
      styles: ctx?.styles,
      linePitchPx: ctx?.linePitchPx,
      normalPx: paragraphNormalPx(node, def).px,
    });
    const layout = layoutImageLines(node, width, strut);
    if (!layout) return null;
    const { lineHeights, lineBreakOffsets } = layout;
    const total = lineHeights.reduce((s, h) => s + h, 0);
    return {
      lineHeight: lineHeights.length > 0 ? total / lineHeights.length : 0,
      lineCount: lineHeights.length,
      lineBreakOffsets,
      block: true,
      lineHeights,
    };
  }
  // A text-only paragraph carrying a pageBreak atom is not split — the atom
  // must stay with the paragraph so it lands at its page's end (Word: a break
  // never opens a new page). Splitting would strand the atom on the tail. An
  // inline-IMAGE paragraph is handled above: it splits at image-row boundaries
  // whose offsets sit before the trailing atom, so the atom stays on the tail
  // and forcesPageBreakAfter closes the tail page. Excluding image paragraphs
  // here let a many-image + pageBreak paragraph return null → never split →
  // the whole multi-thousand-px paragraph sat on one page and overflowed.
  if (hasPageBreakAtom(node)) return null;
  // A section-ending paragraph (carries sectPr) is not split either —
  // sectionProperties must stay on the section's LAST paragraph, and
  // forcesPageBreakAfter closes the page after it. Splitting would strand the
  // sectPr on the head and push the tail onto the next section's first page.
  if ((node.attrs as { sectionProperties?: unknown }).sectionProperties != null) return null;
  const def = defaultRunOf(node, ctx?.styles);
  const items = collectInlineItems(node, def);
  const prepared = getPrepared(items);
  const { px: lineNormalPx, hasCjk } = paragraphNormalPx(node, def);
  const lineHeight = resolveLineHeight({
    spacing: resolveSpacing(node, ctx?.styles),
    normalPx: lineNormalPx,
    linePitchPx: ctx?.linePitchPx,
    snapToGrid: readSnapToGrid(node),
    inTable: ctx?.inTable ?? false,
    hasCjk,
  });
  const firstLinePx = resolveFirstLineIndentPx(node, ctx?.styles);
  const { lineCount, lineBreakOffsets } = layoutLineOffsets(
    prepared,
    width,
    firstLinePx,
    ctx?.floatZones,
    ctx?.startY,
    lineHeight,
  );
  return { lineHeight, lineCount, lineBreakOffsets };
}

/** Lay out an image-only paragraph by content width, mirroring renderHTML: each
 *  inline image occupies attrs.width × attrs.height px, and multiple images sit
 *  on the same line until their combined width exceeds the content width, then
 *  wrap (the same inline rule as text characters). Returns each row's height
 *  (its tallest image) and the PM content offset at the end of each row (a
 *  legal split point). A paragraph mixing text/hardBreak returns null (caller
 *  falls back to not splitting), because Pretext's inline flow holds no image
 *  atom and a unified text+image breaker is out of scope. A single image wider
 *  than the content width sits alone on a row (its width overflows and is
 *  clipped by the page box; its height is still counted). */
function layoutImageLines(
  node: PmNode,
  width: number,
  minHeight: number,
): { lineHeights: number[]; lineBreakOffsets: number[] } | null {
  const lineHeights: number[] = [];
  const lineBreakOffsets: number[] = [];
  let lineW = 0;
  let lineMaxH = 0;
  let lineEndOff = 0;
  let off = 0;
  let imgCount = 0;
  let hasOther = false;
  node.forEach((child) => {
    off += child.nodeSize;
    if (child.type.name === "image") {
      imgCount++;
      const w = (child.attrs as { width?: number }).width;
      const h = (child.attrs as { height?: number }).height;
      let iw = typeof w === "number" && w > 0 ? w : 0;
      let ih = typeof h === "number" && h > 0 ? h : 0;
      // 无尺寸图（attrs 无 width/height，如未精化的 http 图）占位：与
      // image.ts renderHTML 的 CSS（width:100% + aspect-ratio:4/3）一致，
      // 避免分页用 strut（0 高）导致几百张离屏 lazy 图全堆少数页。fetch 完
      // image-cap 精化设实际尺寸后 attrs 有值，此分支跳过。
      if (iw === 0 && ih === 0) {
        iw = width;
        ih = Math.round(width * 0.75);
      }
      // The renderer clamps a wide image with `img { max-width: 100%; height:
      // auto }` so an oversized import never overflows the page. Mirror that
      // here: an image wider than the content width is credited the content
      // width and a height scaled by the same factor, else the row is
      // over-counted and the page breaks early. image-cap leaves an import's
      // source extent untouched (for fidelity), so this clamp must live here.
      if (iw > width) {
        ih = ih > 0 ? Math.round(ih * (width / iw)) : 0;
        iw = width;
      }
      // Current row already has content and adding this image overflows → flush
      // the row and start a new one.
      if (lineW > 0 && lineW + iw > width) {
        lineHeights.push(Math.max(lineMaxH, minHeight));
        lineBreakOffsets.push(lineEndOff);
        lineW = 0;
        lineMaxH = 0;
      }
      lineW += iw;
      lineMaxH = Math.max(lineMaxH, ih);
      lineEndOff = off;
    } else if (child.isText || child.type.name === "hardBreak") {
      hasOther = true;
    }
  });
  if (imgCount === 0 || hasOther) return null;
  lineHeights.push(Math.max(lineMaxH, minHeight));
  lineBreakOffsets.push(lineEndOff);
  return { lineHeights, lineBreakOffsets };
}

/** Effective OOXML pagination properties for a paragraph/heading. `null`
 *  (the document didn't set the property) resolves to the OOXML DEFAULT:
 *  widowControl ON (the spec's default; @office-open leaves it undefined when
 *  `<w:widowControl/>` is absent, so attrs.null → ON), the others OFF.
 *  No style-hierarchy fallback yet (stage 2). */
export interface PaginationAttrs {
  keepLines: boolean;
  keepNext: boolean;
  widowControl: boolean;
  pageBreakBefore: boolean;
}

export function resolvePaginationAttrs(node: PmNode): PaginationAttrs {
  const a = node.attrs as {
    keepLines?: boolean | null;
    keepNext?: boolean | null;
    widowControl?: boolean | null;
    pageBreakBefore?: boolean | null;
  };
  const isHeading = node.type.name === "heading";
  return {
    keepLines: a.keepLines === true,
    // Word's heading styles default to keepNext=true (a heading never orphans
    // at a page bottom — it stays with the following paragraph). The OOXML attr
    // reads null when the style's default applies, so a heading with no explicit
    // keepNext still keeps; a non-heading only keeps when explicitly set.
    keepNext: isHeading ? a.keepNext !== false : a.keepNext === true,
    widowControl: a.widowControl !== false,
    pageBreakBefore: a.pageBreakBefore === true,
  };
}

/** A paragraph's effective indent, resolved through the same cascade the
 *  renderer uses: direct attrs → its style (styleId) → the document default.
 *  The style/default indent reaches the DOM via stylesToCss (a CSS class on the
 *  paragraph), NOT node attrs, so reading only attrs misses it — which is why a
 *  doc whose first-line indent lives in docDefaults measured at the full page
 *  width, under-counted short paragraphs by a line, and overflowed the page. */
type IndentAttrs = {
  left?: number | null;
  right?: number | null;
  firstLine?: number | null;
  firstLineChars?: number | null;
  hanging?: number | null;
};

function resolveIndentAttrs(node: PmNode, styles: unknown): IndentAttrs | null {
  const direct = (node.attrs as { indent?: IndentAttrs | null }).indent;
  if (direct) return direct;
  // Style chain via the renderer's mergeStyleChain (direct style + basedOn
  // ancestors); a styleId-less paragraph → default style. Same source as
  // stylesToCss so measure == render.
  const styleId = (node.attrs as { styleId?: string | null }).styleId;
  const indent = styleChainOf(styles, styleId)?.paragraph?.indent as IndentAttrs | undefined;
  if (indent) return indent;
  const t = styleTableOf(styles);
  const docIndent = t?.default?.document?.paragraph?.indent;
  if (docIndent) return docIndent as IndentAttrs;
  return null;
}

/** A textblock's USABLE wrapping width in px — mirrors renderParagraphStyles
 *  (utils.ts): `indent.left/right` → `margin-left/right` (shrink the content
 *  box); `indent.firstLine` (twips) or `firstLineChars` (→ text-indent N/100 em)
 *  shrinks the FIRST line. Pretext wraps at a single width for all lines, so
 *  the first-line indent is subtracted from the whole block — conservative
 *  (may count a hair more lines than render, never fewer), which beats the
 *  opposite (under-count → overflow). Twips→px and em→px match the renderer. */
export function resolveIndentWidth(
  node: PmNode,
  pageContentWidth: number,
  styles?: unknown,
): number {
  const indent = resolveIndentAttrs(node, styles);
  if (!indent) return pageContentWidth;
  const twipToPx = 4 / 3 / 20;
  const left = (indent.left ?? 0) * twipToPx;
  const right = (indent.right ?? 0) * twipToPx;
  // NOTE: first-line indent is NOT subtracted here — it shrinks only the first
  // line, not the whole block. It's applied per-line in layoutLineOffsets so
  // later lines keep the full width (matching the renderer). Subtracting it here
  // under-counted every line's capacity and left a split head's last line short.
  return Math.max(0, pageContentWidth - left - right);
}

/** The paragraph's first-line indent in px (text-indent shrinks ONLY the first
 *  line). Resolved through the same cascade as renderParagraphStyles +
 *  stylesToCss. Used by layoutLineOffsets so the first line wraps at
 *  width−indent while later lines wrap at the full width. */
function resolveFirstLineIndentPx(node: PmNode, styles?: unknown): number {
  const indent = resolveIndentAttrs(node, styles);
  if (!indent) return 0;
  const twipToPx = 4 / 3 / 20;
  if (indent.firstLine != null) return Math.max(0, indent.firstLine) * twipToPx;
  if (indent.firstLineChars != null && indent.firstLineChars > 0) {
    const sizePt = defaultRunOf(node, styles).size ?? DEFAULT_SIZE_PT;
    return (indent.firstLineChars / 100) * sizePt * PT_TO_PX;
  }
  return 0;
}

// ── block measurement (paragraph / leaf / container) ──

export interface MeasureContext {
  /** DOM height (px) for non-text leaves Pretext cannot measure (images,
   *  embeds). Read once after layout settles; never used for text blocks. */
  domHeightOf?: (node: PmNode) => number | undefined;
  /** Section document-grid pitch (px) — lines snap up to it (Word w:docGrid). */
  linePitchPx?: number;
  /** Document style table (doc.attrs.styles) for spacing/font fallback when a
   *  paragraph carries no direct attrs — measure must mirror the renderer's
   *  style cascade or rows whose spacing lives in a style measure too short. */
  styles?: unknown;
  /** True inside a table cell: line height snaps to max(natural, pitch) so the
   *  row's trHeight atLeast floor — not the line box — governs (matches Word). */
  inTable?: boolean;
  /** Active float zones (page-content-relative Y): square/tight float images
   *  reduce each overlapping line's usable width (text wraps beside them). An
   *  empty/absent list ⇒ measure at full width (no float influence). Table
   *  cells clear this (a cell's width is its column, not the page flow). */
  floatZones?: readonly FloatZone[];
  /** This block's top Y within the page content box — used with floatZones to
   *  derive each line's Y for the per-line width reduction. */
  startY?: number;
  /** Table-level cell insets (table.attrs.margins, the w:tblCellMar default).
   *  Passed to measureRowHeight so a cell lacking its own tcMar inherits the
   *  table default — the same effective source the renderer reads. Without it a
   *  reflow-split clone with transiently-missing margins under-measures. */
  tableCellMargins?: CellMargins | null;
}

/** Measure any top-level block height in px at `width`. Textblocks use Pretext
 *  (deterministic — same input → same output every re-flow pass; a textblock's
 *  DOM height wobbles sub-pixel across layout passes, and preferring it makes
 *  page-count oscillate as boundary blocks flip pages each pass — see file
 *  header); leaves (images) fall back to DOM; containers (list/blockquote)
 *  recurse. */
export function measureBlockHeight(node: PmNode, width: number, ctx?: MeasureContext): number {
  if (node.isTextblock) {
    return measureParagraphHeight(node, width, ctx);
  }
  if (node.isLeaf) {
    return ctx?.domHeightOf?.(node) ?? 0;
  }
  let h = 0;
  node.forEach((child) => {
    h += measureBlockHeight(child, width, ctx);
  });
  return h;
}

// ── table row measurement ──
// Width math mirrors `packages/docx/src/extensions/table.ts` renderHTML so the
// measured column widths match the rendered colgroup (else text wraps to a
// different width than measured and row heights drift).

/** Table content width (px) from its OOXML width attr: pct → % of
 *  pageContentWidth; auto/none → fill the text column; dxa → twips→px. */
export function tableWidthOf(table: PmNode, pageContentWidth: number): number {
  const w = (table.attrs as { width?: TableWidthProperties | null }).width;
  if (!w) return pageContentWidth;
  const num = typeof w.size === "string" ? parseFloat(w.size) : w.size;
  if (w.type === "pct") {
    // office-open normalizes OOXML fiftieths (0-5000) to a percentage number
    // (0-100) on parse, so a numeric size IS the percentage — dividing by 50
    // here (the pre-normalization assumption) collapses a 99.96% table to ~2%
    // width, shrinking every column and inflating measured row heights, which
    // then forces whole tables to the next page despite ample room. A string
    // like "100%" is already literal. Mirrors the pct handling in
    // @docen/docx table renderHTML / DocenTableView.
    const pct = typeof w.size === "string" && w.size.includes("%") ? parseFloat(w.size) : num;
    return (pageContentWidth * pct) / 100;
  }
  if (w.type === "auto") return pageContentWidth;
  return num * TWIP_TO_PX;
}

/** Per-column content widths (px) for a table at `tableWidth`. The grid comes
 *  from the first row's cell colwidths (px) — the real per-cell widths from
 *  DOCX <w:tcW> — else tblGrid columnWidths (twips→px); both scaled to
 *  tableWidth by ratio (Word scales the grid to tblW, never to the raw sum). */
export function tableColumnWidths(table: PmNode, tableWidth: number): number[] {
  // Prefer the table grid (tblGrid, via the columnWidths attr in twips): it is
  // IDENTICAL across every split slice of the same table, so column widths —
  // and thus row heights and split points — stay stable across re-flows. The
  // previous order (first-row cell colwidths first) read each slice's own first
  // row, which for a continuation slice is a mid-table row whose cell layout
  // need not match the grid (merged cells, fewer columns); that made row
  // heights wobble and the table oscillated between split counts every pass.
  const gridPx: number[] = [];
  const cols = (table.attrs as { columnWidths?: number[] | null }).columnWidths;
  if (Array.isArray(cols) && cols.length) {
    for (const w of cols) gridPx.push(Math.round((w || 0) / 15));
  }
  if (!gridPx.some((w) => w > 0)) {
    const firstRow = table.firstChild;
    if (firstRow) {
      firstRow.forEach((cell) => {
        const cw = (cell.attrs as { colwidth?: number[] | null }).colwidth;
        const span = (cell.attrs as { colspan?: number }).colspan ?? 1;
        if (Array.isArray(cw) && cw.length) for (const w of cw) gridPx.push(w || 0);
        else for (let i = 0; i < span; i++) gridPx.push(0);
      });
    }
  }
  const total = gridPx.reduce((a, b) => a + b, 0) || 1;
  return gridPx.map((w) => (w / total) * tableWidth);
}

/** Measure a table row height (px): the tallest cell's stacked block heights at
 *  that cell's column width(s), then floored by the row's trHeight (Word
 *  w:trHeight). rowspan cells count fully on their start row (Word distributes
 *  across spans; this over-estimates the start row and under-estimates the
 *  spanned rows, but the measurement is deterministic — no DOM wobble between
 *  re-flows). Cell margins/padding are not yet added.
 *
 *  trHeight: `atLeast` (absent) is a minimum — a row whose cells render shorter
 *  still reserves trHeight (Word pushes whole short rows to fit). `exact` fixes
 *  the row at trHeight (content overflows but the row does not grow). Without
 *  this floor, a compact table measured shorter than its trHeight sum and spilled
 *  its last row onto the next page even though Word keeps it whole. */
export function measureRowHeight(
  row: PmNode,
  columnWidths: number[],
  ctx?: MeasureContext,
): number {
  // Cells measure under the table-cell line-height rule (max(natural, pitch))
  // so the row's trHeight floor governs — fork the ctx once for the whole row.
  const cellCtx: MeasureContext = {
    ...ctx,
    inTable: true,
    floatZones: undefined,
    startY: undefined,
  };
  let maxHeight = 0;
  let colCursor = 0;
  row.forEach((cell) => {
    const colspan = (cell.attrs as { colspan?: number }).colspan ?? 1;
    const cellWidth = columnWidths.slice(colCursor, colCursor + colspan).reduce((a, b) => a + b, 0);
    colCursor += colspan;
    // Effective cell insets: a cell's own tcMar wins, else inherit the table's
    // tblCellMar (table.attrs.margins, threaded via ctx) — mirroring the
    // renderer, which inherits tblCellMar (resolveTable pushes it onto cells at
    // parse). A cell cloned during a reflow split can transiently lack its own
    // margins; without this fallback measure reads zero insets, over-estimates
    // innerW, and under-counts wrapped lines (a narrow cell can measure ~half
    // its rendered line count), packing too many rows per page until the table
    // overflows the page bottom. Per-side: a present side wins, else table's.
    const cellMargins = effectiveCellMargins(cell, ctx?.tableCellMargins);
    const hMarginPx =
      (marginSizeTw(cellMargins.left) + marginSizeTw(cellMargins.right)) * TWIP_TO_PX;
    // Subtract the cell's left+right borders: under border-collapse the grid
    // column width is the cell's BORDER box, so the text wraps at cellWidth −
    // padding − borders. Omitting borders left innerW ~1.3px wider than the
    // rendered content box — enough, in a narrow CJK column, for Pretext to
    // fit a second char on boundary lines and under-count by several lines/row,
    // cumulating into page-bottom overflow. Mirrors the DOM content box exactly.
    const b = (cell.attrs as { borders?: { left?: BorderEdge; right?: BorderEdge } | null })
      .borders;
    const hBorderPx = borderEdgePx(b?.left) + borderEdgePx(b?.right);
    const innerW = Math.max(0, cellWidth - hMarginPx - hBorderPx);
    let contentH = 0;
    let prevAfter = 0;
    let firstBlock = true;
    cell.forEach((block) => {
      const lineH = measureBlockHeight(block, innerW, cellCtx);
      // Paragraph spacing.before/after renders as margin-top/bottom. The cell is
      // a BFC: adjacent paragraphs inside still collapse vertically (max of the
      // previous after and the current before), but the first before and the
      // last after are NOT collapsed against the cell border — the cell eats
      // both, so a single-paragraph cell reserves before+after. Without this
      // every table row under-measured by its paragraph's before+after and a
      // 41-row table spilled a page short (overflowed instead of splitting).
      const { beforePx, afterPx } = paragraphSpacingMargins(block, ctx?.styles);
      contentH += firstBlock ? beforePx : Math.max(prevAfter, beforePx);
      contentH += lineH;
      prevAfter = afterPx;
      firstBlock = false;
    });
    contentH += prevAfter; // last paragraph's after — the cell eats it
    // Add the renderer's vertical cell overhead (padding + border) so the
    // measured row matches the rendered row — without it rows under-measure
    // and the table's last rows overflow the page bottom (clipped).
    contentH += cellVerticalOverhead(cell, cellMargins);
    maxHeight = Math.max(maxHeight, contentH);
  });
  const h = (row.attrs as { height?: { value?: number; rule?: string } | null }).height;
  if (h && typeof h.value === "number" && h.value > 0) {
    const px = h.value * TWIP_TO_PX;
    if (h.rule === "exact") return px;
    maxHeight = Math.max(maxHeight, px);
  }
  return maxHeight;
}

/** Vertical padding + border the renderer adds around a table cell's content
 *  (edit == render). Without it the paginator under-measures rows and the last
 *  rows overflow the page bottom (clipped by the page's overflow:hidden).
 *  - padding: w:tcMar (twips→px); OOXML defaults tcMar top/bottom to 0
 *    (TableNormal) — the renderer's td is padding-block:0 unless tcMar sets it,
 *    so measure matches (no invented UA offset; the earlier 2px default inflated
 *    every CJK table row ~2px vs Word).
 *  - border: w:tcBorders (w:sz = eighths-of-pt → px); nil/none → 0, an absent
 *    side inherits the Table-Grid default (1px = the CSS border). Under
 *    border-collapse:collapse only the MAX of top/bottom adds to the row height
 *    (adjacent rows share one border). */
/** A cell margin side's size in twips (w:tcMar / w:tblCellMar
 *  TableWidthProperties { size, type }), or 0 when absent. Shared by
 *  cellVerticalOverhead (top/bottom → row height) and measureRowHeight
 *  (left/right → wrapping-width inset). */
/** Per-side cell insets (w:tcMar), each a TableWidthProperties { size (twips),
 *  type }. Same shape on a cell (cell.attrs.margins) and a table
 *  (table.attrs.margins — the w:tblCellMar default cells inherit). */
export type CellMargins = {
  top?: TableWidthProperties | null;
  bottom?: TableWidthProperties | null;
  left?: TableWidthProperties | null;
  right?: TableWidthProperties | null;
};

/** Effective cell insets: a cell's own tcMar wins per side, else the table's
 *  tblCellMar (table.attrs.margins) — mirroring resolveTable's parse-time push
 *  and renderTableCellStyles' inheritance. A cell cloned during a reflow split
 *  can transiently lack margins; resolving here keeps measure == render for the
 *  cell's inner width even then. Returns a shape (sides may be absent —
 *  marginSizeTw reads .size as 0). */
function effectiveCellMargins(cell: PmNode, tableMargins?: CellMargins | null): CellMargins {
  const c = (cell.attrs as { margins?: CellMargins | null }).margins ?? null;
  if (!tableMargins) return c ?? {};
  if (!c) return tableMargins;
  return {
    top: c.top ?? tableMargins.top ?? null,
    bottom: c.bottom ?? tableMargins.bottom ?? null,
    left: c.left ?? tableMargins.left ?? null,
    right: c.right ?? tableMargins.right ?? null,
  };
}

function marginSizeTw(s?: TableWidthProperties | null): number {
  return s && typeof s.size === "number" ? s.size : 0;
}

function cellVerticalOverhead(cell: PmNode, margins?: CellMargins | null): number {
  // Cell insets (w:tcMar, or the table's tblCellMar pushed by resolveTable):
  // each side is TableWidthProperties { size (twips), type } — read .size, not
  // the side object (which coerces to NaN). Only top/bottom grow the row height.
  // `margins` is the already-resolved effective set (cell ?? table) from the
  // caller; falling back to cell.attrs.margins keeps other call sites working.
  const m = margins ?? (cell.attrs as { margins?: CellMargins | null }).margins ?? null;
  const padPx = (marginSizeTw(m?.top) + marginSizeTw(m?.bottom)) * TWIP_TO_PX;
  const b = (
    cell.attrs as {
      borders?: {
        top?: BorderEdge;
        bottom?: BorderEdge;
        left?: BorderEdge;
        right?: BorderEdge;
      } | null;
    }
  ).borders;
  // border-collapse:collapse merges adjacent rows' top/bottom borders into one
  // shared line (thicker wins) — only the MAX adds to the row height. Summing
  // both double-counted a 1px border as 2px, over-measuring rows ~1px and
  // pushing CJK tables onto extra pages vs Word.
  const borderPx = Math.max(borderEdgePx(b?.top), borderEdgePx(b?.bottom));
  return padPx + borderPx;
}

/** A cell border edge: { style, size (eighths of a point) } — OOXML's w:sz. */
type BorderEdge = { style?: string; size?: number } | null | undefined;

// The "Table Grid" fallback the renderer stamps on cells whose OOXML left a
// side open (`.docen-pages table td { border: 1px solid }`): 0.75pt = 1px.
const TABLE_GRID_BORDER_PX = 0.75 * PT_TO_PX;

/** One border edge's rendered width (px). An explicit real border (size in
 *  eighths-of-pt, style not nil/none) wins; an absent/nil side falls back to
 *  the Table-Grid default the renderer applies. Mirrors renderBorderCSS so
 *  measure == render for cell content boxes. */
function borderEdgePx(s: BorderEdge): number {
  if (s && s.style && s.style !== "nil" && s.style !== "none" && s.size != null)
    return (s.size / 8) * PT_TO_PX;
  return TABLE_GRID_BORDER_PX;
}
