import {
  prepareRichInline,
  layoutNextRichInlineLineRange,
  materializeRichInlineLineRange,
  type RichInlineItem,
  type RichInlineCursor,
} from "@chenglou/pretext/rich-inline";
import type { Node as PmNode } from "@tiptap/pm/model";

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
  normalLineHeightCache.clear();
}

function getPrepared(items: RichInlineItem[]): PreparedRichInline {
  const key = items.map((it) => `${it.text}\u0000${it.font}\u0000${it.letterSpacing ?? ""}`).join("");
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
    paragraph?: {
      spacing?: { line?: number | null; lineRule?: string | null } | null;
      indent?: IndentAttrs | null;
    };
    run?: { font?: unknown; size?: number | null; bold?: boolean; italic?: boolean };
  }>;
  default?: {
    document?: {
      paragraph?: {
        spacing?: { line?: number | null; lineRule?: string | null } | null;
        indent?: IndentAttrs | null;
      };
      run?: { font?: unknown; size?: number | null; bold?: boolean; italic?: boolean };
    };
  };
}

function styleTableOf(styles: unknown): StyleTable | null {
  return styles && typeof styles === "object" ? (styles as StyleTable) : null;
}

/** A paragraph's effective spacing: direct attr, else its style's, else the
 *  document default. Without the style/default fallback, paragraphs whose
 *  line-spacing lives in their style (e.g. line=360 on a "Table" style → 1.5×
 *  grid pitch = 31.2px) measured at the bare grid pitch (20.8px), so table rows
 *  measured ~1.5× too short and the paginator never split them. */
function resolveSpacing(
  node: PmNode,
  styles: unknown,
): { line?: number | null; lineRule?: string | null } | null {
  const direct = (
    node.attrs as {
      spacing?: { line?: number | null; lineRule?: string | null } | null;
    }
  ).spacing;
  if (direct && direct.line != null) return direct;
  const t = styleTableOf(styles);
  if (t) {
    const styleId = (node.attrs as { styleId?: string | null }).styleId;
    const ps = styleId ? t.paragraphStyles?.find((p) => p.id === styleId) : null;
    if (ps?.paragraph?.spacing?.line != null) return ps.paragraph.spacing;
    // A paragraph with NO styleId implicitly uses the default paragraph style
    // (w:default="1", usually "Normal") — exactly what the renderer's
    // `.docx-default` class emits. Without this layer, a doc whose Normal sets
    // line=480 (double spacing) measured every plain paragraph at the bare grid
    // pitch and pages packed ~2× too dense (overflowed the page box).
    if (!styleId) {
      const def = t.paragraphStyles?.find((p) => p.default);
      if (def?.paragraph?.spacing?.line != null) return def.paragraph.spacing;
    }
    if (t.default?.document?.paragraph?.spacing?.line != null)
      return t.default.document.paragraph.spacing;
  }
  return null;
}

/** Paragraph default run properties (pPr/rPr) → run style baseline. Falls back
 *  to the paragraph's style and the document default run when attrs are absent,
 *  so the measured font matches the rendered one — a CJK doc default of 宋体
 *  measures wider than the generic "serif" fallback, which under-counted wrapped
 *  lines. */
function defaultRunOf(node: PmNode, styles?: unknown): Partial<RunStyle> {
  const run = (node.attrs as { run?: Record<string, unknown> | null }).run;
  const t = styleTableOf(styles);
  const styleId = (node.attrs as { styleId?: string | null }).styleId;
  const psRun = t && styleId ? t.paragraphStyles?.find((p) => p.id === styleId)?.run : null;
  // A paragraph with NO styleId implicitly uses the default paragraph style
  // (w:default="1", usually "Normal") — exactly what the renderer's `.docx-default`
  // class emits. Without this layer, a doc whose Normal sets run size=14 rendered
  // at 14pt but measured at the bare 12pt default; glyphs measured ~14/12 too
  // narrow → fewer wrapped lines → pages overflowed. (Same cascade gap resolveSpacing
  // already closes for spacing.)
  const defPsRun = !styleId ? t?.paragraphStyles?.find((p) => p.default)?.run : null;
  const defRun = t?.default?.document?.run;
  return {
    size:
      (run?.size as number | undefined) ?? psRun?.size ?? defPsRun?.size ?? defRun?.size ?? null,
    font: run?.font ?? psRun?.font ?? defPsRun?.font ?? defRun?.font ?? null,
    bold:
      (run?.bold as boolean | undefined) ?? psRun?.bold ?? defPsRun?.bold ?? defRun?.bold ?? false,
    italic:
      (run?.italic as boolean | undefined) ??
      psRun?.italic ??
      defPsRun?.italic ??
      defRun?.italic ??
      false,
    characterSpacing: (run?.characterSpacing as number | undefined) ?? null,
  };
}

interface FontSpec {
  size: number; // points
  font: unknown;
  bold: boolean;
  italic: boolean;
}

function buildFont(spec: FontSpec): string {
  const parts: string[] = [];
  if (spec.italic) parts.push("italic");
  if (spec.bold) parts.push("bold");
  parts.push(`${spec.size}pt`);
  parts.push(resolveFamily(spec.font) ?? DEFAULT_FAMILY);
  return parts.join(" ");
}

// ── line-height: canvas font-bounding-box = the browser's "normal" ──
// getComputedStyle().lineHeight returns "normal" (not a number), and approximating
// it as fontSize × 1.2 drifts from the real font metrics. canvas measureText's
// fontBoundingBoxAscent/Descent are exactly the metrics the browser uses for
// "normal", so the measured line-height matches the rendered one — deterministic
// and accurate.
const measureCanvas =
  typeof document !== "undefined" ? document.createElement("canvas").getContext("2d") : null;
const normalLineHeightCache = new Map<string, number>();

function normalLineHeightPx(font: string): number {
  const cached = normalLineHeightCache.get(font);
  if (cached != null) return cached;
  let lh = 0;
  if (measureCanvas) {
    measureCanvas.font = font;
    const m = measureCanvas.measureText("Mg");
    const ascent = (m as TextMetrics).fontBoundingBoxAscent ?? 0;
    const descent = (m as TextMetrics).fontBoundingBoxDescent ?? 0;
    lh = ascent + descent;
  }
  if (lh <= 0) {
    const sizeMatch = font.match(/(\d+(?:\.\d+)?)pt/);
    lh = sizeMatch ? parseFloat(sizeMatch[1]) * PT_TO_PX * 1.2 : 0;
  }
  normalLineHeightCache.set(font, lh);
  return lh;
}

/**
 * Resolve a paragraph's line-height in px from its OOXML spacing. `exact`/
 * `atLeast` → fixed twips→px; `auto`/undefined lineRule → multiple of the
 * font's normal box (line is in 240ths: 240 = single); no spacing.line → the
 * font's normal box. `font` is the canvas font shorthand for the paragraph's
 * dominant run (drives the normal-box measurement).
 *
 * `linePitchPx` applies the section's document grid (w:docGrid): when set,
 * every line snaps UP to that pitch — a normal or loosely-spaced line smaller
 * than the grid grows to it (Chinese docs default linePitch 312tw ≈ 20.8px, so
 * a 12pt normal line renders taller than its font box). `exact`/`atLeast`
 * spacing is an absolute override and ignores the grid, matching Word.
 */
/** Parse the point size out of a canvas font shorthand ("12pt serif" → 12pt →
 *  16px), for CSS-unitless line-height semantics (line-height: m = m × fontSize). */
function fontSizePxOf(font: string): number {
  const m = font.match(/(\d+(?:\.\d+)?)pt/);
  return m ? parseFloat(m[1]) * PT_TO_PX : DEFAULT_SIZE_PT * PT_TO_PX;
}

export function resolveLineHeight(
  spacing: { line?: number | null; lineRule?: string | null } | null | undefined,
  font: string,
  linePitchPx?: number,
): number {
  let lh: number;
  let isExact = false;
  if (spacing?.line) {
    const rule = spacing.lineRule;
    if (rule === "exact" || rule === "exactly" || rule === "atLeast") {
      lh = spacing.line * TWIP_TO_PX;
      isExact = true;
    } else {
      // "auto": a multiple of single spacing. With a document grid (Word
      // w:docGrid), single = the grid pitch (lines snap to it) — matching the
      // renderer's calc(var(--docen-line-pitch) * m). Without a grid, single =
      // the font size (CSS unitless line-height = m × fontSize). The previous
      // (single = font bounding box) over-estimated Word's single line and
      // drifted from the rendered height.
      const single = linePitchPx ?? fontSizePxOf(font);
      lh = (spacing.line / 240) * single;
    }
  } else {
    // No explicit line spacing: single line = the grid pitch (when active, the
    // page renders it via inherited line-height) else the font's normal box.
    lh = linePitchPx ?? normalLineHeightPx(font);
  }
  if (linePitchPx && !isExact && lh < linePitchPx) return linePitchPx;
  return lh;
}

// ── paragraph measurement ──

function collectInlineItems(para: PmNode, def: Partial<RunStyle>): RichInlineItem[] {
  const items: RichInlineItem[] = [];
  const fallbackSpec = (): FontSpec => ({
    size: def.size ?? DEFAULT_SIZE_PT,
    font: def.font,
    bold: def.bold ?? false,
    italic: def.italic ?? false,
  });
  para.forEach((child) => {
    if (child.isText) {
      const ms = markStyleOf(child.marks as readonly MarkLike[] | undefined);
      const spec: FontSpec = {
        size: ms.size ?? def.size ?? DEFAULT_SIZE_PT,
        font: ms.font ?? def.font,
        bold: ms.bold ?? def.bold ?? false,
        italic: ms.italic ?? def.italic ?? false,
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
  let hasContent = false;
  node.forEach((child) => {
    if (child.isText || child.type.name === "hardBreak" || child.type.name === "image") {
      hasContent = true;
    }
  });
  return !hasContent;
}

/** Height of a paragraph's strut line — the line-box minimum when there is no
 *  text (an empty paragraph, or an image row shorter than a text line). Mirrors
 *  renderParagraphStyles: spacing.line wins; else the paragraph-mark run size
 *  (pPr/rPr.sz) renders as `line-height:${size}pt` — an ABSOLUTE value, so
 *  markSize × PT_TO_PX (NOT normalLineHeightPx, whose font-metric box
 *  over-estimates the absolute pt line-height the browser actually produces);
 *  else the section grid pitch or the font's normal box. */
function emptyLineHeight(
  node: PmNode,
  styles: unknown,
  linePitchPx: number | undefined,
  font: string,
): number {
  const spacing = resolveSpacing(node, styles);
  if (spacing?.line) return resolveLineHeight(spacing, font, linePitchPx);
  const markSize = (node.attrs as { run?: { size?: number | null } | null }).run?.size ?? null;
  if (markSize != null) return markSize * PT_TO_PX;
  return linePitchPx ?? normalLineHeightPx(font);
}

// CJK kinsoku (line-start/end prohibition) character classes. Pretext's
// Caveats state it supports only `line-break: auto` (no kinsoku strictness),
// but browsers apply kinsoku to CJK by default: an opening bracket at line end
// (（《「" is pushed to the next line, and a closing punct at line start
// (。，)》 is pulled back to the previous line — both add a line vs Pretext's
// pure width breaking. Paragraphs with 《》"" therefore measured ~1 line short,
// the paginator over-packed the page, and content overflowed the page box. A
// deterministic kinsoku post-pass over Pretext's width-accurate breaks aligns
// with the browser (same input → same output, so convergence is preserved).
// Forbidden at line end (opening brackets / opening quotes):
const _PROHIBIT_TRAILING_RE = /[([{<（［｛〈《「『【〔〖〘〚“‘‟‛]/;
// Forbidden at line start (closing brackets + 、。，．；：！？… + closing quotes):
const _PROHIBIT_LEADING_RE = /[)\]}>.,;:!?）］｝〉》」』】〗〙〛、。，．；：！？…‥·”’]/;

/** Lay out a prepared paragraph line by line: the FIRST line at
 *  `width − firstLinePx` and later lines at `width` — mirroring CSS text-indent
 *  (only the first line is indented). Returns the line count and the PM content
 *  offset at the END of each line (a legal split point). Pretext's walk/stats
 *  take a single width for all lines and can't express a narrower first line,
 *  which under-counted the first line's capacity; layoutNextRichInlineLineRange
 *  lets us shrink only line 1. Offset mapping: each Pretext item is 1:1 with a
 *  PM inline child (text: 1 char = 1 offset; hardBreak: "\n" = 1 offset), so
 *  accumulating fragment text lengths == accumulating PM content offsets.
 *
 *  `oneCharWidthPx` (1-char width of the dominant font size) peeks the next
 *  line's first grapheme for leading kinsoku. */
function layoutLineOffsets(
  prepared: PreparedRichInline,
  width: number,
  firstLinePx: number,
  _oneCharWidthPx: number,
): { lineCount: number; lineBreakOffsets: number[] } {
  const lineBreakOffsets: number[] = [];
  let pmOff = 0;
  let lineCount = 0;
  let cursor: RichInlineCursor | undefined;
  let first = true;
  while (true) {
    const mw = first ? Math.max(0, width - firstLinePx) : width;
    let range = layoutNextRichInlineLineRange(prepared, mw, cursor);
    if (!range) break;
    let line = materializeRichInlineLineRange(prepared, range);
    let text = "";
    let chars = 0;
    for (const frag of line.fragments) {
      text += frag.text;
      chars += frag.text.length;
    }
    // Trailing kinsoku: the line ends on an opening bracket/quote (forbidden at
    // line end) → shrink the wrap width and re-break until the end char is
    // legal, pushing the opening punct to the next line (mirrors the browser).
    // Shrinking the width avoids retreating across a segment boundary — CJK
    // runs are often one grapheme per segment, so graphemeIndex is 0 at line
    // ends and a cursor retreat can't move back within the segment.
    const endCursor: RichInlineCursor = range.end;
    // Leading kinsoku: the next line would start with a closing punct (forbidden
    // at line start) → pull it into this line (the browser moves a line-start
    // closing punct back to the previous line's end). Peek the next line's
    // first grapheme with a 1-char width.
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
  const items = collectInlineItems(node, def);
  const prepared = getPrepared(items);
  const font =
    items[0]?.font ??
    buildFont({
      size: def.size ?? DEFAULT_SIZE_PT,
      font: def.font,
      bold: false,
      italic: false,
    });
  const firstLinePx = resolveFirstLineIndentPx(node, ctx?.styles);
  const strut = emptyLineHeight(node, ctx?.styles, ctx?.linePitchPx, font);
  // Empty paragraph (no text/image children): its sole content is the ¶ glyph,
  // whose height is the paragraph-mark line height (strut) — NOT a text line.
  if (isEmptyTextblock(node)) return strut;
  // Image-only paragraph: lay out images by content width (mirrors renderHTML
  // — each inline image occupies attrs.width × attrs.height px; multiple images
  // share a line until their combined width exceeds the content width, then
  // wrap). Each row's height is its tallest image but never below the strut (a
  // row of small images still occupies a full text line — the line-box minimum
  // from font-size/line-height). Total = sum of per-row heights. Previously
  // each image was credited its own line, over-counting side-by-side images by
  // multiples — the paginator packed them on one page and clipped the rest.
  if (hasInlineImage(node)) {
    const layout = layoutImageLines(node, width, strut);
    if (layout) return layout.lineHeights.reduce((s, h) => s + h, 0);
    // Mixed text+image paragraph: conservative upper bound — text line height
    // plus each image at least the strut (mixed paragraphs are not split; they
    // move whole — over-estimate rather than overflow). Precise wrapping needs
    // a unified text+image line breaker (Pretext's flow holds no image atom);
    // out of scope here.
    const lh = resolveLineHeight(resolveSpacing(node, ctx?.styles), font, ctx?.linePitchPx);
    let imgH = 0;
    node.forEach((child) => {
      if (child.type.name === "image") {
        const h = (child.attrs as { height?: number }).height;
        imgH += typeof h === "number" && h > 0 ? Math.max(h, strut) : strut;
      }
    });
    const lineCount = layoutLineOffsets(prepared, width, firstLinePx, fontSizePxOf(font)).lineCount;
    return Math.max(1, lineCount) * lh + imgH;
  }
  const lh = resolveLineHeight(resolveSpacing(node, ctx?.styles), font, ctx?.linePitchPx);
  const lineCount = layoutLineOffsets(prepared, width, firstLinePx, fontSizePxOf(font)).lineCount;
  return Math.max(1, lineCount) * lh;
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
    const font = buildFont({
      size: def.size ?? DEFAULT_SIZE_PT,
      font: def.font,
      bold: false,
      italic: false,
    });
    const strut = emptyLineHeight(node, ctx?.styles, ctx?.linePitchPx, font);
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
  const font =
    items[0]?.font ??
    buildFont({ size: def.size ?? DEFAULT_SIZE_PT, font: def.font, bold: false, italic: false });
  const lineHeight = resolveLineHeight(resolveSpacing(node, ctx?.styles), font, ctx?.linePitchPx);
  const firstLinePx = resolveFirstLineIndentPx(node, ctx?.styles);
  const { lineCount, lineBreakOffsets } = layoutLineOffsets(
    prepared,
    width,
    firstLinePx,
    fontSizePxOf(font),
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
      const iw = typeof w === "number" && w > 0 ? w : 0;
      const ih = typeof h === "number" && h > 0 ? h : 0;
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
  return {
    keepLines: a.keepLines === true,
    keepNext: a.keepNext === true,
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
  const t = styleTableOf(styles);
  if (t) {
    const styleId = (node.attrs as { styleId?: string | null }).styleId;
    const ps = styleId ? t.paragraphStyles?.find((p) => p.id === styleId) : null;
    if (ps?.paragraph?.indent) return ps.paragraph.indent as IndentAttrs;
    if (t.default?.document?.paragraph?.indent)
      return t.default.document.paragraph.indent as IndentAttrs;
  }
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
  const w = (table.attrs as { width?: { size: number | string; type?: string } | null }).width;
  if (!w) return pageContentWidth;
  const num = typeof w.size === "string" ? parseFloat(w.size) : w.size;
  if (w.type === "pct") {
    const pct = typeof w.size === "string" && w.size.includes("%") ? parseFloat(w.size) : num / 50;
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
 *  that cell's column width(s). rowspan cells count fully on their start row
 *  (Word distributes across spans; this over-estimates the start row and
 *  under-estimates the spanned rows, but the measurement is deterministic — no
 *  DOM wobble between re-flows). Cell margins/padding are not yet added. */
export function measureRowHeight(
  row: PmNode,
  columnWidths: number[],
  ctx?: MeasureContext,
): number {
  let maxHeight = 0;
  let colCursor = 0;
  row.forEach((cell) => {
    const colspan = (cell.attrs as { colspan?: number }).colspan ?? 1;
    const cellWidth = columnWidths.slice(colCursor, colCursor + colspan).reduce((a, b) => a + b, 0);
    colCursor += colspan;
    let contentH = 0;
    cell.forEach((block) => {
      contentH += measureBlockHeight(block, Math.max(0, cellWidth), ctx);
    });
    maxHeight = Math.max(maxHeight, contentH);
  });
  return maxHeight;
}
