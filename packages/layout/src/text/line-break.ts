// The line packer — pretext's rich-inline breaker as the kernel, with the
// OOXML semantics pretext doesn't own layered on top:
// - First-line indent shrinks only line 0 (CSS text-indent): the line query
//   simply runs at width − indent.
// - Float zones reduce each line's usable width at its top Y.
// - Tabs are JUMPS and hard breaks are FORCED ends: both are OOXML atoms
//   outside pretext's model, so the inline flow splits at tab AND break
//   atoms into groups; within a group pretext breaks (browser-grade UAX #14,
//   whitespace collapse, kinsoku), at a tab boundary the packer computes the
//   tab's advance (stop alignment, right-stop lookahead) and continues the
//   SAME line with the next group, at a hard break it ends the line.
// - Pictures ride the flow as unbreakable atoms of known width (empty text +
//   extraWidth — the docen-local pretext patch keeps such items alive).
// - Line heights come from the caller's resolver (the paragraph module owns
//   OOXML line-height semantics); the packer floors picture-only lines at
//   their tallest picture and at the paragraph strut.

import {
  layoutNextRichInlineLineRange,
  materializeRichInlineLineRange,
  measureRichInlineStats,
  prepareRichInline,
  type PreparedRichInline,
  type RichInlineCursor,
  type RichInlineItem,
} from "@docen/pretext/rich-inline";

import type { LayoutFloatZone, LayoutInline, LayoutTabStop } from "../layout-doc";
import type { LaidOutLine, LaidOutLineItem } from "../layout-result";
import { cssFontOf, familyOfSlot, type TextMeasurer } from "./measure";

export interface LineHeightInput {
  /** Max run natural height among the line's text content (0 when none). */
  naturalPx: number;
  /** Whether any of the line's text itemized as CJK (docGrid ceil snap). */
  hasCjk: boolean;
  /** A picture floored this line (its box is the natural): the grid resolver
   *  ceil-snaps it to whole rows like a CJK line's. */
  hasPicture?: boolean;
}

export type LineHeightResolver = (line: LineHeightInput) => number;

export interface PackedLine {
  /** Positioned items; x is relative to the line's content start (the
   *  renderer applies the line origin: indent + zone reduction). Text items
   *  carry the run's WHITESPACE-COLLAPSED slice for this line — the string
   *  the renderer paints, not a slice of the source run. */
  items: LaidOutLineItem[];
  /** The source inline the line's content ends at (the last fragment's
   *  origin); a coarse split-point marker for page breaking. */
  endInlineIndex: number;
  /** The width this line packed against (after indent/zone reductions). */
  maxWidthPx: number;
  /** Resolved height (resolver, floored by pictures on the line and — when
   *  the line has no text — by the paragraph strut). */
  heightPx: number;
  /** The resolver's `naturalPx` input (max text natural height, 0 when the
   *  line has no text) — mirrored for the painter's half-leading math. */
  naturalPx: number;
  /** The line's largest run font size (undefined on textless lines) — the
   *  EM-box reference Word centers in a grid span (smaller than the browser
   *  font box the natural height measures). */
  textEmPx?: number;
  /** The advance of the closing punctuation hanging past this line's right
   *  edge (w:overflowPunct) — 0/undefined when the line ends flush. The hang
   *  never counts against justification or center/right slack. */
  hangPx?: number;
  /** A picture floored this line's height (the resolver's height was smaller)
   *  — a grid line ceils to whole rows and centers the picture box itself. */
  pictureFloored?: boolean;
  /** The uniform advance squeeze this line's pure-CJK run was compressed by
   *  (Word's compressPunctuation, < 1; undefined = natural advances) — the
   *  item x/width are already scaled; the renderer and caret map compress
   *  glyph advances to the same factor. */
  advanceScale?: number;
  /** How far the line's content start sits right of the paragraph's text box
   *  edge — set when a wrapSide right/largest float takes the left side and
   *  the text packs past its right edge (the renderer shifts the line). */
  xOffsetPx?: number;
}

export interface PackLinesOptions {
  measurer: TextMeasurer;
  /** Content width in px (block indents already subtracted by the caller). */
  width: number;
  /** Shrinks line 0 only (first-line indent). */
  firstLineIndentPx?: number;
  /** OOXML line-height semantics, supplied by the paragraph module. */
  lineHeight: LineHeightResolver;
  /** The paragraph-mark strut (px): floors lines with no text content. */
  strutPx?: number;
  /** Paragraph top Y in flow coordinates — pairs with floatZones. */
  startY?: number;
  floatZones?: readonly LayoutFloatZone[];
  /** The paragraph's own wrapped drawings as zones in PARAGRAPH-relative Y
   *  (top 0 = the first line's top): the anchor paragraph wraps beside its
   *  own floats, while `floatZones` (absolute) cover every other paragraph. */
  selfZones?: readonly LayoutFloatZone[];
  /** Explicit tab stops, px from the content-box left edge (w:tabs). Tabs past
   *  the last stop (or with no stops) advance over the default grid — 720
   *  twips, Word's defaultTabStop. */
  tabStops?: LayoutTabStop[];
}

/** Word's defaultTabStop: 720 twips = 0.5 inch = 48 px at 96 dpi. */
const DEFAULT_TAB_PX = 48;

/** Fullwidth CJK closing punctuation (the kinsoku "cannot start a line"
 *  class). Word's w:overflowPunct — on by default — lets a trailing run of
 *  these hang past the line's right edge instead of pushing the break back,
 *  verified against the reference PDF: a line ending 、reports x1end one
 *  full advance beyond the right margin. */
const CLOSING_PUNCT = new Set([
  "、",
  "。",
  "，",
  "．",
  "：",
  "；",
  "？",
  "！",
  "）",
  "〉",
  "》",
  "」",
  "』",
  "】",
  "〕",
  "〗",
  "〙",
  "〛",
  "｝",
  "…",
  "”",
  "’",
]);

/** CJK text (ideographs, kana, fullwidth forms, CJK punctuation) — the only
 *  content the advance squeeze below may touch; Latin runs keep their
 *  natural metrics. */
const CJK_TEXT = /^[　-〿㐀-䶿一-鿿぀-ヿ＀-￯‘’“”…]+$/;

/** Word's w:characterSpacingControl compressPunctuation (its CJK default):
 *  a line whose next pushed run is pure CJK and misses fitting by a hair
 *  squeezes the line's advances to fit instead of breaking (corpus-verified:
 *  the honor table's 17-char cell — natural 272px in a 269.2px cell —
 *  renders ONE line at ~98.9% advance, not two). Bounded at 4% so a genuinely
 *  overflowing line still wraps; the residual squeeze is sub-glyph. */
const SQUEEZE_MAX = 0.04;

/** Prepared flows cache — prepare is the expensive pass (segment + measure);
 *  line queries are pure arithmetic. Keyed by the full item sequence. */
const preparedCache = new Map<string, PreparedRichInline>();
const PREPARED_CACHE_LIMIT = 4000;

/** One tab/break-delimited stretch of the inline flow. */
interface FlowGroup {
  /** RichInlineItem[i] originates from inline[itemInline[i]]. */
  itemInline: number[];
  items: RichInlineItem[];
  prepared: PreparedRichInline;
  /** The tab atom closing this group (null unless the group ends at a tab). */
  tab: Extract<LayoutInline, { kind: "tab" }> | null;
  /** Whether a hard-break atom closes this group (it ends the line). */
  hardBreak: boolean;
  /** The closing atom's inline index — the split-point marker a consumed
   *  closer leaves behind (endInlineIndex). */
  closerIndex: number;
  /** Natural (unwrapped) width of everything after this group's tab up to
   *  the next tab — the lookahead a RIGHT stop positions against. */
  followingPx: number;
}

function groupOf(
  inline: LayoutInline[],
  from: number,
  to: number,
  closer: { tab: FlowGroup["tab"]; hardBreak: boolean; closerIndex: number },
  measurer: TextMeasurer,
): FlowGroup {
  const items: RichInlineItem[] = [];
  const itemInline: number[] = [];
  for (let i = from; i < to; i++) {
    const item = inline[i];
    if (item.kind === "text") {
      // Script itemization: each same-script segment measures (and paints)
      // with its slot's family — the OOXML eastAsia/ascii split.
      const { segments } = measurer.analyze(item.text, item.style);
      for (const seg of segments) {
        items.push({
          text: seg.text,
          font: cssFontOf(item.style, familyOfSlot(item.style.family, seg.isCjk)),
          letterSpacing: item.style.letterSpacingPx,
        });
        itemInline.push(i);
      }
    } else if (item.kind === "picture" || item.kind === "math") {
      // An unbreakable atom of known width (the patched pretext keeps
      // empty-text extraWidth items alive).
      items.push({ text: "", font: "1px serif", break: "never", extraWidth: item.widthPx });
      itemInline.push(i);
    }
    // Tab and break atoms carry no item — the group boundary itself jumps
    // (tab) or ends the line (break).
  }
  const key = items
    .map(
      (it) =>
        `${it.text}\x00${it.font}\x00${it.letterSpacing ?? ""}\x00${it.break ?? ""}\x00${it.extraWidth ?? ""}`,
    )
    .join("\x01");
  let prepared = preparedCache.get(key);
  if (!prepared) {
    prepared = prepareRichInline(items);
    if (preparedCache.size >= PREPARED_CACHE_LIMIT) preparedCache.clear();
    preparedCache.set(key, prepared);
  }
  return { itemInline, items, prepared, ...closer, followingPx: 0 };
}

/** Split the inline flow at tab and break atoms into pretext-prepared
 *  groups; the final group (no closing atom) runs to the flow's end. */
function buildGroups(inline: LayoutInline[], measurer: TextMeasurer): FlowGroup[] {
  const groups: FlowGroup[] = [];
  let start = 0;
  for (let i = 0; i < inline.length; i++) {
    if (inline[i].kind === "tab" || inline[i].kind === "break") {
      groups.push(
        groupOf(
          inline,
          start,
          i,
          inline[i].kind === "tab"
            ? {
                tab: inline[i] as Extract<LayoutInline, { kind: "tab" }>,
                hardBreak: false,
                closerIndex: i,
              }
            : { tab: null, hardBreak: true, closerIndex: i },
          measurer,
        ),
      );
      start = i + 1;
    }
  }
  groups.push(
    groupOf(
      inline,
      start,
      inline.length,
      { tab: null, hardBreak: false, closerIndex: inline.length - 1 },
      measurer,
    ),
  );
  // Right-stop lookahead: each tabbed group's following width is the natural
  // (unwrapped) width of the groups after it, up to the next tab.
  for (let g = 0; g < groups.length; g++) {
    if (!groups[g].tab) continue;
    let following = 0;
    for (let h = g + 1; h < groups.length; h++) {
      following += measureRichInlineStats(groups[h].prepared, 1e9).maxLineWidth;
      if (groups[h].tab) break;
    }
    groups[g].followingPx = following;
  }
  return groups;
}

/** The widest line an inline flow produces when wrapped at `maxWidth` —
 *  wrappable prose packs within the budget, an unbreakable run wider than it
 *  overflows at its own length. That makes this the column floor Word's
 *  autofit grows a grid column past: prose stays at the grid width, a long
 *  word demands its own length (the same pretext pipeline that packs the
 *  real lines, so the fit and the wrap can never drift apart). */
export function minWidthOfInline(
  inline: LayoutInline[],
  measurer: TextMeasurer,
  maxWidth: number,
): number {
  let widest = 0;
  for (const group of buildGroups(inline, measurer)) {
    const w = measureRichInlineStats(group.prepared, maxWidth).maxLineWidth;
    if (w > widest) widest = w;
  }
  return widest;
}

/** The x span of a closed polygon's horizontal slice at `y` (even-odd rule,
 *  points in zone coordinates); undefined when the line misses the polygon. */
function contourSpanAt(
  pts: readonly { x: number; y: number }[],
  y: number,
): { min: number; max: number } | undefined {
  const xs: number[] = [];
  for (let i = 0; i < pts.length; i++) {
    const a = pts[i]!;
    const b = pts[(i + 1) % pts.length]!;
    if ((a.y <= y && b.y > y) || (b.y <= y && a.y > y)) {
      xs.push(a.x + ((y - a.y) / (b.y - a.y)) * (b.x - a.x));
    }
  }
  if (xs.length < 2) return undefined;
  xs.sort((p, q) => p - q);
  return { min: xs[0]!, max: xs[xs.length - 1]! };
}

/** The width — and start offset — available to line `lineIndex` whose box
 *  spans `[y, y+bottom)` (flow Y, for the absolute zones) / `[yPara,
 *  yPara+bottom)` (paragraph-relative Y, for the anchor's own zones). Zones
 *  are intervals in the line's x space: a `textAfter` zone shifts the line
 *  start past its far edge (text right of the float); every other zone caps
 *  the line's right edge at its near edge (text left of the float — a flat
 *  width reduction would run text under a float off the left margin). The
 *  real line height isn't known until the line packs, so the box bottom uses
 *  the resolver's floor — a zone top inside `[y, y+floor)` always overlaps
 *  the real line box, so the check never over-reduces. */
function maxWidthAt(
  opts: PackLinesOptions,
  lineIndex: number,
  y: number,
  yPara: number,
  bottomPx: number,
): { widthPx: number; xOffsetPx: number } {
  let w = opts.width;
  // The first-line indent shifts the line's glyphs right, so it comes OFF the
  // capped width (Word: line 0 packs into min(width, cap) − indent), not off
  // `width` before the cap — a tight float cap would otherwise swallow the
  // indent and line 0 would pack characters Word breaks to the next line.
  const firstLine = lineIndex === 0 ? (opts.firstLineIndentPx ?? 0) : 0;
  let cap = Infinity;
  let xOffset = 0;
  const scan = (zones: readonly LayoutFloatZone[] | undefined, zoneY: number): void => {
    if (!zones || zones.length === 0) return;
    for (const z of zones) {
      if (!(z.bottomPx > zoneY && z.topPx < zoneY + bottomPx)) continue;
      if (z.contour) {
        // A tight/through contour: slice the polygon at the line's mid-height
        // — the slice's bounding span is the forbidden interval (a concave
        // slice's gaps fill in), the text takes the wider side of it.
        const span = contourSpanAt(z.contour, zoneY + bottomPx / 2 - z.topPx);
        if (!span) continue;
        const relX0 = z.x0Px ?? 0;
        const leftRoom = relX0 + span.min;
        const rightRoom = opts.width - relX0 - span.max;
        if (rightRoom > leftRoom) {
          const start = relX0 + span.max;
          if (start > xOffset) xOffset = start;
        } else if (relX0 + span.min < cap) {
          cap = relX0 + span.min;
        }
        continue;
      }
      if (z.textAfter) {
        // Text packs right of the box: the line starts past its far edge.
        const start = (z.x0Px ?? 0) + z.widthPx;
        if (start > xOffset) xOffset = start;
      } else if ((z.x0Px ?? 0) < cap) {
        // Text packs left of the box: the line ends at its near edge.
        cap = z.x0Px ?? 0;
      }
    }
  };
  if (opts.startY != null) scan(opts.floatZones, y);
  scan(opts.selfZones, yPara);
  return {
    widthPx: Math.max(0, Math.min(w, cap) - xOffset - firstLine),
    xOffsetPx: xOffset,
  };
}

/** Tab geometry: the explicit stops and this line's content origin — x=0 of
 *  the walk sits `lineBasePx` into the paragraph's text box (the first line's
 *  indent), and stops are measured from the text box. */
interface TabContext {
  stops?: LayoutTabStop[];
  lineBasePx: number;
  /** This line's usable width: a stop past it clamps to the right edge
   *  (Word: an out-of-margin stop still right-aligns at the margin). */
  maxWidthPx: number;
}

/** The next stop position past `absX`: the first explicit stop beyond it, else
 *  the next slot of the 720-twip default grid. */
function nextStopPast(tabs: TabContext, absX: number): number {
  let best: number | null = null;
  for (const s of tabs.stops ?? []) {
    if (s.positionPx > absX + 0.01 && (best == null || s.positionPx < best)) best = s.positionPx;
  }
  if (best != null) return best;
  return (Math.floor(absX / DEFAULT_TAB_PX) + 1) * DEFAULT_TAB_PX;
}

/** The advance a tab atom produces at walk position `xRaw`: to its explicit
 *  target (a numbering bullet's hop), the next stop, or the default grid. A
 *  RIGHT stop aligns the FOLLOWING text's right edge at the stop, so the tab
 *  itself yields by `followingPx`; when that would cross behind the cursor the
 *  tab degenerates to the default-grid hop (progress, never negative). The
 *  matched explicit stop's leader rides along for the painter; a default-grid
 *  hop (no stop) carries none. */
function tabAdvance(
  tab: { toPx?: number },
  xRaw: number,
  tabs: TabContext,
  followingPx: number,
): { advancePx: number; leader?: LayoutTabStop["leader"] } {
  const absX = xRaw + tabs.lineBasePx;
  if (tab.toPx != null) return { advancePx: Math.max(0, tab.toPx - absX) };
  // A stop past this line's usable width clamps to the right edge (Word: an
  // out-of-margin stop still right-aligns at the margin). maxWidth is measured
  // from the line's content start; stops from the text box — convert, clamp,
  // convert back.
  const stop =
    Math.min(nextStopPast(tabs, absX) - tabs.lineBasePx, tabs.maxWidthPx) + tabs.lineBasePx;
  const matched = tabs.stops?.find((s) => s.positionPx === stop);
  const type = matched?.type ?? "left";
  const w =
    type === "right"
      ? stop - absX - followingPx
      : type === "center"
        ? stop - absX - followingPx / 2
        : stop - absX;
  // A right/center stop the following text cannot reach falls back to the
  // default grid's next slot (the text starts past the stop — progress).
  const advancePx =
    w > 0 ? w : Math.max(0, (Math.floor(absX / DEFAULT_TAB_PX) + 1) * DEFAULT_TAB_PX - absX);
  return { advancePx, leader: matched?.leader };
}

/** The whitespace pretext collapses into the gaps between a line's text items
 *  (CSS white-space semantics; Word keeps the characters). Both consumers of
 *  the gaps — the painter's space dots and the caret map's character lattice
 *  — must agree on where those characters are, so the walk lives here: each
 *  item consumes the whitespace ahead of it in the paragraph's concatenated
 *  run text, which marks every non-text inline's source position with one
 *  U+FFFC placeholder (an atom consumes its placeholder after its gap). When
 *  `fullText` carries no placeholders the atoms miss and `matched: false`
 *  degrades every count to 0 — callers fall back to the item texts alone. */
export function lineSpaceGaps(
  line: Pick<LaidOutLine, "items">,
  fullText: string,
  cursor: number,
): { spaces: number[]; next: number; matched: boolean } {
  const spaces = line.items.map(() => 0);
  let at = cursor;
  let matched = true;
  line.items.forEach((item, itemIndex) => {
    let p = at;
    while (p < fullText.length && (fullText[p] === " " || fullText[p] === "　")) p++;
    if (item.kind !== "text") {
      // An atom's gap: the whitespace ahead of it is trimmed from the laid
      // line just the same (the collapsed run between it and the previous
      // content), and its own placeholder is consumed so the items after it
      // stay aligned.
      if (fullText[p] === "￼") {
        spaces[itemIndex] = p - at;
        at = p + 1;
      } else {
        matched = false;
      }
      return;
    }
    if (fullText.startsWith(item.text, p)) {
      spaces[itemIndex] = p - at;
      at = p + item.text.length;
    } else {
      matched = false;
    }
  });
  return { spaces, next: at, matched };
}

/** One-off advance measurements for hanging closers, keyed by char+font. */
const closerAdvanceCache = new Map<string, number>();

/** The advance of a single grapheme in its run's font (pretext measures with
 *  canvas measureText; a one-item prepared line is the public handle). */
function advanceOfGrapheme(ch: string, font: string, letterSpacing?: number): number {
  const key = `${ch}\x00${font}\x00${letterSpacing ?? 0}`;
  let w = closerAdvanceCache.get(key);
  if (w === undefined) {
    const prepared = prepareRichInline([{ text: ch, font, letterSpacing }]);
    w = measureRichInlineStats(prepared, 1e9).maxLineWidth;
    closerAdvanceCache.set(key, w);
  }
  return w;
}

/** Word's w:overflowPunct probe: what a break at `range`'s end leaves for the
 *  next line, as `{ leadPx, closerPx, closer }` — the closer's advance plus
 *  the advance of the (kinsoku-pushed) glyphs that would precede it. The next
 *  line is queried at minimal width: forced progress yields its opening run
 *  (a closer that cannot start a line comes along, pulled, with whatever was
 *  pushed back before it). Returns undefined when the next line would not
 *  end in a hanging closer. */
function overflowPunctAfter(
  group: FlowGroup,
  range: { width: number; end: RichInlineCursor },
): { leadPx: number; closerPx: number } | undefined {
  const probe = layoutNextRichInlineLineRange(group.prepared, 1, range.end);
  if (!probe) return undefined;
  const fragments = materializeRichInlineLineRange(group.prepared, probe).fragments;
  const text = fragments.map((f) => f.text).join("");
  // CLOSING_PUNCT is all-BMP, so the last UTF-16 unit is the whole closer;
  // an astral tail cannot be in the set and falls to the undefined path.
  const closer = text.at(-1);
  if (closer == null || !CLOSING_PUNCT.has(closer)) return undefined;
  const lastFrag = fragments[fragments.length - 1];
  const src = group.items[lastFrag.itemIndex];
  if (!src) return undefined;
  const closerPx = advanceOfGrapheme(closer, src.font, src.letterSpacing);
  return { leadPx: Math.max(0, probe.width - closerPx), closerPx };
}

/** Pack a paragraph's inline content into lines. Always returns at least one
 *  line when content exists; an empty inline array returns no lines (the
 *  paragraph module supplies the strut height for that case). */
export function packLines(inline: LayoutInline[], opts: PackLinesOptions): PackedLine[] {
  const { measurer } = opts;
  if (inline.length === 0) return [];
  const groups = buildGroups(inline, measurer);
  const cursors: (RichInlineCursor | undefined)[] = groups.map(() => undefined);
  const done = groups.map(() => false);

  const lines: PackedLine[] = [];
  let lineIndex = 0;
  let y = opts.startY ?? 0;
  let yPara = 0;
  // Zone-overlap floor: every line box is at least this tall (the resolver's
  // base, floored by the strut), so zone tops within reach count as overlaps.
  const zoneFloorPx = Math.max(opts.lineHeight({ naturalPx: 0, hasCjk: false }), opts.strutPx ?? 0);

  while (done.some((d) => !d)) {
    const { widthPx: maxWidth, xOffsetPx } = maxWidthAt(opts, lineIndex, y, yPara, zoneFloorPx);
    const tabs: TabContext = {
      stops: opts.tabStops,
      lineBasePx: lineIndex === 0 ? (opts.firstLineIndentPx ?? 0) : 0,
      maxWidthPx: maxWidth,
    };

    const lineItems: LaidOutLineItem[] = [];
    let xLine = 0;
    let naturalPx = 0;
    let textEmPx: number | undefined;
    let hasCjk = false;
    let tallestPicturePx = 0;
    let hasText = false;
    // Whether any closing atom was placed this pass — an intentional blank
    // line always comes from one (a hard break); see the phantom guard below.
    let closedAtom = false;
    let endInlineIndex = 0;
    let brokeMidGroup = false;
    let hangPx = 0;
    // The uniform advance factor applied to this line's fragments (just under
    // 1, or undefined = natural advances).
    let squeeze: number | undefined;

    for (let g = 0; g < groups.length; g++) {
      const group = groups[g];
      // Whether the group finished ON THIS LINE — the only time its closing
      // atom is consumed (a group finished on an earlier line must not jump
      // its tab or end the line again).
      let finishedThisLine = false;
      if (!done[g]) {
        const query = Math.max(1, maxWidth - xLine);
        let range = layoutNextRichInlineLineRange(group.prepared, query, cursors[g]);
        // w:overflowPunct (Word default): when the glyphs the break pushes to
        // the next line END in a closing punctuation and the non-closer part
        // fits this line's slack, grant the closer's advance so it joins this
        // line and hangs past the right edge. Greedy fit is monotone in
        // width, so the re-query's break sits at or past the closer — but a
        // narrow glyph after it could still sneak in, so the hang is kept
        // only when the re-queried line really ends with a closer.
        if (range && range.end.itemIndex < group.items.length) {
          const hang = overflowPunctAfter(group, range);
          if (hang && hang.leadPx <= query - range.width + 0.01) {
            const re = layoutNextRichInlineLineRange(
              group.prepared,
              query + hang.closerPx,
              cursors[g],
            );
            if (re) {
              const reText = materializeRichInlineLineRange(group.prepared, re)
                .fragments.map((f) => f.text)
                .join("");
              // Last code point (for...of iterates code points).
              let last: string | undefined;
              for (const ch of reText) last = ch;
              if (last != null && CLOSING_PUNCT.has(last)) {
                range = re;
                hangPx = hang.closerPx;
              }
            }
          }
        }
        // The advance squeeze (see SQUEEZE_MAX): tried after the hang — a
        // hanging closer already fits its line, squeezing is moot there.
        if (range && range.end.itemIndex < group.items.length && !squeeze) {
          const probe = layoutNextRichInlineLineRange(group.prepared, 1e9, range.end);
          if (probe) {
            const probeFrags = materializeRichInlineLineRange(group.prepared, probe).fragments;
            const lineFrags = materializeRichInlineLineRange(group.prepared, range).fragments;
            const allCjk =
              probeFrags.every((f) => CJK_TEXT.test(f.text)) &&
              lineFrags.every((f) => CJK_TEXT.test(f.text));
            const queryNow = Math.max(1, maxWidth - xLine);
            const total = range.width + probe.width;
            const overflow = total - queryNow;
            if (allCjk && overflow > 0 && overflow <= total * SQUEEZE_MAX) {
              // Re-query at the run's natural width: the greedy break now
              // covers the probed run (or more). The materialized width under
              // the real squeeze below decides the factor.
              const re = layoutNextRichInlineLineRange(group.prepared, total, cursors[g]);
              if (re) {
                const width = materializeRichInlineLineRange(group.prepared, re).width;
                const k = queryNow / width;
                if (k < 1 && k >= 1 - SQUEEZE_MAX) {
                  range = re;
                  squeeze = k;
                }
              }
            }
          }
        }
        if (range) {
          const line = materializeRichInlineLineRange(group.prepared, range);
          const xStart = xLine;
          let xFrag = xLine;
          for (const frag of line.fragments) {
            const inlineIndex = group.itemInline[frag.itemIndex];
            const src = inline[inlineIndex];
            // The fragment's gap is its own leading advance (an inter-run
            // space — pretext trims boundary whitespace into gaps): it moves
            // the fragment's start before the fragment is placed, or the
            // space's width would vanish between two runs.
            xFrag += frag.gapBefore;
            // A squeezed line re-spaces its fragments' advances (k < 1) — the
            // glyphs keep their metrics, only the positions tighten (Word's
            // advance compression; the painter draws at the given x).
            const at = squeeze ? xStart + (xFrag - xStart) * squeeze : xFrag;
            if (src.kind === "text") {
              lineItems.push({
                kind: "text",
                inlineIndex,
                text: frag.text,
                xPx: at,
                widthPx: squeeze ? frag.occupiedWidth * squeeze : frag.occupiedWidth,
                // Synthetic markers stay flagged through materialization — the
                // editor's caret map reads it to keep them out of the PM
                // offset space.
                synthetic: src.synthetic,
              });
              hasText = true;
              const analyzed = measurer.analyze(src.text, src.style);
              if (analyzed.naturalPx > naturalPx) naturalPx = analyzed.naturalPx;
              if (textEmPx == null || src.style.sizePx > textEmPx) textEmPx = src.style.sizePx;
              if (analyzed.hasCjk) hasCjk = true;
            } else if (src.kind === "picture") {
              lineItems.push({
                kind: "picture",
                inlineIndex,
                xPx: at,
                widthPx: src.widthPx,
                heightPx: src.heightPx,
              });
              if (src.heightPx > tallestPicturePx) tallestPicturePx = src.heightPx;
            } else if (src.kind === "math") {
              lineItems.push({
                kind: "math",
                inlineIndex,
                xPx: at,
                widthPx: src.widthPx,
                heightPx: src.heightPx,
                label: src.label,
              });
              if (src.heightPx > tallestPicturePx) tallestPicturePx = src.heightPx;
            }
            xFrag += frag.occupiedWidth;
            endInlineIndex = inlineIndex;
          }
          // A squeezed group fills the line's remaining width exactly.
          xLine += squeeze ? maxWidth - xStart : line.width;
          cursors[g] = range.end;
          if (range.end.itemIndex >= group.items.length) {
            done[g] = true;
            finishedThisLine = true;
          } else {
            brokeMidGroup = true;
            break;
          }
        } else {
          done[g] = true;
          finishedThisLine = true;
        }
      }
      // Group exhausted on this line: consume its closing atom here — a tab
      // jumps and the line continues; a hard break ends the line.
      if (finishedThisLine && !brokeMidGroup) {
        if (group.tab) {
          const { advancePx, leader } = tabAdvance(group.tab, xLine, tabs, group.followingPx);
          // The tab's interval becomes an item so the painter can fill it with
          // the stop's leader (dot leaders in a TOC).
          lineItems.push({
            kind: "tab",
            inlineIndex: group.closerIndex,
            xPx: xLine,
            widthPx: advancePx,
            leader,
          });
          xLine += advancePx;
          endInlineIndex = group.closerIndex;
          closedAtom = true;
        } else if (group.hardBreak) {
          endInlineIndex = group.closerIndex;
          closedAtom = true;
          break;
        }
      }
    }

    // Floor the resolver: pictures size their line; a line with no text at
    // all bottoms at the paragraph strut (the ¶-mark line box).
    let height = opts.lineHeight({ naturalPx, hasCjk });
    if (!hasText && opts.strutPx != null && opts.strutPx > height) height = opts.strutPx;
    const pictureFloored = tallestPicturePx > height;
    if (tallestPicturePx > height) height = tallestPicturePx;
    // A picture is line content: its height is part of the line's natural
    // box. Without this, a docGrid line carrying only a picture reports a
    // text-sized natural, and the painter's grid centering sinks the picture
    // by half its height instead of pinning it to the line top.
    if (tallestPicturePx > naturalPx) naturalPx = tallestPicturePx;
    // A grid line's picture box spans whole rows too: re-resolve so the
    // ceil snap takes the picture's natural, and the painter's half-leading
    // (gridPadOf) centers the box in the span. Pixel-verified against the
    // reference renders — a 681px and a 639px picture on a 17pt pitch land
    // (span − box)/2 = 10.8/9.2px lower, matching the ±0.2px tolerance.
    if (pictureFloored) {
      height = Math.max(height, opts.lineHeight({ naturalPx, hasCjk, hasPicture: true }));
    }

    // An exhausted stream re-enters the packing loop once more before every
    // cursor reports done; that pass places nothing (the resolver emits an
    // empty range for the trailing run) and must not become a line — a
    // phantom blank row under every short centered title. A deliberate
    // blank line closes its own hard-break/tab atom first, so `closedAtom`
    // keeps those while suppressing only the artifact.
    if (lineItems.length === 0 && lines.length > 0 && !closedAtom) {
      lineIndex++;
      continue;
    }

    lines.push({
      items: lineItems,
      endInlineIndex,
      maxWidthPx: maxWidth,
      heightPx: height,
      naturalPx,
      textEmPx,
      ...(pictureFloored ? { pictureFloored: true } : {}),
      ...(squeeze != null ? { advanceScale: squeeze } : {}),
      hangPx: hangPx > 0 ? hangPx : undefined,
      xOffsetPx: xOffsetPx > 0 ? xOffsetPx : undefined,
    });
    y += height;
    yPara += height;
    lineIndex++;
  }
  return lines;
}
