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
} from "@chenglou/pretext/rich-inline";

import type { LayoutFloatZone, LayoutInline } from "../layout-doc";
import type { LaidOutLineItem } from "../layout-result";
import { cssFontOf, familyOfSlot, type TextMeasurer } from "./measure";

export interface LineHeightInput {
  /** Max run natural height among the line's text content (0 when none). */
  naturalPx: number;
  /** Whether any of the line's text itemized as CJK (docGrid ceil snap). */
  hasCjk: boolean;
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
  /** The advance of the closing punctuation hanging past this line's right
   *  edge (w:overflowPunct) — 0/undefined when the line ends flush. The hang
   *  never counts against justification or center/right slack. */
  hangPx?: number;
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
  /** Explicit tab stops, px from the content-box left edge (w:tabs). Tabs past
   *  the last stop (or with no stops) advance over the default grid — 720
   *  twips, Word's defaultTabStop. */
  tabStops?: { positionPx: number; type: "left" | "center" | "right" }[];
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
    } else if (item.kind === "picture") {
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

/** The width available to line `lineIndex` whose top sits at `y`: the content
 *  width (minus the first-line indent on line 0) reduced by the widest float
 *  zone overlapping the line's top (text wraps beside the float). */
function maxWidthAt(opts: PackLinesOptions, lineIndex: number, y: number): number {
  let w = opts.width;
  if (lineIndex === 0 && opts.firstLineIndentPx) w -= opts.firstLineIndentPx;
  if (opts.floatZones && opts.floatZones.length > 0 && opts.startY != null) {
    let reduce = 0;
    for (const z of opts.floatZones) {
      // Top-point check: a zone counts for the line it overlaps at its top.
      // A zone starting mid-line is picked up by the next line.
      if (z.bottomPx > y && z.topPx <= y) reduce = Math.max(reduce, z.widthPx);
    }
    w -= reduce;
  }
  return Math.max(0, w);
}

/** Tab geometry: the explicit stops and this line's content origin — x=0 of
 *  the walk sits `lineBasePx` into the paragraph's text box (the first line's
 *  indent), and stops are measured from the text box. */
interface TabContext {
  stops?: { positionPx: number; type: "left" | "center" | "right" }[];
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
 *  tab degenerates to the default-grid hop (progress, never negative). */
function tabAdvance(
  tab: { toPx?: number },
  xRaw: number,
  tabs: TabContext,
  followingPx: number,
): number {
  const absX = xRaw + tabs.lineBasePx;
  if (tab.toPx != null) return Math.max(0, tab.toPx - absX);
  // A stop past this line's usable width clamps to the right edge (Word: an
  // out-of-margin stop still right-aligns at the margin). maxWidth is measured
  // from the line's content start; stops from the text box — convert, clamp,
  // convert back.
  const stop =
    Math.min(nextStopPast(tabs, absX) - tabs.lineBasePx, tabs.maxWidthPx) + tabs.lineBasePx;
  const type = tabs.stops?.find((s) => s.positionPx === stop)?.type ?? "left";
  const w =
    type === "right"
      ? stop - absX - followingPx
      : type === "center"
        ? stop - absX - followingPx / 2
        : stop - absX;
  // A right/center stop the following text cannot reach falls back to the
  // default grid's next slot (the text starts past the stop — progress).
  return w > 0 ? w : Math.max(0, (Math.floor(absX / DEFAULT_TAB_PX) + 1) * DEFAULT_TAB_PX - absX);
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

  while (done.some((d) => !d)) {
    const maxWidth = maxWidthAt(opts, lineIndex, y);
    const tabs: TabContext = {
      stops: opts.tabStops,
      lineBasePx: lineIndex === 0 ? (opts.firstLineIndentPx ?? 0) : 0,
      maxWidthPx: maxWidth,
    };

    const lineItems: LaidOutLineItem[] = [];
    let xLine = 0;
    let naturalPx = 0;
    let hasCjk = false;
    let tallestPicturePx = 0;
    let hasText = false;
    let endInlineIndex = 0;
    let brokeMidGroup = false;
    let hangPx = 0;

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
        if (range) {
          const line = materializeRichInlineLineRange(group.prepared, range);
          let xFrag = xLine;
          for (const frag of line.fragments) {
            const inlineIndex = group.itemInline[frag.itemIndex];
            const src = inline[inlineIndex];
            if (src.kind === "text") {
              lineItems.push({
                kind: "text",
                inlineIndex,
                text: frag.text,
                xPx: xFrag,
                widthPx: frag.occupiedWidth,
              });
              hasText = true;
              const analyzed = measurer.analyze(src.text, src.style);
              if (analyzed.naturalPx > naturalPx) naturalPx = analyzed.naturalPx;
              if (analyzed.hasCjk) hasCjk = true;
            } else if (src.kind === "picture") {
              lineItems.push({
                kind: "picture",
                inlineIndex,
                xPx: xFrag,
                widthPx: src.widthPx,
                heightPx: src.heightPx,
              });
              if (src.heightPx > tallestPicturePx) tallestPicturePx = src.heightPx;
            }
            xFrag += frag.gapBefore + frag.occupiedWidth;
            endInlineIndex = inlineIndex;
          }
          xLine += line.width;
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
          xLine += tabAdvance(group.tab, xLine, tabs, group.followingPx);
          endInlineIndex = group.closerIndex;
        } else if (group.hardBreak) {
          endInlineIndex = group.closerIndex;
          break;
        }
      }
    }

    // Floor the resolver: pictures size their line; a line with no text at
    // all bottoms at the paragraph strut (the ¶-mark line box).
    let height = opts.lineHeight({ naturalPx, hasCjk });
    if (!hasText && opts.strutPx != null && opts.strutPx > height) height = opts.strutPx;
    if (tallestPicturePx > height) height = tallestPicturePx;
    // A picture is line content: its height is part of the line's natural
    // box. Without this, a docGrid line carrying only a picture reports a
    // text-sized natural, and the painter's grid centering sinks the picture
    // by half its height instead of pinning it to the line top.
    if (tallestPicturePx > naturalPx) naturalPx = tallestPicturePx;

    lines.push({
      items: lineItems,
      endInlineIndex,
      maxWidthPx: maxWidth,
      heightPx: height,
      naturalPx,
      hangPx: hangPx > 0 ? hangPx : undefined,
    });
    y += height;
    lineIndex++;
  }
  return lines;
}
