// The flow strategy — docx page boxing. Blocks stack into fixed-height pages
// (the section's content box); what doesn't fit splits at legal boundaries
// (paragraph lines, table rows, group children) and continues on the next
// page. This ports the C-route paginator's semantics off the DOM:
//
// - A paragraph splits at line boundaries; with widowControl on, head and
//   tail each keep ≥2 lines (Word) — a 3-line paragraph that would split 2+1
//   moves whole instead.
// - keepLines never splits while the paragraph fits a page; a block taller
//   than any page relaxes the keep/widow constraints and splits greedily
//   (progress beats clipping — Word relaxes the same way).
// - keepNext pulls previous block(s) to the next page when this block can't
//   start under them (a heading never orphans at a page bottom); the
//   walk-back cascades through consecutive keepNext blocks.
// - pageBreak atoms and pageBreakBefore close the page before their content.
// - Stacking margins follow the BFC model everywhere: a box (page or cell)
//   contains its first `before` and last `after`, middles collapse at the
//   max — one vertical-margin model shared with table cells.

import { layoutBlock, stackBlocks } from "../block/block";
import { fitExtentPx } from "../block/geometry";
import {
  type LayoutBlock,
  type LayoutBlockContext,
  type LayoutFloatZone,
  type ProjectedColumns,
  wrapEffectsOf,
} from "../layout-doc";
import type {
  LaidOutBlock,
  LaidOutCell,
  LaidOutFootnoteArea,
  LaidOutFootnoteNote,
  LaidOutLine,
  LaidOutParagraph,
  LaidOutRow,
  LaidOutStackItem,
  LaidOutTable,
} from "../layout-result";
import type { TextMeasurer } from "../text/measure";

/** One block placed on a page: `yPx` from the page content top. Split blocks
 *  appear as independent slices (a paragraph's head/tail share nothing but
 *  the line data). */
export interface FlowItem {
  yPx: number;
  block: LaidOutBlock;
  /** The item's column left edge within the content box, px (multi-column
   *  sections, w:cols) — absent in a single-column flow (full width). */
  xPx?: number;
}

export interface FlowPage {
  items: FlowItem[];
  /** Footnotes placed at the bottom of this page (absent when none). */
  footnotes?: LaidOutFootnoteArea;
  /** Unbounded flows only: the y where the content ends (footnote area
   *  included) — the host sizes the continuous page from it. Absent on
   *  paginated pages (their height is the section's paper). */
  contentBottomPx?: number;
}

/** Furniture-driven body insets for one page slot (px, measured from the
 *  content box edges): a tall header pushes the body down (topPx), a tall
 *  footer pushes it up (bottomPx) — Word's overflow rule, each page by its
 *  own slot's stack. The page pattern repeats: page 0 uses `first` when
 *  present, odd indexes use `even`, everything else `default`. */
export interface FlowPageInsets {
  first?: { topPx: number; bottomPx: number };
  even?: { topPx: number; bottomPx: number };
  default?: { topPx: number; bottomPx: number };
}

export interface FlowOptions {
  contentWidthPx: number;
  contentHeightPx: number;
  /** Section document-grid pitch (threads into every block layout). */
  linePitchPx?: number;
  /** w:adjustLineHeightInTable (settings.xml compat) — lets table cell lines
   *  join this grid. Absent leaves cells grid-free (the OOXML default). */
  adjustLinesInTable?: boolean;
  /** Per-slot body insets from the header/footer stacks (absent = margins
   *  rule everywhere). */
  pageInsets?: FlowPageInsets;
  /** This flow's starting page in the document's global page sequence — the
   *  odd/even inset slot keys off the PHYSICAL page number (Word's
   *  evenAndOddHeaders is document-wide), while the first slot stays local
   *  (a section's own first page). Absent = 0 (single-section documents). */
  pageOffset?: number;
  /** Section columns (w:cols) — the content box splits into `count` columns
   *  the flow fills left to right before paging. Absent = one full-width
   *  column. */
  columns?: ProjectedColumns;
  /** Footnote id → definition blocks (absent when document has no footnotes). */
  footnoteDefinitions?: Map<number, readonly LayoutBlock[]>;
  /** Endnote id → definition blocks (absent when document has no endnotes). */
  endnoteDefinitions?: Map<number, readonly LayoutBlock[]>;
  /** Word's Web Layout / Read Mode: the content never pages — every block
   *  stacks onto one continuous page, explicit breaks are inert, and the
   *  page reports its content bottom via {@link FlowPage.contentBottomPx}.
   *  `contentHeightPx` is ignored. */
  unbounded?: boolean;
}

/** Lay a block flow into pages. Always returns at least one page (an empty
 *  flow yields one empty page — a document always has a page). */
export function layoutFlow(
  blocks: readonly LayoutBlock[],
  opts: FlowOptions,
  measurer: TextMeasurer,
): FlowPage[] {
  const flow = new Flow(opts, measurer);
  for (const block of blocks) flow.push(block);
  return flow.finish();
}

/** Split a content width into w:cols column boxes — the single source the
 *  flow fills against and the painter derives separator positions from.
 *  Explicit w:col widths keep their own width/gap; equal (the default)
 *  divides the box minus the gaps evenly — Word shrinks the columns, never
 *  the gap. */
export function columnBoxesOf(
  contentWidthPx: number,
  cols: ProjectedColumns | undefined,
): { xPx: number; widthPx: number }[] {
  if (!cols || cols.count <= 1) return [{ xPx: 0, widthPx: contentWidthPx }];
  if (!cols.equalWidth && cols.columnsPx && cols.columnsPx.length > 0) {
    const boxes: { xPx: number; widthPx: number }[] = [];
    let x = 0;
    for (const widthPx of cols.columnsPx) {
      boxes.push({ xPx: x, widthPx });
      x += widthPx + cols.spacePx;
    }
    return boxes;
  }
  const width = (contentWidthPx - cols.spacePx * (cols.count - 1)) / cols.count;
  return Array.from({ length: cols.count }, (_, i) => ({
    xPx: (width + cols.spacePx) * i,
    widthPx: width,
  }));
}

/** One section of a multi-section document: its block flow plus the flow
 *  geometry it paginates against (page size, margins, grid, insets). */
export interface FlowSection {
  blocks: readonly LayoutBlock[];
  opts: FlowOptions;
}

export interface SectionedFlowPages {
  pages: FlowPage[];
  /** Global page index → section index (parallel to `pages`). */
  sectionOfPage: number[];
}

/** Lay a multi-section document into one continuous page list. Each section
 *  starts on a fresh page (OOXML's nextPage section break); `pageOffset`
 *  threads each section's starting global page number so the odd/even inset
 *  slot follows the physical page across section boundaries. */
export function layoutFlowSections(
  sections: readonly FlowSection[],
  measurer: TextMeasurer,
): SectionedFlowPages {
  const pages: FlowPage[] = [];
  const sectionOfPage: number[] = [];
  sections.forEach((section, i) => {
    for (const page of layoutFlow(
      section.blocks,
      { ...section.opts, pageOffset: pages.length },
      measurer,
    )) {
      pages.push(page);
      sectionOfPage.push(i);
    }
  });
  return { pages, sectionOfPage };
}

/** Footnote separator width in px (Word default: 2 inches = 144 pt = 192 px). */
export const FOOTNOTE_SEPARATOR_WIDTH_PX = 192;
/** Height of the footnote separator stroke and spacing (px):
 *  10px space before separator line, 1px line, 6px space after before first note. */
export const FOOTNOTE_SEPARATOR_HEIGHT_PX = 17;

function noteRefsInBlock(block: LaidOutBlock): { id: number; ordinal: number }[] {
  if (block.kind === "paragraph") {
    const refs: { id: number; ordinal: number }[] = [];
    const seen = new Set<number>();
    for (const line of block.lines) {
      for (const item of line.items) {
        const inl = block.inline[item.inlineIndex];
        if (inl?.kind === "text" && inl.noteRef?.kind === "footnote") {
          if (!seen.has(inl.noteRef.id)) {
            seen.add(inl.noteRef.id);
            refs.push({ id: inl.noteRef.id, ordinal: inl.noteRef.ordinal });
          }
        }
      }
    }
    return refs;
  }
  if (block.kind === "table") {
    const refs: { id: number; ordinal: number }[] = [];
    const seen = new Set<number>();
    for (const row of block.rows) {
      for (const cell of row.cells) {
        for (const item of cell.stack) {
          for (const r of noteRefsInBlock(item.block)) {
            if (!seen.has(r.id)) {
              seen.add(r.id);
              refs.push(r);
            }
          }
        }
      }
    }
    return refs;
  }
  if (block.kind === "group") {
    const refs: { id: number; ordinal: number }[] = [];
    const seen = new Set<number>();
    for (const child of block.children) {
      for (const r of noteRefsInBlock(child.block)) {
        if (!seen.has(r.id)) {
          seen.add(r.id);
          refs.push(r);
        }
      }
    }
    return refs;
  }
  return [];
}

class Flow {
  private readonly pages: FlowPage[] = [];
  private readonly items: FlowItem[] = [];
  private y = 0;
  private prevAfter = 0;
  private firstOnPage = true;
  /** The current page was opened by automatic pagination (overflow), not by
   *  an explicit break — Word never paints space-before on such a page's
   *  first block (COM-verified: a 600-twip before shows after a manual page
   *  break, never after overflow; section-start pages keep it too). */
  private autoBreak = false;
  /** Zero-based index of the page being filled — picks the slot's insets. */
  private pageIndex = 0;
  /** Partial-overlap float zones (wrap square/tight): lines inside the band
   *  wrap beside the box. Registered when the anchor paragraph commits — the
   *  anchor paragraph's own lines wrap via the paragraph module's self-zones,
   *  these cover every later paragraph on the page. */
  private readonly zones: LayoutFloatZone[] = [];
  /** Full-width cleared bands (wrap topAndBottom, or a square box covering
   *  the whole column): no text inside; blocks split at the band top and
   *  resume below it. Both lists are page-local — floats never cross a page
   *  boundary (Word anchors the box to the page its paragraph lands on). */
  private readonly bands: LayoutFloatZone[] = [];
  /** The page's column boxes (content-box-relative x + width, w:cols) and
   *  the index being filled — a single-column flow carries one full-width
   *  box and every column branch below short-circuits. */
  private readonly cols: { xPx: number; widthPx: number }[];
  private colIndex = 0;
  /** Cached laid footnote stacks (id → stack, height, ordinal). */
  private readonly laidFootnoteCache = new Map<
    number,
    { stack: LaidOutStackItem[]; heightPx: number; ordinal: number }
  >();
  /** Footnote IDs referenced on the current page in reference order. */
  private readonly pageFootnoteIds: number[] = [];
  /** Total height of the current page's footnote area (separator + notes). */
  private pageFootnoteHeight = 0;

  constructor(
    private readonly opts: FlowOptions,
    private readonly measurer: TextMeasurer,
  ) {
    this.cols = columnBoxesOf(opts.contentWidthPx, opts.columns);
    // The body starts below the header's push (first page / default slot).
    this.y = this.insets().topPx;
  }

  private get col(): { xPx: number; widthPx: number } {
    return this.cols[this.colIndex] ?? this.cols[this.cols.length - 1]!;
  }

  /** The current page's furniture insets (the section's first page →
   *  `first`, an odd PHYSICAL page → `even`, else `default`; missing slots
   *  fall back to `default`). The odd/even test uses the global page number
   *  (`pageOffset` + the local index) so the pattern is continuous across
   *  section boundaries, Word's evenAndOddHeaders semantics. */
  private insets(): { topPx: number; bottomPx: number } {
    const pi = this.opts.pageInsets;
    if (!pi) return { topPx: 0, bottomPx: 0 };
    const globalIndex = this.pageIndex + (this.opts.pageOffset ?? 0);
    const slot =
      this.pageIndex === 0 && pi.first
        ? pi.first
        : globalIndex % 2 === 1 && pi.even
          ? pi.even
          : pi.default;
    return { topPx: slot?.topPx ?? 0, bottomPx: slot?.bottomPx ?? 0 };
  }

  private get ctx(): LayoutBlockContext {
    return {
      linePitchPx: this.opts.linePitchPx,
      adjustLinesInTable: this.opts.adjustLinesInTable,
      onGrid: true,
      floatZones: this.zones.length > 0 ? this.zones : undefined,
      startY: this.y,
    };
  }

  private getLaidFootnote(
    id: number,
    ordinal: number,
  ): { stack: LaidOutStackItem[]; heightPx: number; ordinal: number } | undefined {
    const cached = this.laidFootnoteCache.get(id);
    if (cached) return cached;
    const blocks = this.opts.footnoteDefinitions?.get(id);
    if (!blocks || blocks.length === 0) return undefined;
    const stacked = stackBlocks(blocks, this.opts.contentWidthPx, undefined, this.measurer);
    const entry = { stack: stacked.stack, heightPx: stacked.heightPx, ordinal };
    this.laidFootnoteCache.set(id, entry);
    return entry;
  }

  private extraFootnoteHeightFor(laid: LaidOutBlock): number {
    if (!this.opts.footnoteDefinitions || this.opts.footnoteDefinitions.size === 0) return 0;
    const refs = noteRefsInBlock(laid);
    if (refs.length === 0) return 0;
    let extra = 0;
    let isFirst = this.pageFootnoteIds.length === 0;
    for (const ref of refs) {
      if (this.pageFootnoteIds.includes(ref.id)) continue;
      if (isFirst) {
        extra += FOOTNOTE_SEPARATOR_HEIGHT_PX;
        isFirst = false;
      }
      const note = this.getLaidFootnote(ref.id, ref.ordinal);
      if (note) extra += note.heightPx;
    }
    return extra;
  }

  private registerFootnotes(laid: LaidOutBlock): void {
    if (!this.opts.footnoteDefinitions || this.opts.footnoteDefinitions.size === 0) return;
    const refs = noteRefsInBlock(laid);
    for (const ref of refs) {
      if (!this.pageFootnoteIds.includes(ref.id)) {
        if (this.pageFootnoteIds.length === 0) {
          this.pageFootnoteHeight += FOOTNOTE_SEPARATOR_HEIGHT_PX;
        }
        this.pageFootnoteIds.push(ref.id);
        const note = this.getLaidFootnote(ref.id, ref.ordinal);
        if (note) this.pageFootnoteHeight += note.heightPx;
      }
    }
  }

  private resyncFootnotes(): void {
    this.pageFootnoteIds.length = 0;
    this.pageFootnoteHeight = 0;
    for (const item of this.items) {
      this.registerFootnotes(item.block);
    }
  }

  private buildPageFootnotes(): LaidOutFootnoteArea | undefined {
    if (this.pageFootnoteIds.length === 0) return undefined;
    const notes: LaidOutFootnoteNote[] = [];
    const items: LaidOutStackItem[] = [];
    let curY = FOOTNOTE_SEPARATOR_HEIGHT_PX;
    for (const id of this.pageFootnoteIds) {
      const laidNote = this.laidFootnoteCache.get(id);
      if (!laidNote) continue;
      notes.push({
        id,
        ordinal: laidNote.ordinal,
        stack: laidNote.stack,
        heightPx: laidNote.heightPx,
      });
      for (const item of laidNote.stack) {
        items.push({
          yPx: curY + item.yPx,
          block: item.block,
        });
      }
      curY += laidNote.heightPx;
    }
    const totalHeightPx = curY;
    const yPx = this.opts.contentHeightPx - this.insets().bottomPx - totalHeightPx;
    return {
      yPx,
      separatorWidthPx: FOOTNOTE_SEPARATOR_WIDTH_PX,
      notes,
      items,
      totalHeightPx,
    };
  }

  private remaining(): number {
    if (this.opts.unbounded) return Infinity;
    return this.opts.contentHeightPx - this.insets().bottomPx - this.pageFootnoteHeight - this.y;
  }

  /** The body a fresh page offers (px): the content box net of BOTH furniture
   *  insets — the mid-row whole-page room test measures against the page the
   *  row would move to, so a tall footer shrinks it too. */
  private freshPageBodyPx(): number {
    const ins = this.insets();
    return this.opts.contentHeightPx - ins.topPx - ins.bottomPx;
  }

  /** The nearest cleared-band top above `this.y`, or Infinity — the ceiling
   *  the current block must not cross (its overflow lines continue below
   *  the band). */
  private bandCeiling(): number {
    let top = Infinity;
    for (const b of this.bands) if (b.topPx > this.y + 0.01 && b.topPx < top) top = b.topPx;
    return top;
  }

  /** Drop `this.y` below the band it sits inside (a block that ended within
   *  a band) and below the nearest upcoming band (nothing fit above it). */
  private dodgeBands(): boolean {
    let moved = false;
    for (const b of this.bands) {
      if (this.y < b.bottomPx && this.y >= b.topPx - 0.01) {
        this.y = b.bottomPx;
        moved = true;
      }
    }
    return moved;
  }

  /** Seal the current page's items and start a fresh page. `auto` marks an
   *  overflow-driven break — its page's first block drops its space-before.
   *  An empty page emits nothing (Word collapses a break at a fresh page
   *  top), so pageIndex re-syncs to the emitted count — a no-op break must
   *  not skew the even/odd inset slots of every page after it. */
  private newPage(auto = false): void {
    if (this.items.length > 0) {
      const footnotes = this.buildPageFootnotes();
      this.pages.push({ items: this.items.splice(0), footnotes });
    }
    this.pageIndex = this.pages.length;
    this.colIndex = 0;
    this.y = this.insets().topPx;
    this.prevAfter = 0;
    this.firstOnPage = true;
    this.autoBreak = auto;
    this.zones.length = 0;
    this.bands.length = 0;
    this.pageFootnoteIds.length = 0;
    this.pageFootnoteHeight = 0;
  }

  /** Close the current column and continue at the top of the next one — a
   *  fresh page past the last column. Column tops keep their space-before
   *  (only overflow page tops drop it, so this routes through the manual
   *  newPage). Zones/bands reset like a page: floats are box-local. */
  private newColumn(): void {
    if (this.colIndex >= this.cols.length - 1) {
      this.newPage();
      return;
    }
    this.colIndex += 1;
    this.y = this.insets().topPx;
    this.prevAfter = 0;
    this.zones.length = 0;
    this.bands.length = 0;
  }

  finish(): FlowPage[] {
    if (this.opts.unbounded) {
      // Seal the one continuous page, reporting where the content ends
      // (footnote area included) — the host sizes the page from it.
      const bottom = this.y + this.pageFootnoteHeight;
      const footnotes = this.buildPageFootnotes();
      this.pages.push({ items: this.items.splice(0), footnotes, contentBottomPx: bottom });
    } else {
      this.newPage();
    }
    if (this.pages.length === 0) this.pages.push({ items: [] });
    return this.pages;
  }

  /** Page-top spacing: when the page was opened by automatic pagination the
   *  first block paints no space-before (Word's overflow-page behavior — see
   *  autoBreak). Manual breaks and section-start pages keep the margin. */
  private spacingBefore(laid: LaidOutBlock): number {
    const m = marginBefore(laid, this.prevAfter, this.firstOnPage);
    return this.firstOnPage && this.autoBreak ? 0 : m;
  }

  /** Place a block: whole, split, or moved to the next page. */
  push(block: LayoutBlock): void {
    // An unbounded flow (Web Layout) has no pages to break — explicit break
    // atoms are inert (Word paints nothing for them there).
    if (this.opts.unbounded && (block.kind === "pageBreak" || block.kind === "columnBreak")) {
      return;
    }
    if (block.kind === "pageBreak") {
      const laid = layoutBlock(block, this.opts.contentWidthPx, this.ctx, this.measurer);
      // A break row the page's last line cannot hold collapses into the page
      // bottom (zero height — the marker still paints at the edge): the page
      // closes without spilling its own marker past the page edge, and the
      // break never opens an empty page for itself. Word's blank-page trap,
      // cut at the root.
      const fits = this.y + laid.heightPx <= this.opts.contentHeightPx - this.insets().bottomPx;
      this.commit(fits ? laid : { ...laid, heightPx: 0 }, 0);
      this.newPage();
      return;
    }
    if (block.kind === "columnBreak") {
      // A column break closes the column committing nothing (Word paints no
      // marker row); in a one-column flow newColumn closes the page instead.
      this.newColumn();
      return;
    }
    // pageBreakBefore is a break atom before the content (inert unbounded —
    // Web Layout has no page edge to force).
    if (block.kind === "paragraph" && block.pageBreakBefore && !this.opts.unbounded) this.newPage();

    const laid = layoutBlock(block, this.col.widthPx, this.ctx, this.measurer);
    if (this.tryPlace(laid)) return;
    // Blocked by a band rather than the page bottom: resume below the band.
    if (this.dodgeBands()) {
      this.pushLaid(laid);
      return;
    }
    // Out of room in this column: Word fills the next one before paging.
    // Re-lay against the fresh column — zones are box-local, so lines laid
    // against the old column's zones/y are stale there.
    if (this.colIndex < this.cols.length - 1) {
      this.newColumn();
      const fresh = layoutBlock(block, this.col.widthPx, this.ctx, this.measurer);
      if (this.tryPlace(fresh)) return;
      this.pushLaid(fresh);
      return;
    }
    // A single-line section-break paragraph the page bottom cannot hold
    // collapses into the page (zero height — its marker still paints at the
    // edge) instead of dropping whole onto a fresh page it would leave empty:
    // the paragraph is the section's last, so a dropped one opens a blank
    // page under the old section (Word's classic undeletable blank page). A
    // multi-line one carries content — it splits normally above.
    if (
      block.kind === "paragraph" &&
      block.sectionEnd &&
      laid.kind === "paragraph" &&
      laid.lines.length === 1 &&
      this.items.length > 0
    ) {
      this.commit({ ...laid, heightPx: 0 }, 0);
      return;
    }
    // Nothing fit here: blocks placed before this one with keepNext move
    // along (a heading stays with the paragraph that follows it). The pulled
    // blocks precede this one — re-place them first, then this block.
    const pulled = this.pullKeepNext();
    this.newPage(true);
    for (const kept of pulled) this.pushLaid(kept);
    // Re-lay on the fresh page: float zones are page-local, so the lines
    // laid against the old page's zones/y are stale here.
    const fresh = layoutBlock(block, this.col.widthPx, this.ctx, this.measurer);
    if (!this.tryPlace(fresh)) {
      // Cannot fit even a full empty page — place whole and overflow.
      this.commit(fresh, this.spacingBefore(fresh));
    }
  }

  /** Continue placing an already-laid block (split tails, pulled keepNext). */
  private pushLaid(laid: LaidOutBlock): void {
    if (this.tryPlace(laid)) return;
    // Out of room: fill the next column before paging (same rule as push's
    // whole-block path). The tail keeps its wrapping — equal-width columns
    // reflow nothing, and a split tail is always born from a column that just
    // filled, i.e. this is the tail's normal continuation.
    if (this.colIndex < this.cols.length - 1) {
      this.newColumn();
      if (this.tryPlace(laid)) return;
    }
    this.newPage(true);
    if (!this.tryPlace(laid)) this.commit(laid, this.spacingBefore(laid));
  }

  /** A grid picture line fits when its spanned rows fit the remaining row
   *  budget — Word quantizes the page fit at line-pitch granularity, so the
   *  padded span may cross the page bottom by the trailing partial row while
   *  the picture box (centered in the span) keeps its ink above it. The ink
   *  rule above is the first test; this catches the boundary case the row
   *  math owns (pixel-verified: a 464px picture spanning 21 rows stays on a
   *  page whose room is 20.6 rows — its ink clears the bottom by 1px, the
   *  padded box crosses by 9). */
  private gridRowsFit(laid: LaidOutBlock, room: number): boolean {
    const pitch = this.opts.linePitchPx;
    const last = laid.kind === "paragraph" ? laid.lines[laid.lines.length - 1] : undefined;
    if (!pitch || !last?.pictureFloored) return false;
    return last.heightPx <= Math.ceil(room / pitch) * pitch;
  }

  /** Try to place `laid` on the current page; on overflow, split at a legal
   *  boundary that fits (even on an empty page — a block taller than the page
   *  still fills it). The room is the smaller of the page bottom and the
   *  nearest cleared band top — lines resume below the band. Returns true
   *  when anything was placed. */
  private tryPlace(laid: LaidOutBlock): boolean {
    // Never place inside a cleared band — drop below it first (whether this
    // is a fresh push, a split tail, or a re-placed keepNext block).
    this.dodgeBands();
    const before = this.spacingBefore(laid);
    const extraFnH = this.extraFootnoteHeightFor(laid);
    // Both spans are relative to this.y: the page bottom and the band top.
    const room = Math.min(this.remaining() - extraFnH, this.bandCeiling() - this.y) - before;
    // `room` is already net of the before-margin, so the check adds the
    // extent alone — adding `before` here too would count the margin twice
    // and evict blocks Word keeps (a picture paragraph with an 8px before on
    // a near-full page, pixel-verified against the reference render).
    if (fitExtentPx(laid) <= room || this.gridRowsFit(laid, room)) {
      this.commit(laid, before);
      return true;
    }
    const { k, midDepth } = this.sliceFitting(laid, room);
    // k = 0 with a midDepth is the force-split of a first row no page could
    // hold — the head is just that row's upper half.
    if (k > 0 || midDepth != null) {
      const [head, tail] = splitLaid(laid, k, midDepth);
      this.commit(head, before);
      this.pushLaid(tail);
      return true;
    }
    return false;
  }

  private commit(laid: LaidOutBlock, before: number): void {
    const yPx = this.y + before;
    // Multi-column flows stamp each item with its column's left edge (the
    // painter and the caret map offset by it); single-column stays undefined.
    const xPx = this.cols.length > 1 ? this.col.xPx : undefined;
    this.items.push({ yPx, block: laid, xPx });
    this.y += before + laid.heightPx;
    this.prevAfter = laid.kind === "paragraph" ? laid.afterPx : 0;
    this.firstOnPage = false;
    this.registerFootnotes(laid);
    if (laid.kind === "paragraph") this.registerFloats(laid, yPx);
  }

  /** Turn the anchor paragraph's wrapped drawings into flow effects: a
   *  paragraph-anchored box offset into the column either shrinks the lines
   *  it overlaps (a zone) or clears its whole band (topAndBottom / a box
   *  covering the full column width). The box grows by the anchor's wrap
   *  distances first (distL/R/T/B), so text keeps its Word gap. Margin/
   *  page-anchored and aligned boxes stay painter-only (registered gaps). */
  private registerFloats(laid: Extract<LaidOutBlock, { kind: "paragraph" }>, yPx: number): void {
    const { zones, bands } = wrapEffectsOf(laid.drawings, yPx, this.opts.contentWidthPx);
    this.zones.push(...zones);
    this.bands.push(...bands);
  }

  /** Detach the trailing run of placed keepNext blocks (cascades through the
   *  whole run — heading + subheading chains) and re-sync the page fill. */
  private pullKeepNext(): LaidOutBlock[] {
    let cut = this.items.length;
    while (cut > 0) {
      const prev = this.items[cut - 1].block;
      if (!(prev.kind === "paragraph" && prev.keepNext)) break;
      cut--;
    }
    if (cut === this.items.length) return [];
    const moved = this.items.splice(cut).map((item) => item.block);
    const last = this.items[this.items.length - 1];
    if (last) {
      this.y = last.yPx + last.block.heightPx;
      this.prevAfter = last.block.kind === "paragraph" ? last.block.afterPx : 0;
      this.firstOnPage = false;
    } else {
      this.y = 0;
      this.prevAfter = 0;
      this.firstOnPage = true;
    }
    this.resyncFootnotes();
    return moved;
  }

  /** The next split: `k` = the prefix count whose stacked height fits `space`,
   *  plus `midDepth` (px from the k-th table row's top) when that row itself
   *  splits mid-content. `k` = 0 means nothing fits — move the whole block. */
  private sliceFitting(laid: LaidOutBlock, space: number): { k: number; midDepth?: number } {
    if (space <= 0) return { k: 0 };
    // A block taller than the page ignores keepLines/widowControl — pushing
    // it to the next page cannot help, so it splits greedily (Word relaxes
    // the same way; progress beats clipping). An unbounded flow never splits
    // mid-block: there is no page edge to relax against.
    const overPage = !this.opts.unbounded && laid.heightPx > this.opts.contentHeightPx;
    switch (laid.kind) {
      case "paragraph": {
        if (laid.keepLines && !overPage) return { k: 0 };
        let k = 0;
        let h = 0;
        while (k < laid.lines.length && h + laid.lines[k].heightPx <= space) {
          h += laid.lines[k].heightPx;
          k++;
        }
        if (k === 0) return { k: 0 };
        if (!overPage && laid.widowControl !== false && laid.lines.length >= 2) {
          if (k === 1) return { k: 0 }; // an orphaned single head line — move whole
          if (laid.lines.length - k === 1) k--; // tail widow — give it a line
          if (k < 2) return { k: 0 }; // can't satisfy both — move whole
        }
        return { k };
      }
      case "table": {
        // Word repeats a leading tblHeader band on every continuation page, so
        // a cut must keep the whole band + at least one body row here, and a
        // band taller than the page's body area gives up repeating (Word's
        // anti-loop rule) — the table then splits as if unmarked.
        const headers = headerRowCount(laid);
        if (headers > 0 && headers < laid.rows.length) {
          const headerH = sumRows(laid.rows.slice(0, headers));
          const body = this.freshPageBodyPx();
          if (headerH <= body) {
            // A body row re-opens under a fresh band copy on the next page,
            // so its whole-page room is the body net of the band.
            const pageRoom = body - headerH;
            let k = headers;
            let h = headerH;
            while (k < laid.rows.length && h + laid.rows[k].heightPx <= space) {
              h += laid.rows[k].heightPx;
              k++;
            }
            // No body row fits under the band whole — try splitting it
            // mid-content, else move the table whole.
            if (k === headers) {
              const mid = this.midRowDepth(laid.rows[k], space - h, pageRoom);
              return mid != null ? { k, midDepth: mid } : { k: 0 };
            }
            const mid =
              k < laid.rows.length
                ? this.midRowDepth(laid.rows[k], space - h, pageRoom)
                : undefined;
            return { k, midDepth: mid };
          }
        }
        let k = 0;
        let h = 0;
        while (k < laid.rows.length && h + laid.rows[k].heightPx <= space) {
          h += laid.rows[k].heightPx;
          k++;
        }
        // The first row alone doesn't fit: force-split it when no page could
        // hold it whole (Word: progress beats clipping), else move whole —
        // midRowDepth itself refuses pageable rows.
        if (k === 0) {
          const mid = this.midRowDepth(laid.rows[0], space, this.freshPageBodyPx());
          return mid != null ? { k: 0, midDepth: mid } : { k: 0 };
        }
        const mid = this.midRowDepth(laid.rows[k], space - h, this.freshPageBodyPx());
        return { k, midDepth: mid };
      }
      case "group": {
        let k = 0;
        let h = 0;
        let prevAfter = 0;
        while (k < laid.children.length) {
          const child = laid.children[k].block;
          const m = marginBefore(child, prevAfter, k === 0);
          if (h + m + child.heightPx > space) break;
          h += m + child.heightPx;
          prevAfter = child.kind === "paragraph" ? child.afterPx : 0;
          k++;
        }
        return { k };
      }
      case "placeholder":
      case "pageBreak":
        return { k: 0 };
    }
  }

  /** The depth a row that doesn't fit splits at (`depth` px from its top), or
   *  undefined when it must move whole: Word only ever splits a row no page
   *  could hold whole — one taller than `pageRoom` (the fresh-page body,
   *  net of the repeated band's own height) force-splits mid-content
   *  (progress beats clipping, cantSplit or not), while any pageable row
   *  moves whole (COM-verified: wrapped 2/3/4-line single paragraphs, a cut
   *  at a paragraph boundary, widowControl off, and a 40-paragraph row all
   *  refuse the cut); exact heights never split (overflow clips); and the
   *  cut needs content on BOTH sides in some cell — an empty-shell head or
   *  tail moves the row instead. */
  private midRowDepth(
    row: LaidOutRow | undefined,
    depth: number,
    pageRoom: number,
  ): number | undefined {
    if (!row || depth <= 0 || depth >= row.heightPx) return undefined;
    if (row.exactHeight || row.heightPx <= pageRoom) return undefined;
    let head = false;
    let tail = false;
    for (const cell of row.cells) {
      const d = depth - (cell.contentOffsetYPx ?? 0);
      let cut = 0;
      while (cut < cell.stack.length && cell.stack[cut].yPx + cell.stack[cut].block.heightPx <= d) {
        cut++;
      }
      if (cut > 0) head = true;
      const crossing = cell.stack[cut];
      if (
        crossing?.block.kind === "paragraph" &&
        fittingLines(crossing.block, d - crossing.yPx) > 0
      ) {
        head = true;
      }
      if (cut < cell.stack.length) tail = true;
    }
    return head && tail ? depth : undefined;
  }
}

/** The stacked `before` margin: the first block of a box counts in full (a
 *  page/cell is a BFC), later blocks collapse against `prevAfter` at max —
 *  except a table, which never inherits the preceding paragraph's space-after
 *  (Word drops space-before-a-table; corpus-verified on the honor table:
 *  ref row-0 top = heading bottom, its 4px after painted nowhere). */
function marginBefore(laid: LaidOutBlock, prevAfter: number, firstInBox: boolean): number {
  const before = laid.kind === "paragraph" ? laid.beforePx : 0;
  if (laid.kind === "table") return before;
  return firstInBox ? before : Math.max(prevAfter, before);
}

/** The table's repeat band: the contiguous tblHeader prefix from the first
 *  row (Word ignores a mark that doesn't start at the top). */
function headerRowCount(laid: LaidOutTable): number {
  let n = 0;
  while (n < laid.rows.length && laid.rows[n].tableHeader) n++;
  return n;
}

/** Split a laid block after its k-th line/row/child — or, for a table with
 *  `midDepth`, through the k-th row's interior at that depth. The head keeps
 *  the `before` margin; the tail carries no `before` (it continues) and keeps
 *  `after`. */
function splitLaid(laid: LaidOutBlock, k: number, midDepth?: number): [LaidOutBlock, LaidOutBlock] {
  switch (laid.kind) {
    case "paragraph": {
      const headLines = laid.lines.slice(0, k);
      const tailLines = rebaseLines(laid.lines.slice(k));
      return [
        { ...laid, lines: headLines, heightPx: sumLines(headLines) },
        // The tail drops the drawings: a float paints and registers its zone
        // on the page of its anchor paragraph (the head), never twice.
        {
          ...laid,
          lines: tailLines,
          heightPx: sumLines(tailLines),
          beforePx: 0,
          drawings: undefined,
        },
      ];
    }
    case "table": {
      const headers = headerRowCount(laid);
      const strip = (row: LaidOutTable["rows"][number]) =>
        row.tableHeader ? { ...row, tableHeader: undefined } : row;
      // Mid-row: the k-th row itself splits — the head keeps its upper half,
      // the tail re-opens with the lower half. Band copies apply exactly as
      // at a row-boundary cut (k ≥ headers keeps the whole band in the head;
      // the tail half is a body-row slice and never carries a mark itself).
      if (midDepth != null) {
        const [headHalf, tailHalf] = splitRowAt(laid.rows[k], midDepth);
        const headRows = [...laid.rows.slice(0, k), headHalf];
        const tailBody = [tailHalf, ...laid.rows.slice(k + 1).map(strip)];
        const tailRows = reopenSpannedCells(
          k >= headers ? [...laid.rows.slice(0, headers), ...tailBody] : tailBody,
          k >= headers
            ? [...leadingSlots(headers), ...rangeSlots(k, laid.rows.length)]
            : rangeSlots(k, laid.rows.length),
          laid.rows,
          laid.columnWidthsPx.length,
        );
        return [
          { ...laid, rows: headRows, heightPx: sumRows(headRows) },
          { ...laid, rows: tailRows, heightPx: sumRows(tailRows) },
        ];
      }
      const headRows = laid.rows.slice(0, k);
      // Word re-opens EVERY continuation with the header band when the cut
      // kept it whole (k > headers) — the tail's leading copies keep the mark
      // so a further cut repeats them again (the copy count stays constant:
      // each cut regenerates slice(0, headers) from the tail's own lead). The
      // non-copy tail rows lose any mark — a mid-table marked row must never
      // widen the band count on a later page. Any other cut (a give-up split,
      // or a table that is all header) continues band-less entirely: the
      // tail's marks are stripped so no later page re-derives a band from a
      // marked row that is no longer at a table top.
      const tailRows = reopenSpannedCells(
        k > headers
          ? [...laid.rows.slice(0, headers), ...laid.rows.slice(k).map(strip)]
          : laid.rows.slice(k).map(strip),
        k > headers
          ? [...leadingSlots(headers), ...rangeSlots(k, laid.rows.length)]
          : rangeSlots(k, laid.rows.length),
        laid.rows,
        laid.columnWidthsPx.length,
      );
      return [
        { ...laid, rows: headRows, heightPx: sumRows(headRows) },
        { ...laid, rows: tailRows, heightPx: sumRows(tailRows) },
      ];
    }
    case "group": {
      const head = laid.children.slice(0, k);
      const tail = rebaseChildren(laid.children.slice(k));
      return [
        { ...laid, children: head, heightPx: stackHeight(head) },
        { ...laid, children: tail, heightPx: stackHeight(tail) },
      ];
    }
    case "placeholder":
    case "pageBreak":
      throw new Error("pageBreak is not splittable");
  }
}

/** Word re-opens every vertical merge a page cut crosses: a tail row keeps an
 *  empty placeholder cell where a row-spanning cell anchored above the cut
 *  still owns columns — without it the tail's cells pack one column left and
 *  the first column appears to vanish. `slots` pairs each tail row with the
 *  source row index it renders (band copies render the leading rows again);
 *  an anchor whose own source row is in the tail re-opens through itself. */
function reopenSpannedCells(
  tailRows: LaidOutTable["rows"],
  slots: number[],
  sourceRows: LaidOutTable["rows"],
  columnCount: number,
): LaidOutTable["rows"] {
  // Walk the source grid once — the same occupancy walk the painter runs —
  // so each row knows where its own cells start and which spans reach down.
  const occ: (LaidOutCell | undefined)[][] = sourceRows.map(() => []);
  const at: { cell: LaidOutCell; c: number; spanW: number; spanH: number }[][] = sourceRows.map(
    () => [],
  );
  sourceRows.forEach((row, r) => {
    let c = 0;
    for (const cell of row.cells) {
      while (c < columnCount && occ[r]![c]) c++;
      if (c >= columnCount) break;
      const spanW = Math.min(cell.colspan ?? 1, columnCount - c);
      const spanH = Math.min(cell.rowspan ?? 1, sourceRows.length - r);
      for (let dr = 0; dr < spanH; dr++)
        for (let dc = 0; dc < spanW; dc++) occ[r + dr]![c + dc] = cell;
      at[r]!.push({ cell, c, spanW, spanH });
      c += spanW;
    }
  });

  const present = new Set(slots);
  return tailRows.map((row, i) => {
    const rt = slots[i];
    if (rt == null) return row;
    const open: { p: (typeof at)[number][number]; r: number }[] = [];
    for (const [r, list] of at.entries()) {
      if (present.has(r)) continue;
      for (const p of list) if (r < rt && rt < r + p.spanH) open.push({ p, r });
    }
    if (open.length === 0) return row;
    // Merge the placeholders into the row's cells by column order — the row's
    // own cells keep their source order (`at[rt]` walks in cells order).
    open.sort((a, b) => a.p.c - b.p.c);
    const cells: LaidOutCell[] = [];
    let ci = 0;
    let oi = 0;
    while (ci < row.cells.length || oi < open.length) {
      const next = open[oi];
      const ownC = at[rt]![ci]?.c ?? Infinity;
      if (next && next.p.c <= ownC) {
        const { p, r } = next;
        cells.push({
          colspan: p.spanW,
          rowspan: r + p.spanH - rt,
          insets: p.cell.insets,
          borders: p.cell.borders,
          fill: p.cell.fill,
          innerWidthPx: p.cell.innerWidthPx,
          stack: [],
        });
        oi++;
      } else {
        cells.push(row.cells[ci++]!);
      }
    }
    return { ...row, cells };
  });
}

/** The tail's source-row slots for the repeated band copies (rows 0..n). */
function leadingSlots(n: number): number[] {
  return Array.from({ length: n }, (_, i) => i);
}

/** The tail's source-row slots for source rows `from` (inclusive) to `to`. */
function rangeSlots(from: number, to: number): number[] {
  return Array.from({ length: Math.max(0, to - from) }, (_, i) => from + i);
}

/** Split a row mid-content at `depth` px from its top: every cell cuts at the
 *  same line (Word's split line crosses the whole row); cells whose content
 *  ends above the line ride whole in the head. The tail half loses any
 *  tblHeader mark — a band only re-opens at a row boundary. */
function splitRowAt(row: LaidOutRow, depth: number): [LaidOutRow, LaidOutRow] {
  const headCells: LaidOutCell[] = [];
  const tailCells: LaidOutCell[] = [];
  for (const cell of row.cells) {
    const [head, tail] = splitCellAt(cell, depth - (cell.contentOffsetYPx ?? 0));
    headCells.push(head);
    tailCells.push(tail);
  }
  return [
    { ...row, heightPx: depth, cells: headCells },
    {
      ...row,
      heightPx: Math.max(row.heightPx - depth, 0),
      cells: tailCells,
      tableHeader: undefined,
    },
  ];
}

/** Cut one cell's stack at `depth` px below its content top: the head keeps
 *  every full item plus the crossing paragraph's fitting lines; the tail
 *  carries the rest, rebased to its own top. A nested table/group crossing
 *  the line rides whole to the tail (only paragraphs split at line bounds). */
function splitCellAt(cell: LaidOutCell, depth: number): [LaidOutCell, LaidOutCell] {
  let cut = 0;
  while (cut < cell.stack.length && cell.stack[cut].yPx + cell.stack[cut].block.heightPx <= depth) {
    cut++;
  }
  const headStack: LaidOutStackItem[] = cell.stack.slice(0, cut);
  const rest: LaidOutStackItem[] = cell.stack.slice(cut);
  // The tail's base: where its first item's top sat in the cell's own space
  // (the crossing paragraph's top when it splits — the tail re-opens there).
  let base = rest[0]?.yPx ?? 0;
  const crossing = rest[0];
  if (crossing?.block.kind === "paragraph") {
    const k = fittingLines(crossing.block, depth - crossing.yPx);
    if (k > 0) {
      const headLines = crossing.block.lines.slice(0, k);
      const tailLines = rebaseLines(crossing.block.lines.slice(k));
      headStack.push({
        yPx: crossing.yPx,
        block: { ...crossing.block, lines: headLines, heightPx: sumLines(headLines) },
      });
      rest[0] = {
        yPx: 0,
        // The tail paragraph continues the split one — its `before` rode in
        // the head's item position already.
        block: { ...crossing.block, lines: tailLines, heightPx: sumLines(tailLines), beforePx: 0 },
      };
      base = crossing.yPx;
    }
  }
  return [
    { ...cell, stack: headStack },
    {
      ...cell,
      stack: rest.map((item) => ({ yPx: item.yPx - base, block: item.block })),
      // The tail is its own row slice — the head row's vAlign slack is void.
      contentOffsetYPx: undefined,
    },
  ];
}

/** Lines of a laid paragraph that fit within `height` px from its top. */
function fittingLines(block: LaidOutParagraph, height: number): number {
  let k = 0;
  let h = 0;
  while (k < block.lines.length && h + block.lines[k].heightPx <= height) {
    h += block.lines[k].heightPx;
    k++;
  }
  return k;
}

/** Re-derive a stack's y offsets after a cut: the new first child's full
 *  `before` counts (fresh box), later ones collapse at the max. */
function rebaseChildren(items: readonly LaidOutStackItem[]): LaidOutStackItem[] {
  let y = 0;
  let prevAfter = 0;
  let first = true;
  return items.map(({ block }) => {
    const m = marginBefore(block, prevAfter, first);
    y += m + block.heightPx;
    prevAfter = block.kind === "paragraph" ? block.afterPx : 0;
    first = false;
    return { yPx: y - block.heightPx, block };
  });
}

/** Stacked height from y-offset items: the last child's bottom plus its
 *  contained `after` margin. */
function stackHeight(items: readonly LaidOutStackItem[]): number {
  const last = items[items.length - 1];
  if (!last) return 0;
  const after = last.block.kind === "paragraph" ? last.block.afterPx : 0;
  return last.yPx + last.block.heightPx + after;
}

/** Re-derive a split tail's line y offsets. The first-line indent needs no
 *  handling here: it lives on the line (`firstLineIndentPx`), and a tail's
 *  leading line is mid-paragraph (slice(k), k ≥ 1), so it carries none. */
function rebaseLines(lines: readonly LaidOutLine[]): LaidOutLine[] {
  let y = 0;
  return lines.map((line) => {
    const out = { ...line, yPx: y };
    y += line.heightPx;
    return out;
  });
}

function sumLines(lines: readonly LaidOutLine[]): number {
  let h = 0;
  for (const l of lines) h += l.heightPx;
  return h;
}

function sumRows(rows: LaidOutTable["rows"]): number {
  let h = 0;
  for (const r of rows) h += r.heightPx;
  return h;
}
