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

import { layoutBlock } from "../block/block";
import type { LayoutBlock, LayoutBlockContext } from "../layout-doc";
import type { LaidOutBlock, LaidOutLine, LaidOutStackItem, LaidOutTable } from "../layout-result";
import type { TextMeasurer } from "../text/measure";

/** One block placed on a page: `yPx` from the page content top. Split blocks
 *  appear as independent slices (a paragraph's head/tail share nothing but
 *  the line data). */
export interface FlowItem {
  yPx: number;
  block: LaidOutBlock;
}

export interface FlowPage {
  items: FlowItem[];
}

export interface FlowOptions {
  contentWidthPx: number;
  contentHeightPx: number;
  /** Section document-grid pitch (threads into every block layout). */
  linePitchPx?: number;
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

class Flow {
  private readonly pages: FlowPage[] = [];
  private readonly items: FlowItem[] = [];
  private y = 0;
  private prevAfter = 0;
  private firstOnPage = true;

  constructor(
    private readonly opts: FlowOptions,
    private readonly measurer: TextMeasurer,
  ) {}

  private get ctx(): LayoutBlockContext {
    return { linePitchPx: this.opts.linePitchPx, onGrid: true };
  }

  private remaining(): number {
    return this.opts.contentHeightPx - this.y;
  }

  /** Seal the current page's items and start a fresh page. */
  private newPage(): void {
    if (this.items.length > 0) this.pages.push({ items: this.items.splice(0) });
    this.y = 0;
    this.prevAfter = 0;
    this.firstOnPage = true;
  }

  finish(): FlowPage[] {
    this.newPage();
    if (this.pages.length === 0) this.pages.push({ items: [] });
    return this.pages;
  }

  /** Place a block: whole, split, or moved to the next page. */
  push(block: LayoutBlock): void {
    if (block.kind === "pageBreak") {
      this.newPage();
      return;
    }
    // pageBreakBefore is a break atom before the content.
    if (block.kind === "paragraph" && block.pageBreakBefore) this.newPage();

    const laid = layoutBlock(block, this.opts.contentWidthPx, this.ctx, this.measurer);
    if (this.tryPlace(laid)) return;
    // Nothing fit here: blocks placed before this one with keepNext move
    // along (a heading stays with the paragraph that follows it). The pulled
    // blocks precede this one — re-place them first, then this block.
    const pulled = this.pullKeepNext();
    this.newPage();
    for (const kept of pulled) this.pushLaid(kept);
    if (!this.tryPlace(laid)) {
      // Cannot fit even a full empty page — place whole and overflow.
      this.commit(laid, marginBefore(laid, this.prevAfter, this.firstOnPage));
    }
  }

  /** Continue placing an already-laid block (split tails, pulled keepNext). */
  private pushLaid(laid: LaidOutBlock): void {
    if (this.tryPlace(laid)) return;
    this.newPage();
    if (!this.tryPlace(laid))
      this.commit(laid, marginBefore(laid, this.prevAfter, this.firstOnPage));
  }

  /** Try to place `laid` on the current page; on overflow, split at a legal
   *  boundary that fits (even on an empty page — a block taller than the page
   *  still fills it). Returns true when anything was placed. */
  private tryPlace(laid: LaidOutBlock): boolean {
    const before = marginBefore(laid, this.prevAfter, this.firstOnPage);
    if (before + laid.heightPx <= this.remaining()) {
      this.commit(laid, before);
      return true;
    }
    const k = this.sliceFitting(laid, this.remaining() - before);
    if (k > 0) {
      const [head, tail] = splitLaid(laid, k);
      this.commit(head, before);
      this.pushLaid(tail);
      return true;
    }
    return false;
  }

  private commit(laid: LaidOutBlock, before: number): void {
    this.items.push({ yPx: this.y + before, block: laid });
    this.y += before + laid.heightPx;
    this.prevAfter = laid.kind === "paragraph" ? laid.afterPx : 0;
    this.firstOnPage = false;
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
    return moved;
  }

  /** Max prefix count whose stacked height fits `space` (0 = none fit). */
  private sliceFitting(laid: LaidOutBlock, space: number): number {
    if (space <= 0) return 0;
    // A block taller than the page ignores keepLines/widowControl — pushing
    // it to the next page cannot help, so it splits greedily (Word relaxes
    // the same way; progress beats clipping).
    const overPage = laid.heightPx > this.opts.contentHeightPx;
    switch (laid.kind) {
      case "paragraph": {
        if (laid.keepLines && !overPage) return 0;
        let k = 0;
        let h = 0;
        while (k < laid.lines.length && h + laid.lines[k].heightPx <= space) {
          h += laid.lines[k].heightPx;
          k++;
        }
        if (k === 0) return 0;
        if (!overPage && laid.widowControl !== false && laid.lines.length >= 2) {
          if (k === 1) return 0; // an orphaned single head line — move whole
          if (laid.lines.length - k === 1) k--; // tail widow — give it a line
          if (k < 2) return 0; // can't satisfy both — move whole
        }
        return k;
      }
      case "table": {
        let k = 0;
        let h = 0;
        while (k < laid.rows.length && h + laid.rows[k].heightPx <= space) {
          h += laid.rows[k].heightPx;
          k++;
        }
        return k;
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
        return k;
      }
      case "placeholder":
      case "pageBreak":
        return 0;
    }
  }
}

/** The stacked `before` margin: the first block of a box counts in full (a
 *  page/cell is a BFC), later blocks collapse against `prevAfter` at max. */
function marginBefore(laid: LaidOutBlock, prevAfter: number, firstInBox: boolean): number {
  const before = laid.kind === "paragraph" ? laid.beforePx : 0;
  return firstInBox ? before : Math.max(prevAfter, before);
}

/** Split a laid block after its k-th line/row/child. The head keeps the
 *  `before` margin; the tail carries no `before` (it continues) and keeps
 *  `after`. */
function splitLaid(laid: LaidOutBlock, k: number): [LaidOutBlock, LaidOutBlock] {
  switch (laid.kind) {
    case "paragraph": {
      const headLines = laid.lines.slice(0, k);
      const tailLines = rebaseLines(laid.lines.slice(k));
      return [
        { ...laid, lines: headLines, heightPx: sumLines(headLines) },
        { ...laid, lines: tailLines, heightPx: sumLines(tailLines), beforePx: 0 },
      ];
    }
    case "table": {
      const headRows = laid.rows.slice(0, k);
      const tailRows = laid.rows.slice(k);
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
