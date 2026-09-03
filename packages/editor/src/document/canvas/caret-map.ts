// Pixel ↔ PM-position mapping for the canvas route — the click-to-caret and
// caret-rendering geometry. Built per relayout: the laid-out pages and the
// PM doc walked in parallel (both document order), each paragraph block
// zipped to its PM textblock. Within a paragraph the line items' collapsed
// character ranges resolve a click's x/y to a doc position and a position
// back to a page-local caret box; justified items scale the natural advances
// into their stretch interval exactly as the painter does.

import {
  cssFontOf,
  familyOfSlot,
  type FlowPage,
  gridPadOf,
  isCjkCodeUnit,
  type ItemGlyphLayout,
  itemGlyphLayout,
  justifiedIntervals,
  type LaidOutLine,
  type LaidOutParagraph,
  leaferBaselinePadPx,
  lineOriginXPx,
  lineSpaceGaps,
  tableGridOf,
  vertAlignBaselineShiftPx,
  vertAlignedSizePx,
} from "@docen/layout";
import type { Node as PmNode } from "@tiptap/pm/model";

export interface CaretRect {
  page: number;
  xPx: number;
  yPx: number;
  heightPx: number;
}

export interface SelectionRect {
  page: number;
  xPx: number;
  yPx: number;
  widthPx: number;
  heightPx: number;
}

/** One laid line with its page-local origin and collapsed-char range. */
interface LineEntry {
  page: number;
  para: LaidOutParagraph;
  owner: ParaEntry;
  line: LaidOutLine;
  /** Page-local top of the line box. */
  yPx: number;
  /** Page-local x of the line's first glyph (indents applied). */
  xPx: number;
  /** Per-item justified stretch-interval ends (null on unjustified lines) —
   *  the same intervals the painter stretches to. */
  intervals: number[] | null;
  /** Collapsed-char range within the paragraph's full text. */
  startChar: number;
  endChar: number;
  /** Per-item whitespace the packer trimmed ahead of the item (same indexes
   *  as `line.items`) — the gap characters this line owns in the collapsed
   *  space. */
  spaces: number[];
  /** Per-item gap start x (line-local; null when the item follows an atom or
   *  nothing) — the left caret boundary of the trimmed whitespace. */
  gapStarts: (number | null)[];
}

interface ParaEntry {
  page: number;
  para: LaidOutParagraph;
  /** PM position just inside the textblock (before its first child). */
  innerPos: number;
  node: PmNode;
  lines: LineEntry[];
  /** Total collapsed chars (concatenated line item texts). */
  chars: number;
  /** The paragraph's concatenated run text and the walk cursor into it —
   *  the laid items match it in order, consuming the whitespace pretext
   *  trimmed into the gaps (the same walk the painted space dots run, so a
   *  caret boundary, a selection edge and a dot share one lattice). */
  fullText: string;
  srcCursor: number;
}

const measureCanvas: HTMLCanvasElement | null =
  typeof document !== "undefined" ? document.createElement("canvas") : null;

/** A PM inline atom the layout projects as a laid box (an embedded picture,
 *  a math placeholder) — Word's "in line with text" graphic. It owns one
 *  selectable slot in the collapsed-char space. Atoms without a laid box
 *  (breaks, tabs — also synthesized by the projection for numbering, paged
 *  runs) stay outside it. */
function boxedInline(node: PmNode): boolean {
  return node.type.name === "image" || node.type.name === "inlinePassthrough";
}

/** The caret band of a line with no measurable text (a picture line, an
 *  empty paragraph): the line box itself, so a caret parked beside an
 *  embedded picture stays visible instead of collapsing to 2px. */
function textlessBand(line: LineEntry, pad: number): { yPx: number; heightPx: number } {
  return {
    yPx: line.yPx + pad,
    heightPx: Math.max(line.line.naturalPx, line.line.heightPx, 2),
  };
}

/** A cell's full grid box, page-local — Word's cell highlight covers the
 *  whole grid slot (insets included), not the text lines inside it. */
interface CellBoxDraft {
  page: number;
  xPx: number;
  yPx: number;
  widthPx: number;
  heightPx: number;
  /** The cell's first laid paragraph — the zip's bridge to a PM position. */
  head: LaidOutParagraph | null;
}

/** A table's placed extent with its column/row boundaries (zone-local) — the
 *  row/column selection bars and the select-all grip hit-test against it. */
export interface TableZone {
  page: number;
  xPx: number;
  yPx: number;
  widthPx: number;
  heightPx: number;
  /** Column left edges + the right rim (nCols + 1, table-local = zone-local). */
  colEdges: number[];
  /** Row top edges + the bottom rim (nRows + 1). */
  rowEdges: number[];
}

/** The cell content stack's first paragraph block — the position the PM zip
 *  pairs. A cell opening with a nested table (rare) has none. */
function firstParaOf(
  stack: readonly { block: import("@docen/layout").LaidOutBlock }[],
): LaidOutParagraph | null {
  const item = stack[0];
  if (!item) return null;
  if (item.block.kind === "paragraph") return item.block;
  if (item.block.kind === "group") return firstParaOf(item.block.children);
  return null;
}

function collectLayoutParas(
  items: readonly {
    yPx: number;
    xPx?: number;
    block: import("@docen/layout").LaidOutBlock;
  }[],
  page: number,
  x: number,
  y: number,
  out: { page: number; para: LaidOutParagraph; xPx: number; yPx: number }[],
  boxes: CellBoxDraft[] | null,
  zones: TableZone[] | null,
): void {
  for (const item of items) {
    const b = item.block;
    // xPx is the item's column left edge (w:cols sections, absent single-column).
    const bx = x + (item.xPx ?? 0);
    const by = y + item.yPx;
    switch (b.kind) {
      case "paragraph":
        out.push({ page, para: b, xPx: bx, yPx: by });
        break;
      case "group":
        collectLayoutParas(b.children, page, bx, by, out, boxes, zones);
        break;
      case "table": {
        // The painter's own walk (tableGridOf), including the w:jc placement
        // it applies before that walk (paintTable adds offsetXPx there):
        // content anchors at the placed origin — column left + insets, row
        // top + insets + vertical-align offset — exactly where painting puts
        // it. Skipping the placement shifted every click and highlight in a
        // centered/right table a full offset left of the glyphs.
        const gx = bx + (b.offsetXPx ?? 0);
        const grid = tableGridOf(b);
        for (const p of grid.cells) {
          collectLayoutParas(
            p.cell.stack,
            page,
            gx + p.contentXPx,
            by + p.contentYPx,
            out,
            boxes,
            zones,
          );
          // The full grid slot: spanned column/row edges (insets included —
          // the highlight is the cell, not its content). Keyed later by the
          // first paragraph's PM position.
          const left = grid.colX[p.col]!;
          const top = grid.rowY[p.row]!;
          boxes?.push({
            page,
            xPx: gx + left,
            yPx: by + top,
            widthPx: grid.colX[p.col + p.spanW]! - left,
            heightPx: grid.rowY[p.row + p.spanH]! - top,
            head: firstParaOf(p.cell.stack),
          });
        }
        zones?.push({
          page,
          xPx: gx,
          yPx: by,
          widthPx: grid.colX[grid.colX.length - 1]!,
          heightPx: grid.rowY[grid.rowY.length - 1]!,
          colEdges: [...grid.colX],
          rowEdges: [...grid.rowY],
        });
        break;
      }
      default:
        break;
    }
  }
}

export class CaretMap {
  readonly valid: boolean;
  private readonly paras: ParaEntry[] = [];
  private readonly lines: LineEntry[] = [];
  /** Cell grid boxes keyed by the cell's PM position — the cell selection's
   *  whole-slot highlight, and the geometry the selection bars resolve a
   *  clicked column/row back to a PM cell from. One cell may paint in several
   *  places: a repeated tblHeader band's cells ARE the original band's cells
   *  (splitLaid's band copies keep the row references), so every placement
   *  keeps its box and a full-table selection highlights them all — Word's
   *  repeated header lights up on every continuation page too. Cells whose
   *  first paragraph never zipped (nested-table openings) stay unmapped. */
  readonly cellBoxes = new Map<number, SelectionRect[]>();
  /** Every placed table's extent + column/row edges — the selection bars. */
  readonly tableZones: TableZone[] = [];

  constructor(
    pages: readonly FlowPage[],
    doc: PmNode,
    private readonly originOf: (page: number) => { contentLeftPx: number; contentTopPx: number },
  ) {
    // Layout side, document order. Multi-section documents give each page its
    // own content origin — the section the page belongs to.
    const laid: { page: number; para: LaidOutParagraph; xPx: number; yPx: number }[] = [];
    const boxDrafts: CellBoxDraft[] = [];
    pages.forEach((p, page) => {
      const origin = originOf(page);
      collectLayoutParas(
        p.items,
        page,
        origin.contentLeftPx,
        origin.contentTopPx,
        laid,
        boxDrafts,
        this.tableZones,
      );
    });
    // PM side: every textblock, document order (descendants).
    const tbs: { node: PmNode; pos: number }[] = [];
    doc.descendants((node, pos) => {
      if (node.isTextblock) {
        tbs.push({ node, pos });
        return false;
      }
      return true;
    });
    // Rendered text of a laid paragraph (inline runs concatenated) — the
    // resync signal when the two sides drift apart.
    const norm = (s: string): string => s.replace(/\s+/g, "");
    const laidText = (para: LaidOutParagraph): string => {
      let text = "";
      for (const line of para.lines) {
        for (const item of line.items) if (item.kind === "text") text += item.text;
      }
      return norm(text);
    };
    // The paragraph's run text — the source the gap walk matches items
    // against (its whitespace is what pretext trimmed into the gaps). Every
    // non-text inline marks its source position with one U+FFFC placeholder,
    // so whitespace between two atoms attributes to the atom it precedes
    // (pictures edge to edge are selectable, the collapsed run included).
    const runTextOf = (para: LaidOutParagraph): string => {
      let text = "";
      for (const inline of para.inline) {
        text += inline.kind === "text" ? inline.text : "￼";
      }
      return text;
    };
    // True zip: walk the laid blocks, consuming one PM textblock per logical
    // paragraph. Three drifts must not invalidate the whole map:
    // 1. PM textblocks the flow never lays (floating-table cells paint in the
    //    scene without flow items) stay unmapped — their positions render a
    //    caret nowhere instead of killing editing document-wide.
    // 2. Render-only laid paragraphs (repeated table headers on continuation
    //    pages, a TOC entry laid from its cached options while the PM field
    //    content stays empty) pair by position instead.
    // 3. A RUN of render-only paragraphs longer than the local scans (a
    //    21-entry TOC lays as 21 paragraphs over one empty field paragraph)
    //    resyncs on the laid-side anchor: where the current textblock's text
    //    next appears in the laid list, skipping the run in between.
    const laidTexts = laid.map((l) => laidText(l.para));
    const laidRuns = new Map<string, number[]>();
    laidTexts.forEach((t, idx) => {
      const run = laidRuns.get(t);
      if (run) run.push(idx);
      else laidRuns.set(t, [idx]);
    });
    this.valid = true;
    let j = 0;
    let i = 0;
    while (i < laid.length) {
      const entry = laid[i]!;
      const para = entry.para;
      const prev = this.paras[this.paras.length - 1];
      const continuation = prev !== undefined && prev.para.inline === para.inline;
      if (continuation && prev) {
        this.appendLines(prev, entry);
        i++;
        continue;
      }
      if (j >= tbs.length) {
        // Render-only tail (a repeated table header row on a continuation
        // page): nothing left to pair with — skip the laid block.
        i++;
        continue;
      }
      const here = laidTexts[i]!;
      const there = norm(tbs[j]!.node.textContent);
      if (here !== there) {
        // Text disagrees at this position. First suspect a PM-side gap: this
        // laid paragraph's text appears further ahead in the textblock list —
        // the textblocks in between (floating-table cells, passthrough runs)
        // are never laid, so skip them unmapped.
        let gap = 0;
        // An EMPTY laid text never gap-skips: "" matches any empty textblock
        // ahead, so the "match" is noise and the skipped range swallows real
        // paragraphs (it must fall through to the anchor resync below). A
        // unique non-empty text pairs with its textblock however far ahead it
        // sits (a floating table parks dozens of never-laid cell textblocks
        // between two laid paragraphs); a text laying more than once (TOC
        // entry + its heading) keeps the short window — the park path below
        // owns those copies.
        const reach =
          here === "" ? 0 : (laidRuns.get(here)?.length ?? 0) === 1 ? tbs.length - j - 1 : 6;
        for (let k = 1; k <= reach; k++) {
          if (here === norm(tbs[j + k]!.node.textContent)) {
            gap = k;
            break;
          }
        }
        if (gap > 0) {
          // But not when this text lays again later (a TOC entry whose
          // heading also lays in the body — the entry's text equals its
          // heading's once the page number drops out): this block is the
          // render-only copy. Park it on the current EMPTY textblock
          // without consuming it — the entries share the field paragraph's
          // position and stay clickable — and let the later laid copy pair
          // with its heading.
          const runs = laidRuns.get(here);
          if (runs && runs[runs.length - 1] !== i && there === "") {
            const { node, pos } = tbs[j]!;
            const paraEntry: ParaEntry = {
              page: entry.page,
              para,
              innerPos: pos + 1,
              node,
              lines: [],
              chars: 0,
              fullText: runTextOf(para),
              srcCursor: 0,
            };
            this.paras.push(paraEntry);
            this.appendLines(paraEntry, entry);
            i++;
            continue;
          }
          j += gap;
        } else {
          const next = i + 1 < laid.length ? laidTexts[i + 1]! : null;
          const anchor = laidRuns.get(there)?.find((idx) => idx > i);
          if (next != null && next === there) {
            // Laid-side gap: the NEXT laid block pairs with this textblock,
            // so this one is render-only — skip it without consuming.
            i++;
            continue;
          }
          if (anchor != null && there !== "") {
            // A render-only RUN: skip straight to the laid block carrying
            // this textblock's text. Without it every TOC entry after the
            // first pair-as-is'd onto a real body paragraph and every click
            // below selected the wrong paragraph's text. Empty `there` (the
            // blank paragraphs after a TOC field) stays on pair-as-is — an
            // empty anchor is never a reliable resync point.
            i = anchor;
            continue;
          }
          // Otherwise a legal same-position drift (a TOC entry laid from its
          // cached text over an empty field-content paragraph) — pair as is.
        }
      }
      const { node, pos } = tbs[j++]!;
      const paraEntry: ParaEntry = {
        page: entry.page,
        para,
        innerPos: pos + 1,
        node,
        lines: [],
        chars: 0,
        fullText: runTextOf(para),
        srcCursor: 0,
      };
      this.paras.push(paraEntry);
      this.appendLines(paraEntry, entry);
      i++;
    }
    // The cell grid boxes: keyed by each cell's PM position — the zip pairs
    // the cell's first paragraph (innerPos), and the cell node sits two
    // positions up (cell pos → first-child pos +1 → textblock innerPos +1).
    const innerOf = new Map<LaidOutParagraph, number>();
    for (const p of this.paras) innerOf.set(p.para, p.innerPos);
    for (const draft of boxDrafts) {
      const inner = draft.head ? innerOf.get(draft.head) : undefined;
      if (inner == null) continue;
      const box: SelectionRect = {
        page: draft.page,
        xPx: draft.xPx,
        yPx: draft.yPx,
        widthPx: draft.widthPx,
        heightPx: draft.heightPx,
      };
      const boxes = this.cellBoxes.get(inner - 2);
      if (boxes) boxes.push(box);
      else this.cellBoxes.set(inner - 2, [box]);
    }
  }

  /** Append one laid block's lines to a ParaEntry (startChar continues from
   *  the accumulated count; a continuation block's first line is not the
   *  paragraph's first line — no first-line indent re-application). */
  private appendLines(
    paraEntry: ParaEntry,
    entry: { page: number; para: LaidOutParagraph; xPx: number; yPx: number },
  ): void {
    const para = entry.para;
    let startChar = paraEntry.chars;
    para.lines.forEach((line) => {
      // The painter's own origin sum — the line carries its first-line flag,
      // so no line-index guessing.
      const xPx = entry.xPx + lineOriginXPx(para, line);
      let chars = 0;
      const spaces: number[] = [];
      const gapStarts: (number | null)[] = [];
      let prevTextEnd: number | null = null;
      // The previous item's laid end edge regardless of kind — the physical
      // left boundary of a gap whose text predecessor was an atom.
      let prevItemEnd: number | null = null;
      // UTF-16 units — the PM side (posOfChar/charOfPos) counts
      // textContent.length, so the collapsed-char space must too; counting
      // code points drifted every boundary after an astral char. The gap
      // walk counts back the whitespace pretext trimmed ahead of each item
      // (those characters exist in the PM text — a doc position, a click
      // boundary and a painted space dot must agree on them).
      const gaps = lineSpaceGaps(line, paraEntry.fullText, paraEntry.srcCursor);
      if (gaps.matched) paraEntry.srcCursor = gaps.next;
      line.items.forEach((item, itemIndex) => {
        if (item.kind !== "text") {
          // The trimmed whitespace ahead of the item is its gap in the
          // collapsed space too (Word's inline graphic is selectable edge to
          // edge, the collapsed run around it included). A boxed inline (an
          // embedded picture, a math placeholder) additionally owns one
          // character slot; breaks and tabs count their gap only (a break IS
          // the line split; a tab can also be synthesized by the projection
          // for numbering).
          const gap = gaps.matched ? gaps.spaces[itemIndex]! : 0;
          chars += gap;
          spaces.push(gap);
          gapStarts.push(prevTextEnd ?? prevItemEnd);
          if (item.kind === "picture" || item.kind === "math") chars += 1;
          prevTextEnd = null;
          prevItemEnd = item.xPx + item.widthPx;
          return;
        }
        // A synthetic item (a list marker) paints but has no document-model
        // character behind it — it stays outside the PM offset space, so it
        // neither advances the char count nor participates in the gap
        // lattice (lineSpaceGaps still consumed its text above, keeping the
        // walk over the layout-side fullText aligned).
        if (item.synthetic) {
          spaces.push(0);
          gapStarts.push(null);
          return;
        }
        const gap = gaps.matched ? gaps.spaces[itemIndex]! : 0;
        chars += gap + item.text.length;
        spaces.push(gap);
        gapStarts.push(prevTextEnd ?? prevItemEnd);
        prevTextEnd = item.xPx + item.widthPx;
        prevItemEnd = prevTextEnd;
      });
      const lineEntry: LineEntry = {
        page: entry.page,
        para,
        owner: paraEntry,
        line,
        yPx: entry.yPx + line.yPx,
        xPx,
        intervals: justifiedIntervals(line),
        startChar,
        endChar: startChar + chars,
        spaces,
        gapStarts,
      };
      paraEntry.lines.push(lineEntry);
      this.lines.push(lineEntry);
      startChar += chars;
    });
    paraEntry.chars = startChar;
  }

  /** The PM position just inside the textblock a laid paragraph paired with
   *  (null when the paragraph never paired — render-only or unmapped). The
   *  drawing selection resolves its hit box's host paragraph through this;
   *  a split paragraph's continuation blocks all resolve to the same entry. */
  posOfPara(para: LaidOutParagraph): number | null {
    return this.paras.find((p) => p.lines.some((l) => l.para === para))?.innerPos ?? null;
  }

  /** A click's page-local coordinates → the nearest doc position. Clamping
   *  (drag extends) drops the distance cap: the overshoot past a line's band
   *  resolves to that nearest line, so dragging below the last line selects
   *  through its end instead of stalling. */
  posAtPoint(page: number, x: number, y: number, clamp = false): number | null {
    // Lines at the best (smallest) vertical distance — one band per click.
    let bestDist = Infinity;
    const band: LineEntry[] = [];
    for (const entry of this.lines) {
      if (entry.page !== page) continue;
      const within = y >= entry.yPx && y <= entry.yPx + entry.line.heightPx;
      const dist = within
        ? 0
        : Math.min(Math.abs(y - entry.yPx), Math.abs(y - (entry.yPx + entry.line.heightPx)));
      if (dist > 40 && !clamp) continue;
      if (dist < bestDist) {
        bestDist = dist;
        band.length = 0;
      }
      if (dist === bestDist) band.push(entry);
    }
    if (!band.length) return null;
    // The line the x falls inside. A point in the gutter BETWEEN two lines
    // (column gap, cell spacing) belongs to the line it just LEFT — columns
    // share one y band, and x proximity would fling a drag overshooting a
    // column's right edge clear across the gutter into the next column's
    // text. Document order (band is in it) clamps to the text behind the
    // gutter instead — Word's drag behavior.
    const inside = band.find((entry) => {
      const items = entry.line.items;
      const last = items[items.length - 1];
      const right = entry.xPx + (last ? last.xPx + last.widthPx : 0);
      return x >= entry.xPx && x <= right;
    });
    // Outside every line's span: the gutter's LEFT line (last line whose left
    // edge is at/behind x) — an overshoot past a column's right edge clamps
    // to that column's line end; before the first line clamps to its start.
    let hit = inside ?? null;
    if (!hit) {
      for (const entry of band) {
        if (entry.xPx <= x) hit = entry;
        else break;
      }
      hit ??= band[0]!;
    }
    return this.posInLine(hit, x);
  }

  /** The position one line above/below a position's line, at the same
   *  character column (clamped to the target line's length — the Word goal
   *  column; a pixel target would drift across differently stretched
   *  justified lines). Null at the paragraph's vertical edge. */
  posVertical(pos: number, dir: -1 | 1): number | null {
    const located = this.locate(pos);
    if (!located) return null;
    const lines = located.entry.lines;
    const target = lines[lines.indexOf(located.line) + dir];
    if (!target) return null;
    const col = Math.min(
      located.offset - located.line.startChar,
      target.endChar - target.startChar,
    );
    return this.posOfChar(located.entry, target.startChar + col);
  }

  /** The vertical caret box of a line — anchored where the painter actually
   *  draws: the baseline sits `leaferBaselinePadPx` below the element top
   *  (itself at the grid-centered pad). Sizing to the runs' ink box keeps
   *  the highlight hugging the glyphs; the font-box model floated a ~0.3em
   *  gap above them. */
  private bandOf(line: LineEntry): { yPx: number; heightPx: number } {
    const pad = gridPadOf(line.line);
    const ctx = measureCanvas?.getContext("2d");
    if (!ctx) return textlessBand(line, pad);
    let top = Infinity;
    let bottom = -Infinity;
    for (const item of line.line.items) {
      if (item.kind !== "text") continue;
      const inline = line.para.inline[item.inlineIndex];
      if (inline?.kind !== "text") continue;
      // The slot test the layout's own measurement used (isCjkCodeUnit over
      // the engine's CJK ranges) — the band hugs glyphs measured AND painted
      // in the same face.
      const font = cssFontOf(
        inline.style,
        familyOfSlot(inline.style.family, isCjkCodeUnit(item.text, 0)),
      );
      ctx.font = font;
      // The painter's own baseline: a vertAlign run paints at the scaled
      // size on a shifted baseline (vertAlignedSizePx in cssFontOf above,
      // vertAlignBaselineShiftPx in the shared measure module) — the band
      // must anchor there too, or a footnote reference's highlight rides
      // below its glyphs.
      const baseline =
        line.yPx +
        pad +
        vertAlignBaselineShiftPx(inline.style) +
        leaferBaselinePadPx(vertAlignedSizePx(inline.style));
      // The item's own ink box (first graphemes carry its script's shape); the
      // deepest run's descent and highest run's ascent bound the highlight.
      const metrics = ctx.measureText(Array.from(item.text).slice(0, 8).join(""));
      top = Math.min(top, baseline - metrics.actualBoundingBoxAscent);
      bottom = Math.max(bottom, baseline + metrics.actualBoundingBoxDescent);
    }
    if (!Number.isFinite(top)) {
      return textlessBand(line, pad);
    }
    return { yPx: top, heightPx: Math.max(bottom - top, 2) };
  }

  /** The selection rectangles for a range — Word's highlight model: every
   *  fully crossed line spans to the wrap's right edge (not the last glyph),
   *  the end line stops at the selection's last boundary, heights are the
   *  full line box (contiguous down the paragraph, covering the pitch gap),
   *  and a paragraph's trailing spacing highlights when the next paragraph is
   *  selected too. Empty paragraphs show a caret-width block. */
  selectionRects(from: number, to: number): SelectionRect[] {
    const rects: SelectionRect[] = [];
    this.paras.forEach((entry, paraIndex) => {
      const start = Math.max(from, entry.innerPos);
      const end = Math.min(to, entry.innerPos + entry.node.content.size);
      // An empty paragraph's position range is degenerate ([innerPos,
      // innerPos]) — it is selected whenever the range crosses that position.
      const emptyBlock = entry.node.content.size === 0;
      if (start > end || (start === end && !(emptyBlock && from <= entry.innerPos))) return;
      const offA = this.charOfPos(entry, start);
      const offB = this.charOfPos(entry, end);
      const next = this.paras[paraIndex + 1];
      const nextSelected =
        next !== undefined &&
        Math.max(from, next.innerPos) < Math.min(to, next.innerPos + next.node.content.size);
      for (const [li, line] of entry.lines.entries()) {
        const empty = line.startChar === line.endChar;
        // A line the PM side maps no chars into (offA === offB) belongs to a
        // paragraph painted from cached options over an empty field paragraph
        // (a TOC entry): there is nothing to intersect — the selection
        // crossing the paragraph highlights the whole painted line below.
        if (empty) {
          if (offA > line.startChar || offB < line.startChar) continue;
        } else if (offA !== offB && (line.endChar <= offA || line.startChar >= offB)) {
          continue;
        }
        // Line-box geometry (the layout's own pitch) keeps multi-line
        // highlights contiguous — the caret's ink band would fragment them.
        // The line-box floor keeps the height sane across a column split's
        // line-y rewind: a tail block restarts at the right column's top on
        // the SAME page, so the next line's y can sit ABOVE the current one
        // and taking it raw would flip the height negative (invisible).
        const boxBottom = line.yPx + line.line.heightPx;
        const nextLine = entry.lines[li + 1];
        const nextFirst = next?.lines[0];
        const bottom = nextLine
          ? Math.max(nextLine.yPx, boxBottom)
          : nextSelected && nextFirst && nextFirst.page === line.page
            ? // The paragraph gap (after+before spacing) belongs to the
              // selection once the next paragraph is in it too.
              Math.max(nextFirst.yPx, boxBottom)
            : boxBottom;
        if (empty) {
          // An atom line (a picture) paints no characters but still owns a
          // box — highlight the items' span; a textless paragraph keeps the
          // caret-width stub.
          let first: number | null = null;
          let last = 0;
          for (const item of line.line.items) {
            if (item.kind === "text") continue;
            if (first == null || item.xPx < first) first = item.xPx;
            last = Math.max(last, item.xPx + item.widthPx);
          }
          rects.push({
            page: line.page,
            xPx: line.xPx + (first ?? 0),
            yPx: line.yPx,
            widthPx:
              first == null ? Math.min(8, line.line.maxWidthPx ?? 8) : Math.max(last - first, 2),
            heightPx: bottom - line.yPx,
          });
          continue;
        }
        if (offA === offB) {
          // From the line's own content start (a leading atom's item x carries
          // the centering/hang offset — line.xPx alone would sit in the
          // margin), through the line's packed width.
          const startXPx = this.xOfChar(line, line.startChar);
          rects.push({
            page: line.page,
            xPx: startXPx,
            yPx: line.yPx,
            widthPx: Math.max((line.line.maxWidthPx ?? 0) - (startXPx - line.xPx), 2),
            heightPx: bottom - line.yPx,
          });
          continue;
        }
        const fromOff = Math.max(offA, line.startChar);
        // A line's highlight reaches the wrap edge when the selection passes
        // its end — but the document's final boundary stops at the last glyph
        // (Word: Ctrl+A's last line is not stretched).
        const endPos = this.posOfChar(entry, line.endChar);
        const coversEnd =
          offB > line.endChar || (offB === line.endChar && (to === endPos || nextSelected));
        const left = this.xOfChar(line, fromOff);
        const right = coversEnd
          ? line.xPx + (line.line.maxWidthPx ?? 0) + (line.line.hangPx ?? 0)
          : this.xOfChar(line, Math.min(offB, line.endChar));
        rects.push({
          page: line.page,
          xPx: left,
          yPx: line.yPx,
          widthPx: Math.max(right - left, 2),
          heightPx: bottom - line.yPx,
        });
      }
    });
    return rects;
  }

  /** A cell selection's highlight — every selected cell as one full grid box
   *  per placement (Word highlights the cell, insets included, and a repeated
   *  tblHeader band lights up on every continuation page too). Cells the zip
   *  never paired stay unhighlighted. */
  cellSelectionRects(selection: {
    forEachCell(f: (node: PmNode, pos: number) => void): void;
  }): SelectionRect[] {
    const rects: SelectionRect[] = [];
    selection.forEachCell((_node, pos) => {
      for (const box of this.cellBoxes.get(pos) ?? []) rects.push(box);
    });
    return rects;
  }

  /** The innermost table zone containing a page-local point, if any — the
   *  selection bars' hover test. `pad` widens the test past the zone's edges
   *  (the bars hover just OUTSIDE the table). Nested tables collect after
   *  their parent, so the last match is the innermost. */
  tableZoneAt(page: number, x: number, y: number, pad = 0): TableZone | null {
    let hit: TableZone | null = null;
    for (const z of this.tableZones) {
      if (
        z.page === page &&
        x >= z.xPx - pad &&
        x <= z.xPx + z.widthPx + pad &&
        y >= z.yPx - pad &&
        y <= z.yPx + z.heightPx + pad
      ) {
        hit = z;
      }
    }
    return hit;
  }

  /** The x-resolved position within one line (round to the nearest boundary). */
  private posInLine(entry: LineEntry, x: number): number | null {
    let char = entry.startChar;
    // Initialized via a cast so TS keeps the declared union at the read site
    // (closure writes aren't tracked and a plain null narrows to never).
    let bestOffset = undefined as { pos: number; dist: number } | undefined;
    const push = (xAt: number, pos: number): void => {
      const dist = Math.abs(xAt - x);
      if (!bestOffset || dist < bestOffset.dist) bestOffset = { pos, dist };
    };
    push(entry.xPx, this.posOfChar(entry.owner, char));
    for (const [itemIndex, item] of entry.line.items.entries()) {
      const boxed = item.kind === "picture" || item.kind === "math";
      if (!boxed && item.kind !== "text") continue;
      // A synthetic item (a list marker) carries no PM characters — its
      // glyphs sit outside the offset space, so skip both its gap and its
      // grapheme boundaries.
      if (item.kind === "text" && item.synthetic) continue;
      // The trimmed gap ahead of the item: its characters' left boundaries
      // share the gap the space dots center in (the previous item's laid end
      // → this item's x, evenly split), so a click inside the gap lands on
      // the space char the PM side knows about. Boxed inlines share the
      // lattice — their gap is the collapsed run between them and the
      // previous content.
      const gap = entry.spaces[itemIndex]!;
      if (gap > 0) {
        const gs = entry.gapStarts[itemIndex]!;
        const span = gs != null ? item.xPx - gs : 0;
        for (let k = 0; k < gap; k++) {
          push(
            entry.xPx + (gs != null && span > 0 ? gs + (span * k) / gap : item.xPx),
            this.posOfChar(entry.owner, char),
          );
          char++;
        }
      }
      // A boxed inline offers its two edges as caret boundaries — clicking
      // left of it lands before the box, right of it after (Word's inline
      // graphic is one character wide); clicking ON it selects the drawing
      // before this map is ever asked.
      if (boxed) {
        push(entry.xPx + item.xPx, this.posOfChar(entry.owner, char));
        char += 1;
        push(entry.xPx + item.xPx + item.widthPx, this.posOfChar(entry.owner, char));
        continue;
      }
      const glyphs = this.glyphLayout(entry, itemIndex);
      if (!glyphs) continue;
      for (let g = 0; g < glyphs.layout.lens.length; g++) {
        // layout.xs[g] is the g-th grapheme's LEFT edge — the boundary with
        // exactly `char` collapsed UTF-16 units before it. Push it against
        // the pre-increment char; pairing it with char+1 shifted every
        // boundary one position right and clicks landed a full character off.
        push(glyphs.base + glyphs.layout.xs[g]!, this.posOfChar(entry.owner, char));
        char += glyphs.layout.lens[g]!;
      }
    }
    // The line-end edge: clicking past the last glyph lands here (Word's
    // line-end click puts the caret at this line's end).
    push(this.xOfChar(entry, entry.endChar), this.posOfChar(entry.owner, char));
    // Empty lines push nothing — fall back to the line-start position.
    return bestOffset?.pos ?? this.posOfChar(entry.owner, entry.endChar);
  }

  /** One text item's glyph placement anchored at the line — the shared
   *  itemGlyphLayout model (the exact distribution the painter's Text
   *  renders, Leafer's CharLayout) with the item's stretch/compress
   *  interval: a justified item's interval end, or on a squeezed line the
   *  item's own right edge (the painter runs both-letter at negative
   *  slack — the same uniform per-grapheme delta as justification). */
  private glyphLayout(
    entry: LineEntry,
    itemIndex: number,
  ): { layout: ItemGlyphLayout; base: number; end: number | undefined } | null {
    const item = entry.line.items[itemIndex]!;
    if (item.kind !== "text") return null;
    const inline = entry.para.inline[item.inlineIndex];
    if (inline?.kind !== "text") return null;
    const end =
      entry.intervals?.[itemIndex] ??
      (entry.line.advanceScale != null ? item.xPx + item.widthPx : undefined);
    const font = cssFontOf(
      inline.style,
      familyOfSlot(inline.style.family, isCjkCodeUnit(item.text, 0)),
    );
    return {
      layout: itemGlyphLayout(
        item.text,
        font,
        inline.style.letterSpacingPx,
        end != null ? end - item.xPx : undefined,
      ),
      base: entry.xPx + item.xPx,
      end,
    };
  }

  /** Collapsed-char offset → doc position (walk the textblock's children). */
  private posOfChar(entry: ParaEntry, char: number): number {
    let pos = entry.innerPos;
    let remaining = char;
    let lastAtomStart = -1;
    entry.node.content.forEach((child) => {
      if (remaining < 0) return;
      if (child.isText) {
        if (remaining <= child.textContent.length) {
          pos += remaining;
          remaining = -1;
        } else {
          remaining -= child.textContent.length;
          pos += child.nodeSize;
        }
      } else if (boxedInline(child)) {
        // A boxed inline owns one offset slot: crossing it consumes that
        // slot, so the offset past it lands right after the box. Offset 0
        // relative to it stays before it (the walk just stops here).
        if (remaining > 0) {
          pos += child.nodeSize;
          remaining -= 1;
        }
      } else if (remaining > 0) {
        // An atom with no laid box carries no collapsed chars — step over it
        // while offsets remain to reach the text after it.
        lastAtomStart = pos;
        pos += child.nodeSize;
      }
    });
    // Ghost chars — a rendered line longer than its paragraph's PM content
    // (a TOC entry paints the cached entry text over a single field atom)
    // — have no position of their own: fold them back before the last atom
    // instead of onto the paragraph's far edge, or every selection anchored
    // inside the line clipped the paragraph to an empty range and the line
    // highlighted nothing (the next line lit up instead).
    if (remaining > 0 && lastAtomStart >= 0) return lastAtomStart;
    return pos;
  }

  /** The paragraph/line/collapsed-offset a doc position lands in. */
  private locate(pos: number): { entry: ParaEntry; offset: number; line: LineEntry } | null {
    const entry = this.paras.find(
      (p) => pos >= p.innerPos && pos <= p.innerPos + p.node.content.size,
    );
    if (!entry) return null;
    const offset = this.charOfPos(entry, pos);
    // A non-empty line's end offset belongs to that line (clicking past the
    // last glyph must place the caret at THIS line's end, not roll into the
    // next line's coordinate space); an empty line hands its offset to the
    // next one.
    const line =
      entry.lines.find(
        (l) => offset >= l.startChar && offset <= l.endChar && l.endChar > l.startChar,
      ) ?? entry.lines[entry.lines.length - 1];
    return line ? { entry, offset, line } : null;
  }

  /** A doc position → the caret's page-local box (null when unmappable). */
  caretRect(pos: number): CaretRect | null {
    const located = this.locate(pos);
    if (!located) return null;
    const { offset, line } = located;
    const band = this.bandOf(line);
    return {
      // The line's page — a split paragraph's ParaEntry.page stays at its
      // first block's page, the caret belongs where the line actually lays.
      page: line.page,
      xPx: this.xOfChar(line, offset),
      yPx: band.yPx,
      heightPx: band.heightPx,
    };
  }

  /** A doc position's line-start/-end positions (null when unmappable). */
  lineEdges(pos: number): { home: number; end: number } | null {
    const located = this.locate(pos);
    if (!located) return null;
    return {
      home: this.posOfChar(located.entry, located.line.startChar),
      end: this.posOfChar(located.entry, located.line.endChar),
    };
  }

  /** The first doc position rendered on a page (null when the page has no
   *  text lines). */
  firstPosOfPage(page: number): number | null {
    const line = this.lines.find((l) => l.page === page);
    return line ? this.posOfChar(line.owner, line.startChar) : null;
  }

  /** Doc position → collapsed-char offset in its paragraph. */
  private charOfPos(entry: ParaEntry, pos: number): number {
    let char = 0;
    let cursor = entry.innerPos;
    let resolved: number | null = null;
    entry.node.content.forEach((child) => {
      if (resolved !== null) return;
      if (child.isText) {
        const end = cursor + child.textContent.length;
        if (pos <= end) {
          resolved = char + (pos - cursor);
          return;
        }
        char += child.textContent.length;
        cursor = end;
      } else if (boxedInline(child)) {
        // A boxed inline owns one offset slot: positions before it map to
        // the running offset, positions after it to the slot past it.
        const start = cursor;
        cursor += child.nodeSize;
        if (pos <= start) {
          resolved = char;
          return;
        }
        char += 1;
        if (pos <= cursor) {
          resolved = char;
          return;
        }
      } else {
        // Atoms with no laid box (breaks, tabs, paged runs) carry no
        // collapsed chars — a caret beside one maps to the running offset.
        cursor += child.nodeSize;
        if (pos <= cursor) {
          resolved = char;
          return;
        }
      }
    });
    return resolved ?? char;
  }

  /** Collapsed-char offset → page-local x within its line. */
  private xOfChar(line: LineEntry, offset: number): number {
    let char = line.startChar;
    for (const [itemIndex, item] of line.line.items.entries()) {
      const boxed = item.kind === "picture" || item.kind === "math";
      if (!boxed && item.kind !== "text") continue;
      // A synthetic item (a list marker) sits outside the offset space — its
      // glyphs paint before the paragraph's own characters and never answer
      // a boundary query.
      if (item.kind === "text" && item.synthetic) continue;
      // The trimmed gap ahead of the item: the boundary rides the gap's even
      // split — the same lattice the space dots center in. Boxed inlines
      // share it (their gap is the collapsed run before the box).
      const gap = line.spaces[itemIndex]!;
      if (gap > 0 && offset <= char + gap) {
        const gs = line.gapStarts[itemIndex]!;
        const span = gs != null ? item.xPx - gs : 0;
        return line.xPx + (gs != null && span > 0 ? gs + (span * (offset - char)) / gap : item.xPx);
      }
      char += gap;
      // A boxed inline's own slot maps to its left edge (the boundary before
      // it); the offset past it resolves against the following content.
      if (boxed) {
        if (offset === char) return line.xPx + item.xPx;
        char += 1;
        continue;
      }
      // Inside the item's own glyphs: the shared per-grapheme lattice.
      if (offset <= char + item.text.length) {
        const glyphs = this.glyphLayout(line, itemIndex);
        if (!glyphs) return line.xPx + item.xPx;
        // UTF-16 offset → grapheme index: an offset that splits a multi-unit
        // grapheme lands at that grapheme's left edge.
        const local = offset - char;
        let g = 0;
        let cum = 0;
        while (g < glyphs.layout.lens.length && cum + glyphs.layout.lens[g]! <= local) {
          cum += glyphs.layout.lens[g]!;
          g++;
        }
        // Past the last grapheme the boundary is the item's end edge: a
        // justified/squeezed item at its interval's end (the painter fills
        // the width), a natural item at its advance sum.
        return g >= glyphs.layout.xs.length
          ? glyphs.end != null
            ? line.xPx + glyphs.end
            : glyphs.base + glyphs.layout.endX
          : glyphs.base + glyphs.layout.xs[g]!;
      }
      char += item.text.length;
    }
    const last = line.line.items[line.line.items.length - 1];
    return line.xPx + (last ? last.xPx + last.widthPx : 0);
  }
}
