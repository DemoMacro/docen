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
  type LaidOutLine,
  type LaidOutParagraph,
  leaferWordIndices,
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
  /** Collapsed-char range within the paragraph's full text. */
  startChar: number;
  endChar: number;
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
}

const measureCanvas: HTMLCanvasElement | null =
  typeof document !== "undefined" ? document.createElement("canvas") : null;

const SEGMENTER = new Intl.Segmenter(undefined, { granularity: "grapheme" });

/** Per-grapheme advance widths of a text in a font — measured once, reused
 *  across every caret query against the same item (drag-select rescans its
 *  line on every mousemove; per-grapheme sums approximate the kerned run,
 *  caret placement only). */
const widthCache = new Map<string, number[]>();

function graphemeWidths(text: string, font: string): number[] {
  const key = `${font}\u0000${text}`;
  const cached = widthCache.get(key);
  if (cached) return cached;
  const ctx = measureCanvas?.getContext("2d");
  if (!ctx) return [];
  ctx.font = font;
  const widths: number[] = [];
  for (const { segment } of SEGMENTER.segment(text)) widths.push(ctx.measureText(segment).width);
  if (widthCache.size >= 4000) widthCache.clear();
  widthCache.set(key, widths);
  return widths;
}

/** Sum of the first `n` grapheme widths. */
function prefixWidth(widths: number[], n: number): number {
  let w = 0;
  for (let i = 0; i < n && i < widths.length; i++) w += widths[i]!;
  return w;
}

function collectLayoutParas(
  items: readonly { yPx: number; block: import("@docen/layout").LaidOutBlock }[],
  page: number,
  x: number,
  y: number,
  out: { page: number; para: LaidOutParagraph; xPx: number; yPx: number }[],
): void {
  for (const item of items) {
    const b = item.block;
    const bx = x;
    const by = y + item.yPx;
    switch (b.kind) {
      case "paragraph":
        out.push({ page, para: b, xPx: bx, yPx: by });
        break;
      case "group":
        collectLayoutParas(b.children, page, bx, by, out);
        break;
      case "table": {
        // Mirror the painter's occupancy walk (colspans shift the column
        // cursor); content anchors at the start row + cell insets.
        const colX = [0];
        for (const w of b.columnWidthsPx) colX.push(colX[colX.length - 1] + w);
        const rowY = [0];
        for (const row of b.rows) rowY.push(rowY[rowY.length - 1] + row.heightPx);
        const nCols = b.columnWidthsPx.length;
        const nRows = b.rows.length;
        const occ: (boolean | undefined)[][] = Array.from({ length: nRows }, () =>
          Array.from<boolean | undefined>({ length: nCols }),
        );
        b.rows.forEach((row, r) => {
          let col = 0;
          for (const cell of row.cells) {
            while (col < nCols && occ[r][col]) col++;
            if (col >= nCols) break;
            const spanW = Math.min(cell.colspan, nCols - col);
            const spanH = Math.min(cell.rowspan ?? 1, nRows - r);
            for (let dr = 0; dr < spanH; dr++)
              for (let dc = 0; dc < spanW; dc++) occ[r + dr][col + dc] = true;
            collectLayoutParas(
              cell.stack,
              page,
              bx + colX[col] + (cell.insets.left ?? 0),
              by + rowY[r] + (cell.insets.top ?? 0),
              out,
            );
            col += spanW;
          }
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

  constructor(
    pages: readonly FlowPage[],
    doc: PmNode,
    private readonly flow: { contentLeftPx: number; contentTopPx: number },
  ) {
    // Layout side, document order.
    const laid: { page: number; para: LaidOutParagraph; xPx: number; yPx: number }[] = [];
    pages.forEach((p, page) => {
      collectLayoutParas(p.items, page, flow.contentLeftPx, flow.contentTopPx, laid);
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
    if (laid.length < tbs.length) {
      this.valid = false;
      return;
    }
    this.valid = true;
    // Page-split paragraphs lay out as one block per page but share their
    // inline array by reference — such continuation blocks append lines to
    // the running ParaEntry instead of consuming the next textblock.
    let j = 0;
    for (const entry of laid) {
      const para = entry.para;
      const prev = this.paras[this.paras.length - 1];
      const continuation = prev !== undefined && prev.para.inline === para.inline;
      if (continuation && prev) {
        this.appendLines(prev, entry);
        continue;
      }
      if (j >= tbs.length) {
        this.valid = false;
        return;
      }
      const { node, pos } = tbs[j++]!;
      const paraEntry: ParaEntry = {
        page: entry.page,
        para,
        innerPos: pos + 1,
        node,
        lines: [],
        chars: 0,
      };
      this.paras.push(paraEntry);
      this.appendLines(paraEntry, entry);
    }
    if (j !== tbs.length) this.valid = false;
  }

  /** Append one laid block's lines to a ParaEntry (startChar continues from
   *  the accumulated count; a continuation block's first line is not the
   *  paragraph's first line — no first-line indent re-application). */
  private appendLines(
    paraEntry: ParaEntry,
    entry: { page: number; para: LaidOutParagraph; xPx: number; yPx: number },
  ): void {
    const para = entry.para;
    const first = paraEntry.lines.length === 0;
    let startChar = paraEntry.chars;
    para.lines.forEach((line, lineIndex) => {
      const xPx =
        entry.xPx +
        (para.indent?.leftPx ?? 0) +
        (first && lineIndex === 0 ? (para.indent?.firstLinePx ?? 0) : 0);
      let chars = 0;
      for (const item of line.items) if (item.kind === "text") chars += [...item.text].length;
      const lineEntry: LineEntry = {
        page: entry.page,
        para,
        owner: paraEntry,
        line,
        yPx: entry.yPx + line.yPx,
        xPx,
        startChar,
        endChar: startChar + chars,
      };
      paraEntry.lines.push(lineEntry);
      this.lines.push(lineEntry);
      startChar += chars;
    });
    paraEntry.chars = startChar;
  }

  /** A click's page-local coordinates → the nearest doc position. */
  posAtPoint(page: number, x: number, y: number): number | null {
    let best: { entry: LineEntry; dist: number; xDist: number } | null = null;
    for (const entry of this.lines) {
      if (entry.page !== page) continue;
      const within = y >= entry.yPx && y <= entry.yPx + entry.line.heightPx;
      const dist = within
        ? 0
        : Math.min(Math.abs(y - entry.yPx), Math.abs(y - (entry.yPx + entry.line.heightPx)));
      if (dist > 40) continue;
      // Table columns share one y band — x proximity breaks the tie, else
      // every click in the row lands on the first cell's paragraph.
      const items = entry.line.items;
      const last = items[items.length - 1];
      const right = entry.xPx + (last ? last.xPx + last.widthPx : 0);
      const xDist = x < entry.xPx ? entry.xPx - x : x > right ? x - right : 0;
      if (!best || dist < best.dist || (dist === best.dist && xDist < best.xDist)) {
        best = { entry, dist, xDist };
      }
    }
    if (!best) return null;
    return this.posInLine(best.entry, x);
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
   *  draws: Leafer's Text defaults to a 150% lineHeight, which puts the
   *  alphabetic baseline at 1.1 × fontSize below the element top (itself at
   *  the grid-centered pad). Sizing to the runs' ink box keeps the highlight
   *  hugging the glyphs; the font-box model floated a ~0.3em gap above them. */
  private bandOf(line: LineEntry): { yPx: number; heightPx: number } {
    const pad = line.line.grid ? Math.max(0, (line.line.heightPx - line.line.naturalPx) / 2) : 0;
    const ctx = measureCanvas?.getContext("2d");
    if (!ctx) return { yPx: line.yPx + pad, heightPx: Math.max(line.line.naturalPx, 2) };
    let top = Infinity;
    let bottom = -Infinity;
    for (const item of line.line.items) {
      if (item.kind !== "text") continue;
      const inline = line.para.inline[item.inlineIndex];
      if (inline?.kind !== "text") continue;
      const font = cssFontOf(
        inline.style,
        familyOfSlot(inline.style.family, /[一-鿿぀-ヿ가-힯]/.test(item.text[0] ?? "")),
      );
      ctx.font = font;
      const baseline = line.yPx + pad + 1.1 * inline.style.sizePx;
      // The item's own ink box (first graphemes carry its script's shape); the
      // deepest run's descent and highest run's ascent bound the highlight.
      const metrics = ctx.measureText([...item.text].slice(0, 8).join(""));
      top = Math.min(top, baseline - metrics.actualBoundingBoxAscent);
      bottom = Math.max(bottom, baseline + metrics.actualBoundingBoxDescent);
    }
    if (!Number.isFinite(top)) {
      return { yPx: line.yPx + pad, heightPx: Math.max(line.line.naturalPx, 2) };
    }
    return { yPx: top, heightPx: Math.max(bottom - top, 2) };
  }

  /** The selection rectangles for a range — one per crossed line, in the
   *  caret's geometry (same grid pad and text height). */
  selectionRects(from: number, to: number): SelectionRect[] {
    const rects: SelectionRect[] = [];
    for (const entry of this.paras) {
      const start = Math.max(from, entry.innerPos);
      const end = Math.min(to, entry.innerPos + entry.node.content.size);
      if (start >= end) continue;
      const offA = this.charOfPos(entry, start);
      const offB = this.charOfPos(entry, end);
      for (const line of entry.lines) {
        if (line.endChar <= offA || line.startChar >= offB) continue;
        const fromOff = Math.max(offA, line.startChar);
        const toOff = Math.min(offB, line.endChar);
        const band = this.bandOf(line);
        rects.push({
          page: line.page,
          xPx: this.xOfChar(line, fromOff),
          yPx: band.yPx,
          widthPx: Math.max(this.xOfChar(line, toOff) - this.xOfChar(line, fromOff), 2),
          heightPx: band.heightPx,
        });
      }
    }
    return rects;
  }

  /** The x-resolved position within one line (round to the nearest boundary). */
  private posInLine(entry: LineEntry, x: number): number | null {
    const { line } = entry;
    let char = entry.startChar;
    // Initialized via a cast so TS keeps the declared union at the read site
    // (closure writes aren't tracked and a plain null narrows to never).
    let bestOffset = undefined as { pos: number; dist: number } | undefined;
    const push = (xAt: number, pos: number): void => {
      const dist = Math.abs(xAt - x);
      if (!bestOffset || dist < bestOffset.dist) bestOffset = { pos, dist };
    };
    push(entry.xPx, this.posOfChar(entry.owner, char));
    for (const [itemIndex, item] of line.items.entries()) {
      if (item.kind !== "text") continue;
      const inline = entry.para.inline[item.inlineIndex];
      if (inline?.kind !== "text") continue;
      const font = cssFontOf(
        inline.style,
        familyOfSlot(inline.style.family, /[一-鿿぀-ヿ가-힯]/.test(item.text[0] ?? "")),
      );
      const widths = graphemeWidths(item.text, font);
      let prefix = 0;
      for (let g = 0; g < widths.length; g++) {
        char++;
        push(this.xInItem(entry, itemIndex, g, prefix, widths), this.posOfChar(entry.owner, char));
        prefix += widths[g]!;
      }
    }
    // Empty lines push nothing — fall back to the line-start position.
    return bestOffset?.pos ?? this.posOfChar(entry.owner, entry.endChar);
  }

  /** The x of one grapheme inside an item — the exact distribution the
   *  painter's justified Text applies: CJK items advance each grapheme by a
   *  uniform share of the stretch interval (Leafer "both-letter"), Latin
   *  items shift each word by its word index × the per-gap share ("both-
   *  justify"). The layout's item widths are pretext's compressed measures —
   *  scaling plain prefixes by them (the old proportional model) drifted up
   *  to a full punctuation run mid-line. */
  private xInItem(
    entry: LineEntry,
    itemIndex: number,
    graphemeIndex: number,
    plainPrefix: number,
    widths: number[],
  ): number {
    const item = entry.line.items[itemIndex]!;
    const base = entry.xPx + item.xPx + plainPrefix;
    const { line } = entry;
    if (line.justifyGapPx == null || item.kind !== "text") return base;
    // The painter's interval math: the last item stretches to maxWidth(+hang),
    // each earlier one to the next item's post-justify x.
    let intervalEnd = (line.maxWidthPx ?? 0) + (line.hangPx ?? 0);
    for (let i = line.items.length - 1; i > itemIndex; i--) intervalEnd = line.items[i]!.xPx;
    const interval = intervalEnd - item.xPx;
    const count = widths.length;
    if (count <= 1) return base;
    let natural = 0;
    for (const w of widths) natural += w;
    if (/[一-鿿぀-ヿ가-힯]/.test(item.text)) {
      // The line-end offset (index == count) sits right after the last glyph —
      // no further gap follows it (the painter's last word absorbs none).
      return base + Math.min(graphemeIndex, count - 1) * ((interval - natural) / (count - 1));
    }
    // Word gaps: a grapheme's shift is its Leafer word index × gap.
    const indices = leaferWordIndices(item.text);
    const words = (indices[count - 1] ?? 0) + 1;
    if (words <= 1) return base;
    return base + ((indices[graphemeIndex] ?? 0) * (interval - natural)) / (words - 1);
  }

  /** Collapsed-char offset → doc position (walk the textblock's children). */
  private posOfChar(entry: ParaEntry, char: number): number {
    let pos = entry.innerPos;
    let remaining = char;
    entry.node.content.forEach((child) => {
      if (child.isText) {
        if (remaining <= child.textContent.length) {
          pos += remaining;
          remaining = -1;
        } else {
          remaining -= child.textContent.length;
          pos += child.nodeSize;
        }
      } else {
        pos += child.nodeSize;
      }
    });
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
      } else {
        // Atoms carry no collapsed chars — a caret beside one maps to the
        // running offset either way.
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
      if (item.kind !== "text") continue;
      const graphemes = [...item.text];
      if (offset <= char + graphemes.length) {
        const inline = line.para.inline[item.inlineIndex];
        if (inline?.kind !== "text") return line.xPx + item.xPx;
        const font = cssFontOf(
          inline.style,
          familyOfSlot(inline.style.family, /[一-鿿぀-ヿ가-힯]/.test(item.text[0] ?? "")),
        );
        const widths = graphemeWidths(item.text, font);
        return this.xInItem(
          line,
          itemIndex,
          offset - char,
          prefixWidth(widths, offset - char),
          widths,
        );
      }
      char += graphemes.length;
    }
    const last = line.line.items[line.line.items.length - 1];
    return line.xPx + (last ? last.xPx + last.widthPx : 0);
  }
}
