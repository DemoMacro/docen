/**
 * Scene painter — walks a laid-out page (the @docen/layout result) and builds
 * the LeaferJS tree. The layout engine owns ALL geometry; painting positions
 * what it is given and never measures. Text elements carry explicit width AND
 * height — Leafer never paints an element whose height is still 0.
 */
import {
  isCjkCodeUnit,
  stackBlocks,
  TextMeasurer,
  type FlowItem,
  type LaidOutBlock,
  type LaidOutCell,
  type LaidOutLineItem,
  type LaidOutParagraph,
  type LaidOutStackItem,
  type LaidOutTable,
  type LayoutBorderEdge,
  type LayoutDrawing,
  type LayoutDrawingMember,
  type LayoutInline,
  type LayoutTextStyle,
} from "@docen/layout";
import {
  Ellipse,
  Image as LeaferImage,
  Line,
  Path as LeaferPath,
  Rect,
  Text,
  type IGroup,
} from "leafer-ui";

import type { CanvasStageContext } from "./stage";

/** OOXML prstDash tokens → dash patterns in px (line-width units, the host's
 *  preset line styles); unlisted tokens render solid. */
const PRSTDASH_PATTERN: Record<string, number[]> = {
  dot: [1, 3],
  sysDot: [1, 1],
  dash: [4, 3],
  sysDash: [4, 2],
  dashDot: [4, 3, 1, 3],
  sysDashDot: [4, 2, 1, 2],
  dashDotDot: [4, 3, 1, 3, 1, 3],
  sysDashDotDot: [4, 2, 1, 2, 1, 2],
  lgDash: [12, 3],
  lgDashDot: [12, 3, 1, 3],
  lgDashDotDot: [12, 3, 1, 3, 1, 3],
};

/** The paint context for one page — the stage context plus the page's own
 *  identity (page-number fields resolve against it). */
export interface PaintContext extends CanvasStageContext {
  pageIndex: number;
  pageCount: number;
}

/** Paint one page's flow items (block + page-content y). */
export function paintScene(tree: IGroup, items: readonly FlowItem[], ctx: PaintContext): void {
  for (const item of items) {
    paintBlock(tree, item.block, ctx.flow.contentLeftPx, ctx.flow.contentTopPx + item.yPx, ctx);
  }
}

/** Paint a pre-laid header/footer stack at its page position. */
export function paintFurnitureStack(
  tree: IGroup,
  stack: readonly LaidOutStackItem[],
  x: number,
  y: number,
  ctx: PaintContext,
): void {
  for (const item of stack) {
    paintBlock(tree, item.block, x, y + item.yPx, ctx);
  }
}

function paintBlock(
  tree: IGroup,
  block: LaidOutBlock,
  x: number,
  y: number,
  ctx: PaintContext,
): void {
  switch (block.kind) {
    case "paragraph":
      paintParagraph(tree, block, x, y, ctx);
      return;
    case "table":
      paintTable(tree, block, x, y, ctx);
      return;
    case "group":
      for (const child of block.children) {
        paintBlock(tree, child.block, x, y + child.yPx, ctx);
      }
      return;
    case "placeholder":
      paintPlaceholder(tree, block, x, y);
      return;
    case "pageBreak":
      return;
  }
}

function paintParagraph(
  tree: IGroup,
  para: LaidOutParagraph,
  x: number,
  y: number,
  ctx: PaintContext,
): void {
  // behindDoc drawings first — everything below sits on top of them.
  for (const drawing of para.drawings ?? []) {
    if (drawing.behind) paintDrawing(tree, drawing, x, y, ctx);
  }
  // w:pBdr horizontal rules (the "Education" underline shape): between the
  // text and `spacePx` below it, spanning the wrapping width.
  for (const side of ["top", "bottom"] as const) {
    const edge = para.borders?.[side];
    if (!edge || !edge.px || edge.style === "nil" || edge.style === "none") continue;
    const top = side === "top";
    const lineY = top ? y - (edge.spacePx ?? 0) : y + para.heightPx + (edge.spacePx ?? 0);
    const right = ctx.flow.contentLeftPx + ctx.flow.contentWidthPx;
    tree.add(
      new Rect({
        x,
        y: lineY - (top ? 0 : edge.px),
        width: Math.max(0, right - x),
        height: edge.px,
        fill: "#1b1b1b",
      }),
    );
  }
  for (const line of para.lines) {
    const lineY = y + line.yPx;
    // In-line vertical placement: a docGrid body line centers its natural box
    // in the grid span (half-leading); every other regime (multiple without a
    // grid, atLeast, text boxes, headers/footers) anchors the text at the line
    // top and sinks the slack below. Both verified against the reference PDF.
    const pad = line.grid ? Math.max(0, (line.heightPx - line.naturalPx) / 2) : 0;
    // Line x origin: the left indent (every line) plus THIS line's own
    // first-line indent — a split tail's leading line carries none (it is
    // mid-paragraph), so the flag lives on the line, not the line index.
    const lineX = x + (para.indent?.leftPx ?? 0) + (line.firstLineIndentPx ?? 0);
    // A justified line stretches each text item to the next item's x (the
    // last one to the line's content width): Leafer's textAlign "both-letter"
    // spreads the slack as uniform letter spacing inside that interval —
    // Word's CJK justification model.
    const justified = line.justifyGapPx != null;
    const rights: number[] = [];
    if (justified) {
      // The last item's stretch target extends past the content width by the
      // overflow-punct hang: the full glyphs fill the width (both-letter
      // spacing covers the wider interval), and the closer lands in the
      // margin at ~its natural advance — Word's justified hang.
      let nextLeft = (line.maxWidthPx ?? 0) + (line.hangPx ?? 0);
      for (let i = line.items.length - 1; i >= 0; i--) {
        rights[i] = nextLeft;
        nextLeft = line.items[i].xPx;
      }
    }
    for (const [itemIndex, item] of line.items.entries()) {
      const inline: LayoutInline | undefined = para.inline[item.inlineIndex];
      if (!inline) continue;
      if (item.kind === "text" && inline.kind === "text") {
        const family = familyOf(inline.style, item.text);
        const intervalPx = justified ? rights[itemIndex] - item.xPx : undefined;
        // A page-number field paints its live value; the measured `text` was
        // only a placeholder.
        const label =
          inline.field === "page"
            ? String(ctx.pageIndex + 1)
            : inline.field === "numPages"
              ? String(ctx.pageCount)
              : item.text;
        tree.add(
          new Text({
            x: lineX + item.xPx,
            y: lineY + pad,
            // width ONLY on justified items (their stretch interval): a width
            // on every line would let Leafer wrap the slice again with its
            // own metrics (a phantom second line). textWrap "none" keeps the
            // interval from wrapping; height keeps the element paintable
            // (height 0 is skipped by Leafer).
            width: intervalPx,
            textWrap: justified ? "none" : undefined,
            // CJK items spread per glyph (both-letter); Latin items spread
            // per word gap (both-justify — Leafer's word mode, Word's Latin
            // justification). "both" keeps the single-row Text justifiable.
            textAlign: justified
              ? /[一-鿿぀-ヿ가-힯]/.test(item.text)
                ? "both-letter"
                : "both-justify"
              : undefined,
            height: Math.max(1, line.heightPx),
            text: label,
            fill: inline.style.color ? `#${inline.style.color}` : "#1b1b1b",
            textDecoration:
              inline.style.underline && inline.style.strikethrough
                ? "under-delete"
                : inline.style.underline
                  ? "under"
                  : inline.style.strikethrough
                    ? "delete"
                    : undefined,
            fontFamily: family,
            fontSize: inline.style.sizePx,
            // Numbers only: Leafer's fontWeight setter treats strings as named
            // weights ("bold"/"thin"…) and silently maps unknown strings to 400,
            // so a string "700" would lose bold. Italic is the `italic` boolean
            // property — there is no fontStyle.
            fontWeight: inline.style.bold ? 700 : 400,
            italic: inline.style.italic,
            letterSpacing: inline.style.letterSpacingPx
              ? { type: "px", value: inline.style.letterSpacingPx }
              : undefined,
          }),
        );
      } else if (item.kind === "picture" && inline.kind === "picture") {
        if (inline.src) {
          tree.add(
            new LeaferImage({
              url: inline.src,
              x: lineX + item.xPx,
              y: lineY + pad,
              width: item.widthPx,
              height: item.heightPx,
            }),
          );
        } else {
          // Linked-only picture (no bytes in the package): an empty frame.
          tree.add(
            new Rect({
              x: lineX + item.xPx,
              y: lineY,
              width: item.widthPx,
              height: item.heightPx,
              fill: "#f3f3f3",
              stroke: "#c4c4c4",
              strokeWidth: 1,
            }),
          );
        }
      } else if (item.kind === "tab" && inline.kind === "tab") {
        if (item.leader && item.widthPx > 1) {
          // Leader fill across the tab's advance (a TOC's dot row). The glyph
          // metrics come from the line's text — the tab atom carries no style.
          let sizePx = 0;
          let color = "#1b1b1b";
          for (const other of line.items) {
            const src = para.inline[other.inlineIndex];
            if (src?.kind === "text" && src.style.sizePx > sizePx) {
              sizePx = src.style.sizePx;
              color = src.style.color ? `#${src.style.color}` : color;
            }
          }
          if (sizePx > 0) paintTabLeader(tree, item, lineX, lineY, pad, sizePx, color);
        }
      }
    }
  }
  // Floating drawings anchored to this paragraph: wrap-none boxes painted
  // over the text — the flow reserved them no height. (behindDoc ones went
  // first, above.)
  for (const drawing of para.drawings ?? []) {
    if (!drawing.behind) paintDrawing(tree, drawing, x, y, ctx);
  }
}

/** w:leader fill across a tab's interval: dots/hyphens/underscores drawn on
 *  the text baseline (a hair below it for the underscore, Word's placement). */
function paintTabLeader(
  tree: IGroup,
  item: Extract<LaidOutLineItem, { kind: "tab" }>,
  lineX: number,
  lineY: number,
  pad: number,
  sizePx: number,
  color: string,
): void {
  const style = item.leader ? TAB_LEADER_STYLES[item.leader] : undefined;
  if (!style) return;
  const x1 = lineX + item.xPx;
  const x2 = x1 + item.widthPx;
  if (x2 - x1 < 2) return;
  const y = lineY + pad + sizePx * (style.underside ? 0.9 : 0.82);
  tree.add(
    new Line({
      points: [x1, y, x2, y],
      stroke: color,
      strokeWidth: style.widthPx,
      dashPattern: style.dash,
    }),
  );
}

/** Per-leader dash patterns: [on, off] in px. A sub-pixel `on` value would
 *  round to zero in Leafer's dash pass and paint nothing, so every leader
 *  uses a positive on-width (Word's dots render as tiny squares anyway). */
const TAB_LEADER_STYLES: Record<
  NonNullable<Extract<LaidOutLineItem, { kind: "tab" }>["leader"]>,
  { dash?: number[]; widthPx: number; underside?: boolean }
> = {
  dot: { dash: [1.2, 2.7], widthPx: 1.2 },
  heavy: { dash: [2.2, 1.7], widthPx: 2.2 },
  middleDot: { dash: [1.6, 2.3], widthPx: 1.6 },
  hyphen: { dash: [3, 2.5], widthPx: 1 },
  underscore: { widthPx: 1, underside: true },
};

/** One floating drawing: members absolutely positioned in the drawing's box,
 *  itself placed by the anchor spec against the page geometry. A text box
 *  stacks its own paragraphs inside its insets (the same stackBlocks the
 *  header/footer furniture uses). */
function paintDrawing(
  tree: IGroup,
  drawing: LayoutDrawing,
  x: number,
  y: number,
  ctx: PaintContext,
): void {
  const { flow } = ctx;
  // The reference box each axis resolves against: the content box (column /
  // topMargin), the page box, an edge (leftMargin/rightMargin/bottomMargin),
  // or — vertically — the anchor paragraph's own top (extent 0: offsets and
  // align both hang off the edge itself).
  const hBox =
    drawing.anchor.horizontal.relative === "page"
      ? { left: 0, width: flow.pageWidthPx }
      : drawing.anchor.horizontal.relative === "rightMargin"
        ? { left: flow.contentLeftPx + flow.contentWidthPx, width: 0 }
        : drawing.anchor.horizontal.relative === "leftMargin"
          ? { left: flow.contentLeftPx, width: 0 }
          : { left: flow.contentLeftPx, width: flow.contentWidthPx };
  const vBox =
    drawing.anchor.vertical.relative === "page"
      ? { top: 0, height: flow.pageHeightPx }
      : drawing.anchor.vertical.relative === "paragraph"
        ? { top: y, height: 0 }
        : drawing.anchor.vertical.relative === "bottomMargin"
          ? { top: flow.contentTopPx + flow.contentHeightPx, height: 0 }
          : { top: flow.contentTopPx, height: flow.contentHeightPx };
  // Axis position: align inside the reference box, else the offset (px, or a
  // fraction of the reference extent) from its leading edge.
  const axisPos = (
    spec: { offsetPx?: number; percent?: number; align?: string },
    base: number,
    extent: number,
    size: number,
  ): number => {
    if (spec.align === "center") return base + (extent - size) / 2;
    if (spec.align === "right" || spec.align === "bottom") return base + extent - size;
    if (spec.align) return base; // left / top
    if (spec.percent != null) return base + spec.percent * extent;
    return base + (spec.offsetPx ?? 0);
  };
  const boxX = axisPos(drawing.anchor.horizontal, hBox.left, hBox.width, drawing.width);
  const boxY = axisPos(drawing.anchor.vertical, vBox.top, vBox.height, drawing.height);
  for (const m of drawing.members) {
    const mx = boxX + m.x;
    const my = boxY + m.y;
    if (m.kind === "picture") {
      if (m.src && m.crop) {
        addCroppedImage(tree, m, mx, my);
      } else if (m.src) {
        tree.add(
          new LeaferImage({
            url: m.src,
            x: mx,
            y: my,
            width: m.width,
            height: m.height,
            // Mirrors flip around the element's (x,y) origin: shifting the
            // origin to the far edge first makes the reflection cover the
            // original box exactly.
            ...(m.flipH ? { x: mx + m.width, scaleX: -1 } : {}),
            ...(m.flipV ? { y: my + m.height, scaleY: -1 } : {}),
          }),
        );
      } else {
        tree.add(
          new Rect({
            x: mx,
            y: my,
            width: m.width,
            height: m.height,
            fill: "#f3f3f3",
            stroke: "#c4c4c4",
            strokeWidth: 1,
          }),
        );
      }
    } else if (m.kind === "path") {
      tree.add(
        new LeaferPath({
          x: mx,
          y: my,
          width: m.width,
          height: m.height,
          // Leafer's Path takes SVG path data under `path` (its `data` holds
          // the parsed command array — a string there paints nothing).
          path: m.d,
          fill: m.fill ? `#${m.fill}` : undefined,
          stroke: m.line ? (m.line.color ? `#${m.line.color}` : "#000000") : undefined,
          strokeWidth: m.line?.px,
          strokeCap:
            m.line?.cap === "round" ? "round" : m.line?.cap === "square" ? "square" : undefined,
          strokeJoin:
            m.line?.join === "round" ? "round" : m.line?.join === "bevel" ? "bevel" : undefined,
          dashPattern: m.line?.dash ? PRSTDASH_PATTERN[m.line.dash] : undefined,
        }),
      );
    } else if (m.kind === "shape") {
      const fill = m.fill ? `#${m.fill}` : undefined;
      const stroke = m.line ? (m.line.color ? `#${m.line.color}` : "#000000") : undefined;
      const strokeWidth = m.line?.px;
      if (m.preset === "ellipse") {
        tree.add(
          new Ellipse({
            x: mx,
            y: my,
            width: m.width,
            height: m.height,
            fill,
            stroke,
            strokeWidth,
          }),
        );
      } else {
        // rect/roundRect and every box-like preset; unknown presets skip —
        // an honest absence beats a wrong geometry.
        if (m.preset !== "rect" && m.preset !== "roundRect") continue;
        tree.add(
          new Rect({
            x: mx,
            y: my,
            width: m.width,
            height: m.height,
            fill,
            stroke,
            strokeWidth,
            cornerRadius: m.preset === "roundRect" ? Math.min(m.width, m.height) / 6 : undefined,
          }),
        );
      }
    } else {
      const measurer = new TextMeasurer(ctx.metrics);
      const left = m.insets?.left ?? 0;
      const inner = Math.max(0, m.width - left - (m.insets?.right ?? 0));
      const laid = stackBlocks(m.blocks, inner, undefined, measurer);
      let oy = my + (m.insets?.top ?? 0);
      if (m.anchor !== "top") {
        const slack = m.height - (m.insets?.top ?? 0) - (m.insets?.bottom ?? 0) - laid.heightPx;
        oy += m.anchor === "center" ? Math.max(0, slack / 2) : Math.max(0, slack);
      }
      for (const item of laid.stack) {
        paintBlock(tree, item.block, mx + left, oy + item.yPx, ctx);
      }
    }
  }
}

/** A cropped picture (a:srcRect): Leafer paints whole sources only, so the
 *  sub-region renders through an offscreen canvas copy, added when decoded
 *  (the stage re-paints on the next sync regardless). */
function addCroppedImage(
  tree: IGroup,
  m: Extract<LayoutDrawingMember, { kind: "picture" }>,
  mx: number,
  my: number,
): void {
  const el = new Image();
  el.onload = () => {
    const sx = Math.round(m.crop!.left * el.naturalWidth);
    const sy = Math.round(m.crop!.top * el.naturalHeight);
    const sw = Math.max(1, el.naturalWidth - sx - Math.round(m.crop!.right * el.naturalWidth));
    const sh = Math.max(1, el.naturalHeight - sy - Math.round(m.crop!.bottom * el.naturalHeight));
    const canvas = document.createElement("canvas");
    canvas.width = sw;
    canvas.height = sh;
    canvas.getContext("2d")?.drawImage(el, sx, sy, sw, sh, 0, 0, sw, sh);
    tree.add(
      new LeaferImage({
        url: canvas.toDataURL("image/png"),
        x: mx,
        y: my,
        width: m.width,
        height: m.height,
        ...(m.flipH ? { x: mx + m.width, scaleX: -1 } : {}),
        ...(m.flipV ? { y: my + m.height, scaleY: -1 } : {}),
      }),
    );
  };
  el.src = m.src!;
}

function paintTable(
  tree: IGroup,
  table: LaidOutTable,
  x: number,
  y: number,
  ctx: PaintContext,
): void {
  // w:jc: the whole grid (borders included) shifts as one box.
  x += table.offsetXPx ?? 0;
  const cols = table.columnWidthsPx;
  const nRows = table.rows.length;
  const nCols = cols.length;
  // Boundary coordinates (column left edges + the right rim, row tops + bottom).
  const colX = [0];
  for (const w of cols) colX.push(colX[colX.length - 1] + w);
  const rowY = [0];
  for (const row of table.rows) rowY.push(rowY[rowY.length - 1] + row.heightPx);

  // Occupancy: occ[r][c] = the cell covering that grid slot. A merged cell
  // fills its whole span so boundary resolution sees across rows/columns; a
  // boundary inside a span has the same cell on both sides and draws nothing.
  const occ: (LaidOutCell | undefined)[][] = Array.from({ length: nRows }, () =>
    Array.from<LaidOutCell | undefined>({ length: nCols }),
  );
  table.rows.forEach((row, r) => {
    let col = 0;
    for (const cell of row.cells) {
      while (col < nCols && occ[r][col]) col++;
      if (col >= nCols) break;
      const spanW = Math.min(cell.colspan, nCols - col);
      const spanH = Math.min(cell.rowspan ?? 1, nRows - r);
      // Shading covers the merged box; content anchors to the start row (the
      // engine measured it there).
      if (cell.fill) {
        tree.add(
          new Rect({
            x: x + colX[col],
            y: y + rowY[r],
            width: colX[col + spanW] - colX[col],
            height: rowY[r + spanH] - rowY[r],
            fill: `#${cell.fill}`,
          }),
        );
      }
      for (let dr = 0; dr < spanH; dr++) {
        for (let dc = 0; dc < spanW; dc++) occ[r + dr][col + dc] = cell;
      }
      const contentX = x + colX[col] + (cell.insets.left ?? 0);
      const contentY = y + rowY[r] + (cell.insets.top ?? 0) + (cell.contentOffsetYPx ?? 0);
      for (const stacked of cell.stack) {
        paintBlock(tree, stacked.block, contentX, contentY + stacked.yPx, ctx);
      }
      col += spanW;
    }
  });

  // Collapsed borders: every shared boundary resolves to its heaviest
  // candidate — the two adjacent cells' own edges and the table-level default
  // (rim edges for the grid's outline, inside edges between cells). Word's
  // conflict rule is width-first; ties keep the earlier candidate.
  const tb = table.borders;
  /** A horizontal boundary (row edge `b`) at column `c`: the cell above ends
   *  here, the cell below starts here. */
  const pickH = (b: number, c: number): LayoutBorderEdge | undefined => {
    const above = b > 0 ? occ[b - 1][c] : undefined;
    const below = b < nRows ? occ[b][c] : undefined;
    if (above && above === below) return undefined;
    // An explicitly declared cell edge (w:tcBorders, nil included) suppresses
    // the table-level default — Word resolves tcBorders over tblBorders
    // outright; only cells silent on the edge fall back to the grid default.
    const aboveEdge = above?.borders?.bottom;
    const belowEdge = below?.borders?.top;
    const def =
      aboveEdge || belowEdge
        ? undefined
        : b === 0
          ? tb?.top
          : b === nRows
            ? tb?.bottom
            : tb?.insideHorizontal;
    return heaviest(aboveEdge, heaviest(belowEdge, def));
  };
  /** A vertical boundary (column edge `b`) at row `r`. */
  const pickV = (b: number, r: number): LayoutBorderEdge | undefined => {
    const left = b > 0 ? occ[r][b - 1] : undefined;
    const right = b < nCols ? occ[r][b] : undefined;
    if (left && left === right) return undefined;
    const leftEdge = left?.borders?.right;
    const rightEdge = right?.borders?.left;
    const def =
      leftEdge || rightEdge
        ? undefined
        : b === 0
          ? tb?.left
          : b === nCols
            ? tb?.right
            : tb?.insideVertical;
    return heaviest(leftEdge, heaviest(rightEdge, def));
  };
  // Contiguous boundary slots with an identical winner merge into one stroke.
  for (let b = 0; b <= nRows; b++) {
    let segStart = -1;
    let seg: LayoutBorderEdge | undefined;
    for (let c = 0; c <= nCols; c++) {
      const winner = c < nCols ? pickH(b, c) : undefined;
      if (seg && winner && sameEdge(seg, winner)) continue;
      if (seg) {
        drawEdge(tree, x + colX[segStart], y + rowY[b], colX[c] - colX[segStart], true, seg);
      }
      seg = winner && edgeWeight(winner) > 0 ? winner : undefined;
      segStart = seg ? c : -1;
    }
  }
  for (let b = 0; b <= nCols; b++) {
    let segStart = -1;
    let seg: LayoutBorderEdge | undefined;
    for (let r = 0; r <= nRows; r++) {
      const winner = r < nRows ? pickV(b, r) : undefined;
      if (seg && winner && sameEdge(seg, winner)) continue;
      if (seg) {
        drawEdge(tree, x + colX[b], y + rowY[segStart], rowY[r] - rowY[segStart], false, seg);
      }
      seg = winner && edgeWeight(winner) > 0 ? winner : undefined;
      segStart = seg ? r : -1;
    }
  }
}

/** One border edge's conflict weight: nil/none/absent carry none. */
function edgeWeight(edge: LayoutBorderEdge | undefined): number {
  return edge && edge.style && edge.style !== "nil" && edge.style !== "none" && edge.px != null
    ? edge.px
    : 0;
}

/** Word's border conflict resolution: the wider edge wins; ties keep `a`. */
function heaviest(
  a: LayoutBorderEdge | undefined,
  b: LayoutBorderEdge | undefined,
): LayoutBorderEdge | undefined {
  return edgeWeight(b) > edgeWeight(a) ? b : a;
}

function sameEdge(a: LayoutBorderEdge, b: LayoutBorderEdge): boolean {
  return a === b || (a.px === b.px && a.style === b.style && a.color === b.color);
}

/** dashPattern per OOXML border style (stroke-only in Leafer); styles without
 *  an entry render solid — the visual fallback for wave/3D composites. */
const DASH_PATTERN: Record<string, number[]> = {
  dashed: [4, 2],
  dashSmallGap: [2, 2],
  dotted: [1, 2],
  dashDot: [4, 2, 1, 2],
  dashDotDot: [4, 2, 1, 2, 1, 2],
};

/** Draw one collapsed edge centered on its boundary: a stroked Line (dash
 *  styles apply to strokes, not fills); double/triple split the width into
 *  parallel strokes. */
function drawEdge(
  tree: IGroup,
  ex: number,
  ey: number,
  len: number,
  horizontal: boolean,
  edge: LayoutBorderEdge,
): void {
  // Word's screen rendering lifts hairlines to a full pixel at 100% zoom —
  // a sub-pixel stroke here would render as a faint half-transparent line.
  const px = Math.max(edge.px ?? 1, 1);
  const color = edge.color ? `#${edge.color}` : "#000000";
  const style = edge.style ?? "single";
  const dash = DASH_PATTERN[style];
  const stroke = (offset: number, thickness: number): void => {
    const wx = horizontal ? ex : ex + offset;
    const wy = horizontal ? ey + offset : ey;
    tree.add(
      new Line({
        x: wx,
        y: wy,
        // Line points are relative to x/y; a zero-length second point pins the
        // direction (horizontal → +x, vertical → +y).
        points: horizontal ? [0, 0, len, 0] : [0, 0, 0, len],
        stroke: color,
        strokeWidth: thickness,
        dashPattern: dash,
      }),
    );
  };
  if (style === "double" || style === "triple") {
    const unit = px / (style === "double" ? 3 : 4);
    stroke(-px / 2 + unit / 2, unit);
    if (style === "triple") stroke(0, unit);
    stroke(px / 2 - unit / 2, unit);
    return;
  }
  stroke(0, px);
}

function paintPlaceholder(
  tree: IGroup,
  block: { heightPx: number; label?: string },
  x: number,
  y: number,
): void {
  const width = 240;
  const height = Math.max(20, block.heightPx);
  tree.add(
    new Rect({
      x,
      y,
      width,
      height,
      fill: "#fafafa",
      stroke: "#d0d0d0",
      strokeWidth: 1,
      dashPattern: [4, 3],
    }),
  );
  if (block.label) {
    tree.add(
      new Text({
        x: x + 8,
        y: y + 4,
        text: `${block.label} (not rendered yet)`,
        fill: "#9a9a9a",
        fontFamily: "Inter, sans-serif",
        fontSize: 11,
      }),
    );
  }
}

/** The font family a text slice paints in: a slot-aware pick over the item's
 *  own text (the engine already itemized runs; CJK slices take the eastAsia
 *  slot — the resolveFontFamily rule, applied at paint time). */
function familyOf(style: LayoutTextStyle, text: string): string {
  if (typeof style.family === "string") return style.family;
  const slots = style.family;
  if (text && isCjkCodeUnit(text, 0)) {
    return slots.eastAsia ?? slots.latin ?? "Inter";
  }
  return slots.latin ?? "Inter";
}
