/**
 * Scene painter — walks a laid-out page (the @docen/layout result) and builds
 * the LeaferJS tree. The layout engine owns ALL geometry; painting positions
 * what it is given and never measures. Text elements carry explicit width AND
 * height — Leafer never paints an element whose height is still 0.
 *
 * @module
 */
import {
  familyOfSlot,
  gridPadOf,
  isCjkCodeUnit,
  justifiedIntervals,
  justifyPerGrapheme,
  lineOriginXPx,
  stackBlocks,
  tableGridOf,
  TextMeasurer,
  vertAlignedSizePx,
  vertAlignBaselineShiftPx,
  type FlowItem,
  type FontMetrics,
  type LaidOutBlock,
  type LaidOutLineItem,
  type LaidOutParagraph,
  type LaidOutStackItem,
  type LaidOutTable,
  type LayoutBlockContext,
  type LayoutBorderEdge,
  type LayoutDrawing,
  type LayoutDrawingMember,
  type LayoutInline,
  type LayoutTextStyle,
  type ProjectedFlowBox,
  type ProjectedPageBackground,
  type ProjectedPageFurniture,
} from "@docen/layout";
import {
  Box,
  Ellipse,
  Group,
  Image as LeaferImage,
  ImageManager,
  Line,
  Path as LeaferPath,
  Rect,
  Text,
  type IGroup,
  type ILeaferImage,
} from "leafer-ui";

/** OOXML prstDash tokens → dash patterns in px (line-width units, the host's
 *  preset line styles); unlisted tokens render solid. */
const PRSTDASH_PATTERN: Record<string, number[]> = {
  dot: [1, 3],
  // A 1px-on/1px-off antialiased hairline blends to a faint tint — Word's
  // sysDot boxes read as clear 2px dots at hairline widths (user-verified).
  sysDot: [2, 2],
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

/** One hit-testable drawing box, page-local px — what a click needs to grab a
 *  drawing (Word: clicking a picture selects it). `para` is the laid host
 *  paragraph (the caret map resolves it to the PM position) and `index` the
 *  drawing's position among that paragraph's drawings, matching the run order
 *  projectDrawings collected them in. */
export interface DrawingHitBox {
  page: number;
  x: number;
  y: number;
  width: number;
  height: number;
  para: LaidOutParagraph;
  index: number;
  /** "drawing" — a floating picture/shape from para.drawings (index counts
   *  that sequence); "inline" — a picture line item (index counts the
   *  paragraph's inline pictures). The PM side re-finds the node per kind. */
  kind: "drawing" | "inline";
}

/** The paint context for one page — the stage context plus the page's own
 *  identity (page-number fields resolve against it) and which of Word's two
 *  text-underlapping layers is being composed right now (the stage paints a
 *  page twice: once for behind-doc floats, once for everything else with
 *  header/footer furniture between them — Word renders footer furniture and
 *  the body over those floats, so furniture must not sit under them).
 *
 *  The flow box and furniture are the PAGE's OWN section's (multi-section
 *  documents give every page the box of the section it belongs to). */
export interface PaintContext {
  metrics: FontMetrics;
  flow: ProjectedFlowBox;
  furniture?: ProjectedPageFurniture;
  background?: ProjectedPageBackground;
  pageIndex: number;
  pageCount: number;
  layer: "behind" | "body";
  /** Forces a frame after an async image insert: Leafer's change-driven
   *  scheduling stalls on apps created while offscreen (see stage.repaint),
   *  so a decode completing after repaint would otherwise never show. */
  rerender: () => void;
  /** Accumulates this page's drawing boxes as the body pass paints them —
   *  the stage turns the list into its click hit table. */
  hitBoxes?: DrawingHitBox[];
}

/** The text column a block paints inside: the page's content box for body
 *  blocks, the cell's inner box for table content (a text box's insets box
 *  for its paragraphs). Cell-anchored floats clamp inside it — Word's
 *  layoutInCell containment, matching the wrap zones the layout built. */
interface PaintColumn {
  width: number;
  inCell: boolean;
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
  col?: PaintColumn,
): void {
  switch (block.kind) {
    case "paragraph":
      paintParagraph(tree, block, x, y, ctx, col);
      return;
    // Only paragraphs can carry drawings; the behind pass therefore skips
    // every other block so nothing paints twice.
    case "table":
    case "placeholder":
    case "pageBreak":
      if (ctx.layer === "behind") return;
      if (block.kind === "table") paintTable(tree, block, x, y, ctx);
      else if (block.kind === "placeholder") paintPlaceholder(tree, block, x, y);
      return;
    case "group":
      for (const child of block.children) {
        paintBlock(tree, child.block, x, y + child.yPx, ctx, col);
      }
      return;
  }
}

function paintParagraph(
  tree: IGroup,
  para: LaidOutParagraph,
  x: number,
  y: number,
  ctx: PaintContext,
  col?: PaintColumn,
): void {
  // The stage composes a page in two passes; this one paints only its own
  // layer's drawings. Behind-doc floats land beneath the furniture pass.
  const behind = ctx.layer === "behind";
  if (behind) {
    let index = 0;
    for (const drawing of para.drawings ?? []) {
      if (drawing.behind) paintDrawing(tree, drawing, x, y, ctx, col, { para, index: index++ });
      else index++;
    }
    return;
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
  let inlinePicIndex = 0;
  for (const line of para.lines) {
    const lineY = y + line.yPx;
    // In-line vertical placement: a docGrid line centers its natural box in
    // the grid span (half-leading — body flow and text-box stacks alike);
    // every other regime (multiple without a grid, atLeast, plain text
    // boxes, header/footer stories) anchors the text at the line top and
    // sinks the slack below. All verified against the reference PDF.
    const pad = gridPadOf(line);
    // Line x origin — the shared sum (left indent + the line's own first-line
    // indent + a wrapSide float's shift) the caret map anchors by too.
    const lineX = x + lineOriginXPx(para, line);
    // A justified line stretches each text item to the next item's x (the
    // last one past the content width by the overflow-punct hang): Leafer's
    // textAlign "both-letter" spreads the slack as uniform letter spacing
    // inside that interval — Word's CJK justification model.
    const rights = justifiedIntervals(line);
    for (const [itemIndex, item] of line.items.entries()) {
      const inline: LayoutInline | undefined = para.inline[item.inlineIndex];
      if (!inline) continue;
      if (item.kind === "text" && inline.kind === "text") {
        const family = familyOf(inline.style, item.text);
        const intervalPx = rights ? rights[itemIndex]! - item.xPx : undefined;
        // Comment range tint (w:commentRangeStart..End): a translucent box
        // under the item's glyphs, Word's light-amber reviewer tint. Painted
        // before the text so the glyphs stay on top.
        if (inline.commentIds?.length) {
          tree.add(
            new Rect({
              x: lineX + item.xPx,
              y: lineY + pad,
              width: intervalPx ?? item.widthPx,
              height: Math.max(1, line.naturalPx || line.heightPx),
              fill: "rgba(255, 222, 89, 0.45)",
            }),
          );
        }
        // A page-number field paints its live value; the measured `text` was
        // only a placeholder.
        const label =
          inline.field === "page"
            ? String(ctx.pageIndex + 1)
            : inline.field === "numPages"
              ? String(ctx.pageCount)
              : item.text;
        const textEl = new Text({
          x: lineX + item.xPx,
          // A raised/lowered run (w:vertAlign — the footnote reference) paints
          // at the scaled size on a shifted baseline; the scaling itself is
          // the shared vertAlignedSizePx so measure and paint agree.
          y: lineY + pad + vertAlignBaselineShiftPx(inline.style),
          // width ONLY on justified items (their stretch interval): a width
          // on every line would let Leafer wrap the slice again with its
          // own metrics (a phantom second line). textWrap "none" keeps the
          // interval from wrapping; height keeps the element paintable
          // (height 0 is skipped by Leafer).
          width: intervalPx,
          textWrap: rights ? "none" : undefined,
          // CJK items spread per glyph (both-letter); Latin items spread
          // per word gap (both-justify — Leafer's word mode, Word's Latin
          // justification). "both" keeps the single-row Text justifiable.
          textAlign: rights
            ? justifyPerGrapheme(item.text)
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
          fontSize: vertAlignedSizePx(inline.style),
          // Leafer's default 150% line spacing half-leads the glyphs ~0.25×
          // fontSize below the line-box top the layout handed over (text-box
          // text riding low). The px form pins one line's spacing to the font
          // size — the percent form (`{ type: "percent" }`) silently blanks
          // every body Text when combined with an explicit height.
          lineHeight: vertAlignedSizePx(inline.style),
          // Numbers only: Leafer's fontWeight setter treats strings as named
          // weights ("bold"/"thin"…) and silently maps unknown strings to 400,
          // so a string "700" would lose bold. Italic is the `italic` boolean
          // property — there is no fontStyle.
          fontWeight: inline.style.bold ? 700 : 400,
          italic: inline.style.italic,
          letterSpacing: inline.style.letterSpacingPx
            ? { type: "px", value: inline.style.letterSpacingPx }
            : undefined,
        });
        tree.add(textEl);
      } else if (item.kind === "picture" && inline.kind === "picture") {
        // An inline picture is a grab target just like a floating drawing —
        // without a hit box a click lands behind the art (Word selects the
        // picture). Index counts the paragraph's inline pictures; the PM side
        // re-finds the same k-th non-floating image node.
        ctx.hitBoxes?.push({
          page: ctx.pageIndex,
          x: lineX + item.xPx,
          y: lineY + pad,
          width: item.widthPx,
          height: item.heightPx,
          para,
          index: inlinePicIndex++,
          kind: "inline",
        });
        if (inline.members) {
          // A metafile source replayed into members (WMF vector layers): the
          // structured scene paints in place of the flat image, clipped to
          // the extent — a srcRect leaves records reaching past the box and
          // GDI never lets metafile ink out of the playback rect. Leafer's
          // Group ignores `overflow` (a Box-only data getter clips children),
          // so the clip holder must be a Box.
          const holder = new Box({
            x: lineX + item.xPx,
            y: lineY + pad,
            width: item.widthPx,
            height: item.heightPx,
            overflow: "hide",
          });
          paintMembers(holder, inline.members, 0, 0, ctx);
          tree.add(holder);
        } else if (inline.src && inline.crop) {
          // A cropped flat source (a:srcRect): the visible remainder fills
          // the extent box — the whole source would stretch into it.
          addCroppedImage(
            tree,
            inline.src,
            inline.crop,
            lineX + item.xPx,
            lineY + pad,
            item.widthPx,
            item.heightPx,
            ctx,
          );
        } else if (inline.src) {
          pinImage(inline.src);
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
            if (src?.kind === "text") {
              // Raised/lowered runs count at their scaled size — a footnote
              // reference must not pull the leader dots up.
              const px = vertAlignedSizePx(src.style);
              if (px > sizePx) {
                sizePx = px;
                color = src.style.color ? `#${src.style.color}` : color;
              }
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
  // The body pass collects EVERY drawing's hit box — behind-doc floats
  // painted by the earlier pass included (their boxes are just as clickable).
  let hitIndex = 0;
  for (const drawing of para.drawings ?? []) {
    const host = { para, index: hitIndex++ };
    if (!drawing.behind) paintDrawing(tree, drawing, x, y, ctx, col, host);
    else if (ctx.hitBoxes) recordDrawingHit(drawing, x, y, ctx, ctx.hitBoxes, host);
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

/** The page-local box a drawing's anchor spec resolves to — the single
 *  implementation behind both the painter and the hit recorder, so what a
 *  click tests is exactly what painted. */
function drawingBoxOf(
  drawing: LayoutDrawing,
  x: number,
  y: number,
  ctx: PaintContext,
  col?: PaintColumn,
): { x: number; y: number } {
  const { flow } = ctx;
  // The reference box each axis resolves against: the content box (column /
  // topMargin), the page box, an edge (leftMargin/rightMargin/bottomMargin),
  // or — vertically — the anchor paragraph's own top (extent 0: offsets and
  // align both hang off the edge itself). The column's left edge is the
  // CALLER's text column — the page's for body paragraphs, the cell's for
  // cell-anchored floats (Word's layoutInCell), matching the wrap zones the
  // layout computes against the same base.
  const hBox =
    drawing.anchor.horizontal.relative === "page"
      ? { left: 0, width: flow.pageWidthPx }
      : drawing.anchor.horizontal.relative === "rightMargin"
        ? { left: flow.contentLeftPx + flow.contentWidthPx, width: 0 }
        : drawing.anchor.horizontal.relative === "leftMargin"
          ? { left: flow.contentLeftPx, width: 0 }
          : { left: x, width: col?.width ?? flow.contentWidthPx };
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
  let boxX = axisPos(drawing.anchor.horizontal, hBox.left, hBox.width, drawing.width);
  const boxY = axisPos(drawing.anchor.vertical, vBox.top, vBox.height, drawing.height);
  // Word's layoutInCell: a cell-anchored object never extends past its cell —
  // an offset that overflows the right edge shifts the whole box left to touch
  // it (the wrap zones shifted with it at layout time). Body floats keep
  // their raw position (they may hang into the margins).
  if (col?.inCell && drawing.anchor.horizontal.relative === "column") {
    boxX = Math.min(Math.max(boxX, hBox.left), hBox.left + Math.max(0, hBox.width - drawing.width));
  }
  return { x: boxX, y: boxY };
}

/** Record a drawing's hit box without painting it — the body pass catalogs
 *  behind-doc floats the earlier pass already painted. */
function recordDrawingHit(
  drawing: LayoutDrawing,
  x: number,
  y: number,
  ctx: PaintContext,
  boxes: DrawingHitBox[],
  host: DrawingHost,
): void {
  const box = drawingBoxOf(drawing, x, y, ctx);
  boxes.push({
    page: ctx.pageIndex,
    x: box.x,
    y: box.y,
    width: drawing.width,
    height: drawing.height,
    para: host.para,
    index: host.index,
    kind: "drawing",
  });
}

/** One floating drawing: members absolutely positioned in the drawing's box,
 *  itself placed by the anchor spec against the page geometry. A text box
 *  stacks its own paragraphs inside its insets (the same stackBlocks the
 *  header/footer furniture uses). */
/** Which paragraph carries a drawing, and its position among that paragraph's
 *  drawings (run order — how the PM side re-finds the node). */
interface DrawingHost {
  para: LaidOutParagraph;
  index: number;
}

function paintDrawing(
  tree: IGroup,
  drawing: LayoutDrawing,
  x: number,
  y: number,
  ctx: PaintContext,
  col: PaintColumn | undefined,
  host: DrawingHost,
): void {
  const box = drawingBoxOf(drawing, x, y, ctx, col);
  const boxX = box.x;
  const boxY = box.y;
  ctx.hitBoxes?.push({
    page: ctx.pageIndex,
    x: boxX,
    y: boxY,
    width: drawing.width,
    height: drawing.height,
    para: host.para,
    index: host.index,
    kind: "drawing",
  });
  if (drawing.clipMembers) {
    // A srcRect-cropped metafile replay reaches past the extent (GDI clips
    // metafile playback to the rect); wps text boxes must NOT clip — their
    // text may legitimately overflow a stale declared extent. Leafer clips
    // children only on a Box (`overflow` is Box data, Group ignores it).
    const holder = new Box({
      x: boxX,
      y: boxY,
      width: drawing.width,
      height: drawing.height,
      overflow: "hide",
    });
    paintMembers(holder, drawing.members, 0, 0, ctx);
    tree.add(holder);
  } else {
    paintMembers(tree, drawing.members, boxX, boxY, ctx);
  }
}

/** The members of a drawing box (or of an inline picture's metafile replay),
 *  each positioned at its own offset inside the box origin. Shared by the
 *  anchored-drawing and the inline-picture paths — the member shapes are the
 *  same; only the box origin differs. */
function paintMembers(
  tree: IGroup,
  members: readonly LayoutDrawingMember[],
  boxX: number,
  boxY: number,
  ctx: PaintContext,
): void {
  for (let i = 0; i < members.length; i++) {
    const m = members[i];
    const mx = boxX + m.x;
    const my = boxY + m.y;
    if (m.kind === "picture" && m.src && !m.crop) {
      // A masked GDI blt sequence (SRCPAINT then SRCAND halves) composites
      // against the metafile's own backdrop. The run is flattened into one
      // image before painting: canvas blend modes only see the editor's
      // layered App canvases, whose destinations are transparent — not the
      // underlying members a ternary raster-op needs.
      const run: Extract<LayoutDrawingMember, { kind: "picture" }>[] = [m];
      let end = i + 1;
      while (end < members.length) {
        const cur = members[end];
        if (cur.kind !== "picture" || !cur.src || cur.crop) break;
        run.push(cur);
        end++;
      }
      if (run.some((p) => p.blend)) {
        // A masked layer is meaningful only inside a composited run: painted
        // alone its opaque mask background (SRCPAINT halves are black-backed)
        // would lay a black slab over the page. A run of one blends against
        // nothing — honest absence beats a wrong slab.
        if (run.length > 1) addBlendedPictureRun(tree, run, boxX, boxY, ctx);
        i = end - 1;
        continue;
      }
    }
    if (m.kind === "picture") {
      if (m.src && m.crop) {
        addCroppedImage(tree, m.src, m.crop, mx, my, m.width, m.height, ctx, m.flipH, m.flipV);
      } else if (m.src) {
        addPlainImage(tree, m, mx, my, ctx);
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
          // A dashed hairline disappears entirely: the antialiased 1 px stroke
          // fades under the dash gaps and nothing survives. Hold a 1.5 px
          // floor on dashed strokes so the pattern renders; solid lines keep
          // their true width (fading is invisible there).
          strokeWidth:
            m.line?.px != null ? (m.line.dash ? Math.max(m.line.px, 1.5) : m.line.px) : undefined,
          strokeCap:
            m.line?.cap === "round" ? "round" : m.line?.cap === "square" ? "square" : undefined,
          strokeJoin:
            m.line?.join === "round" ? "round" : m.line?.join === "bevel" ? "bevel" : undefined,
          dashPattern: m.line?.dash ? PRSTDASH_PATTERN[m.line.dash] : undefined,
          // Leafer spells the SVG fill-rule attribute `windingRule`.
          windingRule: m.fillRule,
        }),
      );
    } else if (m.kind === "shape") {
      paintShapeBox(tree, { ...m, x: mx, y: my }, false);
    } else {
      // The txbx shape's own paint (prstGeom silhouette + spPr fill + a:ln)
      // sits under its text — Word draws the box even when the body is empty
      // (a plain text box is white fill + an accent hairline, visible on a
      // white page).
      paintShapeBox(tree, { ...m, x: mx, y: my }, true);
      const measurer = new TextMeasurer(ctx.metrics);
      const left = m.insets?.left ?? 0;
      // Metafile runs carry nowrap: GDI draws the string as-is, so the width
      // re-fit must not re-break it into phantom lines.
      const inner = m.nowrap
        ? Number.POSITIVE_INFINITY
        : Math.max(0, m.width - left - (m.insets?.right ?? 0));
      // A text box shares its STORY's doc grid with the surrounding text: a
      // body box snaps to the section grid and centers its grid rows
      // (onGrid — the half-leading the reference renders), while a
      // header/footer box gets no pitch at all (the furniture paint context
      // clears it — the story keeps natural line heights). bodyPr
      // @compatLnSpc plays no role: Word ignores it for wps txbxContent.
      // Metafile text carries nowrap: GDI strings are absolutely positioned
      // by the replay (baseline-derived y), not story rows — the grid pad
      // would shove them half a pitch off their drawn spot.
      const grid: LayoutBlockContext | undefined =
        ctx.flow.linePitchPx && !m.nowrap
          ? { linePitchPx: ctx.flow.linePitchPx, onGrid: true }
          : undefined;
      const laid = stackBlocks(m.blocks, inner, grid, measurer);
      let oy = m.insets?.top ?? 0;
      if (m.anchor === "center" || m.anchor === "bottom") {
        // spAutoFit shrinks the drawn box to the text, so slack resolves
        // against the fitted height — an oversized declared extent (stale
        // cy from a template) must not push the text down/center the box.
        const boxH = m.autoFit
          ? (m.insets?.top ?? 0) + laid.heightPx + (m.insets?.bottom ?? 0)
          : m.height;
        const slack = boxH - (m.insets?.top ?? 0) - (m.insets?.bottom ?? 0) - laid.heightPx;
        oy += m.anchor === "center" ? Math.max(0, slack / 2) : Math.max(0, slack);
      }
      if (m.rotation) {
        // Vertical metafile text (a rotated GDI world transform): the body
        // shapes horizontally as usual, then rotates about the box origin —
        // a group keeps the offsets in text space and carries the angle. A
        // clockwise 90° (vertical punctuation) swings the box left of the
        // origin, so the group shifts right by the box width and the ink
        // anchors the cell's top-right corner.
        const group = new Group({
          x: m.rotation === 90 ? mx + m.width : mx,
          y: my,
          rotation: m.rotation,
        });
        for (const item of laid.stack) {
          paintBlock(group, item.block, left, oy + item.yPx, ctx, { width: inner, inCell: true });
        }
        tree.add(group);
      } else {
        for (const item of laid.stack) {
          paintBlock(tree, item.block, mx + left, my + oy + item.yPx, ctx, {
            width: inner,
            inCell: true,
          });
        }
      }
    }
  }
}

/** Hex RRGGBB + opacity 0-1 → a CSS rgba color (Leafer parses CSS strings;
 *  the alpha must ride on the fill, not element opacity, so a stroke on the
 *  same shape stays opaque). */
function rgbaOf(hex: string, opacity: number): string {
  const n = parseInt(hex, 16);
  return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${opacity})`;
}

/** One box-like shape's own paint — preset silhouette + solid fill + outline
 *  stroke. Shared by standalone shape members and text-box shapes (a txbx is
 *  a shape carrying text; Word paints its prstGeom under the body). Unknown
 *  presets degrade to a plain rectangle when `rectFallback` (a txbx is a box
 *  by nature) and skip otherwise (an honest absence). */
function paintShapeBox(
  tree: IGroup,
  box: {
    x: number;
    y: number;
    width: number;
    height: number;
    preset?: string;
    fill?: string;
    opacity?: number;
    line?: { px: number; color?: string; dash?: string };
  },
  rectFallback: boolean,
): void {
  if (
    box.preset != null &&
    box.preset !== "rect" &&
    box.preset !== "roundRect" &&
    box.preset !== "ellipse"
  ) {
    if (!rectFallback) return;
  }
  const fill = box.fill
    ? box.opacity != null
      ? rgbaOf(box.fill, box.opacity)
      : `#${box.fill}`
    : undefined;
  const stroke = box.line ? (box.line.color ? `#${box.line.color}` : "#000000") : undefined;
  // A dashed hairline vanishes in the dash gaps — hold the same 1.5 px floor
  // the member path uses so preset dashes render.
  const strokeWidth =
    box.line?.px != null && box.line.dash ? Math.max(box.line.px, 1.5) : box.line?.px;
  const common = {
    x: box.x,
    y: box.y,
    width: box.width,
    height: box.height,
    fill,
    stroke,
    strokeWidth,
    // Closed shapes default to an inside stroke, under which Leafer's dash
    // pass paints nothing — center stroke renders the dashPattern.
    strokeAlign: "center",
    dashPattern: box.line?.dash ? PRSTDASH_PATTERN[box.line.dash] : undefined,
  };
  if (box.preset === "ellipse") {
    tree.add(new Ellipse(common));
    return;
  }
  tree.add(
    new Rect({
      ...common,
      cornerRadius: box.preset === "roundRect" ? Math.min(box.width, box.height) / 6 : undefined,
    }),
  );
}

/** Leafer evicts a decoded image larger than its 4MP cache threshold the
 *  moment a paint's use count drops back to zero (ImageManager.recycle →
 *  Resource.remove), so page recycling under scroll re-loaded every big
 *  banner/photo from its data URL and painted it a second late — the pop-in.
 *  One pinned use per media url keeps the decoded entry resident for the
 *  stage's lifetime: every later paint hits the ready entry and renders
 *  synchronously. The per-position `LeaferImage` elements stay thin shells —
 *  Leafer's paint resolves the bitmap through this shared entry. */
const pinnedImages = new Map<string, ILeaferImage>();

function pinImage(url: string): void {
  if (pinnedImages.has(url)) return;
  const image = ImageManager.get({ url }, "image");
  pinnedImages.set(url, image);
  image.load();
}

/** Release the pins at stage teardown — Leafer's own recycle then evicts the
 *  large entries it no longer shares. */
export function releasePinnedImages(): void {
  for (const image of pinnedImages.values()) ImageManager.recycle(image);
  pinnedImages.clear();
}

/** An uncropped picture. The element must join the tree only once its encoding
 *  is decoded (the stage renders eagerly right after paint — Leafer's
 *  change-driven re-render stalls afterwards, so an Image added before its
 *  url has a bitmap would be force-rendered empty and never picked back up).
 *  A transparent placeholder holds the paint-order slot meanwhile: the async
 *  insert must not land on top of members painted after this one, or z-order
 *  would follow decode completion instead of record order (labels would sink
 *  under their plates). */
function addPlainImage(
  tree: IGroup,
  m: Extract<LayoutDrawingMember, { kind: "picture" }>,
  mx: number,
  my: number,
  ctx: PaintContext,
): void {
  pinImage(m.src!);
  const slot = new Rect({ x: mx, y: my, width: m.width, height: m.height });
  tree.add(slot);
  const el = new Image();
  el.onload = () => {
    // A repaint since the decode started cleared the tree (slot included) —
    // that repaint's own decode now owns the paint-order slot.
    if (!slot.parent) return;
    tree.addAfter(
      new LeaferImage({
        url: m.src!,
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
      slot,
    );
    tree.remove(slot);
    // The stage's eager render already ran when this decode finished; without
    // a fresh frame the inserted image waits for a repaint that may never come.
    ctx.rerender();
  };
  el.src = m.src!;
}

/** A run of masked GDI blt members (SRCPAINT/SRCAND halves, optionally over a
 *  plain backdrop picture): decode all sources, flatten them in record order
 *  through canvas `screen`/`multiply` compositing — the ternary raster-op
 *  semantics — and insert the result as one image. Decode failures drop that
 *  member; if nothing survives the run falls back to individual painting. */
function addBlendedPictureRun(
  tree: IGroup,
  run: Extract<LayoutDrawingMember, { kind: "picture" }>[],
  boxX: number,
  boxY: number,
  ctx: PaintContext,
): void {
  const x0 = Math.min(...run.map((p) => p.x));
  const y0 = Math.min(...run.map((p) => p.y));
  const width = Math.ceil(Math.max(...run.map((p) => p.x + p.width)) - x0);
  const height = Math.ceil(Math.max(...run.map((p) => p.y + p.height)) - y0);
  if (width < 1 || height < 1 || width > 8192 || height > 8192) return;
  const slot = new Rect({ x: boxX + x0, y: boxY + y0, width, height });
  tree.add(slot);
  const loads = run.map(
    (p) =>
      new Promise<HTMLImageElement | null>((resolve) => {
        const el = new Image();
        el.onload = () => resolve(el);
        el.onerror = () => resolve(null);
        el.src = p.src!;
      }),
  );
  void Promise.all(loads).then((decoded) => {
    // A repaint since the decode started cleared the tree (slot included) —
    // that repaint's own run now owns the paint-order slot.
    if (!slot.parent) return;
    if (!decoded.some(Boolean)) {
      tree.remove(slot);
      // Masked halves never paint alone (their opaque mask background would
      // slab the page) — only plain backdrops fall back.
      for (const p of run) if (!p.blend) addPlainImage(tree, p, boxX + p.x, boxY + p.y, ctx);
      return;
    }
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const c2d = canvas.getContext("2d")!;
    // GDI replays these blts against its live destination surface — the page
    // behind the metafile — so the composite stays transparent wherever the
    // records keep the destination: a screen half only marks the shape mask
    // (never painted), and a multiply half lands just its colored content
    // inside that mask. Page color and lower members show through.
    let maskData: Uint8ClampedArray | undefined;
    for (let k = 0; k < run.length; k++) {
      const img = decoded[k];
      if (!img) continue;
      const dx = run[k].x - x0;
      const dy = run[k].y - y0;
      const dw = run[k].width;
      const dh = run[k].height;
      if (run[k].blend === "screen") {
        maskData = shapeMaskAt(img, dx, dy, dw, dh, width, height);
        continue;
      }
      if (run[k].blend === "multiply") {
        const content = maskedContent(img, dx, dy, dw, dh, maskData, width);
        maskData = undefined;
        if (!content) continue;
        c2d.drawImage(content, dx, dy, dw, dh);
        continue;
      }
      maskData = undefined;
      c2d.drawImage(img, dx, dy, dw, dh);
    }
    // The composite is a brand-new data URL: decode it through a DOM Image
    // first (the same protocol addPlainImage follows) — inserting a url
    // Leafer hasn't decoded rides the stage's eager render as an empty
    // bitmap that the stalled re-render never picks back up.
    const url = canvas.toDataURL("image/png");
    const el = new Image();
    el.onload = () => {
      // A repaint since the decode started cleared the tree (slot included)
      // — that repaint's own run now owns the paint-order slot.
      if (!slot.parent) return;
      pinImage(url);
      tree.addAfter(new LeaferImage({ url, x: boxX + x0, y: boxY + y0, width, height }), slot);
      tree.remove(slot);
      ctx.rerender();
    };
    el.src = url;
  });
}

/** A screen half's brightness as the shape mask, sampled on the run's union
 *  box: each pixel's alpha takes its max channel — the 1bpp white shape
 *  lights up, the black backdrop drops out. This is the raster-op's "where
 *  the shape is" term, derived from the record's own bytes. */
function shapeMaskAt(
  img: HTMLImageElement,
  dx: number,
  dy: number,
  dw: number,
  dh: number,
  width: number,
  height: number,
): Uint8ClampedArray {
  const c = document.createElement("canvas");
  c.width = width;
  c.height = height;
  const g = c.getContext("2d")!;
  g.drawImage(img, dx, dy, dw, dh);
  const d = g.getImageData(0, 0, width, height);
  const a = d.data;
  for (let i = 0; i < a.length; i += 4) a[i + 3] = Math.max(a[i], a[i + 1], a[i + 2]);
  return a;
}

/** A multiply half reduced to its colored content inside the pending shape
 *  mask: per pixel, white keeps the destination (alpha 0) and every other
 *  color lands verbatim at the mask's coverage — GDI's AND over a white page
 *  writes the source pixel itself wherever it is not white, not a blend. A
 *  distance-from-white ramp would turn light fills and antialiased ink into
 *  translucent washes over the page. An unconsumed multiply half falls back
 *  to its own non-white key. */
function maskedContent(
  img: HTMLImageElement,
  dx: number,
  dy: number,
  dw: number,
  dh: number,
  maskData: Uint8ClampedArray | undefined,
  maskWidth: number,
): HTMLCanvasElement | undefined {
  if (dw < 1 || dh < 1) return undefined;
  const c = document.createElement("canvas");
  c.width = dw;
  c.height = dh;
  const g = c.getContext("2d")!;
  g.drawImage(img, 0, 0, dw, dh);
  const d = g.getImageData(0, 0, dw, dh);
  const a = d.data;
  for (let j = 0; j < dh; j++) {
    for (let i = 0; i < dw; i++) {
      const p = (j * dw + i) * 4;
      const m =
        maskData && dy + j >= 0 && dy + j < maskData.length / 4 / maskWidth
          ? maskData[((dy + j) * maskWidth + dx + i) * 4 + 3]
          : 255;
      a[p + 3] = Math.min(a[p], a[p + 1], a[p + 2]) >= 250 ? 0 : m;
    }
  }
  g.putImageData(d, 0, 0);
  return c;
}

/** A cropped picture (a:srcRect): Leafer paints whole sources only, so the
 *  sub-region renders through an offscreen canvas copy, added when decoded —
 *  into the paint-order slot a placeholder kept open for it (see
 *  addPlainImage; the stage re-paints on the next sync regardless). Mirrors
 *  flip the cropped result (the xfrm flip applies to the blip, post-crop).
 *  Shared by drawing members and inline picture atoms. */
function addCroppedImage(
  tree: IGroup,
  src: string,
  crop: { left: number; top: number; right: number; bottom: number },
  x: number,
  y: number,
  width: number,
  height: number,
  ctx: PaintContext,
  flipH?: boolean,
  flipV?: boolean,
): void {
  const slot = new Rect({ x, y, width, height });
  tree.add(slot);
  const el = new Image();
  el.onload = () => {
    // A repaint since the decode started cleared the tree (slot included) —
    // that repaint's own decode now owns the paint-order slot.
    if (!slot.parent) return;
    const sx = Math.round(crop.left * el.naturalWidth);
    const sy = Math.round(crop.top * el.naturalHeight);
    const sw = Math.max(1, el.naturalWidth - sx - Math.round(crop.right * el.naturalWidth));
    const sh = Math.max(1, el.naturalHeight - sy - Math.round(crop.bottom * el.naturalHeight));
    const canvas = document.createElement("canvas");
    canvas.width = sw;
    canvas.height = sh;
    canvas.getContext("2d")?.drawImage(el, sx, sy, sw, sh, 0, 0, sw, sh);
    const croppedUrl = canvas.toDataURL("image/png");
    pinImage(croppedUrl);
    tree.addAfter(
      new LeaferImage({
        url: croppedUrl,
        x,
        y,
        width,
        height,
        // Mirrors flip around the element's (x,y) origin — same shift as
        // addPlainImage: move the origin to the far edge first.
        ...(flipH ? { x: x + width, scaleX: -1 } : {}),
        ...(flipV ? { y: y + height, scaleY: -1 } : {}),
      }),
      slot,
    );
    tree.remove(slot);
    // Same eager-render gap as addPlainImage.
    ctx.rerender();
  };
  el.src = src;
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
  // The shared walk: boundaries, the occupancy grid, and every cell's
  // content origin — the caret map consumes the same tableGridOf output.
  const { colX, rowY, occ, cells } = tableGridOf(table);
  const nRows = table.rows.length;
  const nCols = table.columnWidthsPx.length;

  for (const p of cells) {
    // Shading covers the merged box; content anchors to the start row (the
    // engine measured it there).
    if (p.cell.fill) {
      tree.add(
        new Rect({
          x: x + colX[p.col]!,
          y: y + rowY[p.row]!,
          width: colX[p.col + p.spanW]! - colX[p.col]!,
          height: rowY[p.row + p.spanH]! - rowY[p.row]!,
          fill: `#${p.cell.fill}`,
        }),
      );
    }
    const contentX = x + p.contentXPx;
    const contentY = y + p.contentYPx;
    for (const stacked of p.cell.stack) {
      paintBlock(tree, stacked.block, contentX, contentY + stacked.yPx, ctx, {
        width: p.cell.innerWidthPx,
        inCell: true,
      });
    }
  }

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

/** The font family a text slice paints in: the measurement side's slot pick
 *  (cssFontOf builds `, serif` on top of it — layout, caret map and paint
 *  must resolve empty slots to the same face or glyph advances drift from
 *  the boundaries the caret map computed). */
function familyOf(style: LayoutTextStyle, text: string): string {
  return familyOfSlot(style.family, isCjkCodeUnit(text, 0)) || "serif";
}
