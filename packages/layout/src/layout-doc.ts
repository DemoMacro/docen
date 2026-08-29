// LayoutDoc — the engine's own input projection. Adapters (docx from
// Tiptap/ProseMirror, pptx from its shape tree, xlsx from its grid) build
// these plain blocks; the engine never sees a Tiptap node or an OOXML part.
//
// Projection contract:
// - All geometry is px. Unit conversions happen once, in the adapter.
// - Style cascades are resolved. `spacing`/`indent`/text styles arrive as the
//   effective values (direct attrs → style chain → document defaults merged);
//   the engine has no styles table and no `styleId` — re-resolving a cascade
//   here would duplicate each format's semantics it must not know about.
// - Only layout semantics keep their OOXML shape (a line rule, a grid pitch,
//   a snap flag) — those are the rules this engine exists to implement.

import type { FontSlots } from "./font";

// ── inline content ──

/** Resolved run styling. `family` may be script-slotted; the engine itemizes
 *  a mixed Latin/CJK run by code point and picks the slot per segment. */
export interface LayoutTextStyle {
  family: string | FontSlots;
  sizePx: number;
  bold?: boolean;
  italic?: boolean;
  /** Text color, hex RRGGBB (w:color; absent/auto → the renderer's ink). */
  color?: string;
  /** Underlined (w:u, any non-none pattern). */
  underline?: boolean;
  /** Struck through (w:strike / w:dstrike). */
  strikethrough?: boolean;
  /** Extra per-character spacing in px (OOXML w:spacing resolved). */
  letterSpacingPx?: number;
  /** Raised/lowered run (w:vertAlign): glyphs paint at the scaled size on a
   *  shifted baseline (Word's FootnoteReference look). Measuring applies the
   *  same scaling — see vertAlignedSizePx. */
  verticalAlign?: "superscript" | "subscript";
}

/** One inline item of a paragraph. Text, hard breaks, tabs, and inline
 *  pictures are all line-box content — one box packer handles the four (the
 *  unified text+picture breaker the DOM route never had). A picture's `src`
 *  (data URL or object URL) is renderer-only passthrough — the engine measures
 *  the box and never loads it. A tab advances to the next stop: the explicit
 *  `toPx` (a numbering bullet's hop to the body text), the paragraph's
 *  `tabStops`, or the default grid (720 twips). */
/** a:srcRect crop as fractions of the image edge (0-1, each side inward);
 *  the painted region is the remainder. */
export interface LayoutPictureCrop {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export type LayoutInline =
  /** A `field` marker makes the text a dynamic page-number atom (w:fldSimple /
   *  complexField PAGE / NUMPAGES): the value only exists after pagination, so
   *  `text` is a single-digit placeholder for measuring and the painter swaps
   *  in the real page number. */
  | {
      kind: "text";
      text: string;
      style: LayoutTextStyle;
      field?: "page" | "numPages";
      /** Ids of the comments whose range covers this atom (w:commentRangeStart
       *  /commentRangeEnd): the painter tints the text's box (sorted, unique).
       *  Pure paint metadata — measuring and wrapping ignore it. */
      commentIds?: number[];
    }
  | { kind: "break" }
  | { kind: "tab"; toPx?: number }
  | {
      kind: "picture";
      widthPx: number;
      heightPx: number;
      src?: string;
      /** a:srcRect crop: the flat `src` paints only the visible remainder
       *  (Leafer paints whole sources, so the renderer sub-regions it). */
      crop?: LayoutPictureCrop;
      /** Vector replay members (a WMF/EMF metafile source): when present, the
       * renderer paints these instead of loading `src` — same renderer-only
       * passthrough contract; the engine never reads beyond widthPx/heightPx. */
      members?: LayoutDrawingMember[];
    };

// ── floating drawings (anchored shape groups) ──

/** One absolutely-positioned member of a floating drawing. Coordinates and
 *  sizes are px in the drawing's own box (top-left corner the origin) — the
 *  adapter already resolved the group's child coordinate space (chOff/chExt
 *  scaling). */
export type LayoutDrawingMember =
  | {
      kind: "picture";
      x: number;
      y: number;
      width: number;
      height: number;
      /** Renderer media source (data URL); absent → an empty frame. */
      src?: string;
      /** Mirrored content flips (a:xfrm @flipH/@flipV). */
      flipH?: boolean;
      flipV?: boolean;
      /** Source rectangle crop (a:srcRect): the painted region is the
       *  visible remainder of the image edge. */
      crop?: LayoutPictureCrop;
    }
  | {
      kind: "shape";
      x: number;
      y: number;
      width: number;
      height: number;
      /** Preset geometry (a:prstGeom @prst). The renderer maps the presets it
       *  knows and skips the rest; custom geometry stays unprojected. */
      preset?: string;
      /** Solid fill, hex RRGGBB; absent → no fill. */
      fill?: string;
      /** Fill opacity 0-1 (the solid color's a:alpha percent ÷ 100);
       *  absent → fully opaque. Fades the fill only, never the stroke. */
      opacity?: number;
      /** Outline stroke: width in px + hex color (absent color → ink). */
      line?: { px: number; color?: string };
    }
  | {
      kind: "path";
      x: number;
      y: number;
      width: number;
      height: number;
      /** SVG path data in box coordinates (0,0 … width,height) — the adapter
       *  scaled the geometry's own space (custGeom path w/h) into the box. */
      d: string;
      /** Fill rule for self-intersecting outlines (wmf SetPolyFillMode). */
      fillRule?: "evenodd" | "nonzero";
      /** Solid fill, hex RRGGBB; absent → no fill. */
      fill?: string;
      /** Outline stroke (a:ln): width px + hex color + cap/join/dash. */
      line?: {
        px: number;
        color?: string;
        cap?: "round" | "square" | "flat";
        join?: "round" | "bevel" | "miter";
        dash?: string;
      };
    }
  | {
      kind: "textBox";
      x: number;
      y: number;
      width: number;
      height: number;
      /** The shape's own solid fill, hex RRGGBB; absent → no fill (a plain
       *  wps:txbx draws its spPr fill under the text). */
      fill?: string;
      /** Fill opacity 0-1 (the solid color's a:alpha percent ÷ 100). */
      opacity?: number;
      /** The shape's outline stroke (a:ln): width px + hex color + the line-
       *  dressing tokens (cap/join full-word, dash the prstDash token). Word
       *  draws the txbx box even when the body is empty. */
      line?: {
        px: number;
        color?: string;
        cap?: "round" | "square" | "flat";
        join?: "round" | "bevel" | "miter";
        dash?: string;
      };
      /** Preset geometry (a:prstGeom @prst) — a txbx can live in any shape
       *  (a text-carrying ellipse); the box paints in that shape. */
      preset?: string;
      /** Text insets px (wps:bodyPr lIns/tIns/rIns/bIns, DrawingML defaults
       *  applied by the adapter). */
      insets?: { left?: number; top?: number; right?: number; bottom?: number };
      /** Vertical anchoring (bodyPr @anchor); absent = top (the OOXML
       *  ST_TextAnchoringType default). */
      anchor?: "top" | "center" | "bottom";
      /** bodyPr a:spAutoFit — the shape hugs its text, so vertical anchoring
       *  resolves against the fitted height, not the declared extent (Word
       *  shrinks a stale oversized cy to the text's line height). */
      autoFit?: boolean;
      /** Paint the runs as one unwrapped line at the box origin — metafile
       *  text (GDI ExtTextOut / EMF+ DrawString) draws the string as-is and
       *  never word-wraps, so the width re-fit must not re-break it. */
      nowrap?: boolean;
      /** Clockwise rotation of the text body about the box origin, degrees —
       *  metafile vertical text: a rotated GDI world transform lays runs down
       *  a column; shaping stays horizontal, the paint rotates. */
      rotation?: number;
      /** The box's content as projected blocks — the renderer stacks them
       *  inside the box (the flow never sees them; the drawing wraps none). */
      blocks: LayoutBlock[];
    };

/** A floating shape group (wpg:wgp under a wp:anchor): a drawing box anchored
 *  to its paragraph, members absolutely positioned inside. The flow gives the
 *  anchor paragraph no extra height — the anchor's `relative` axes resolve to
 *  page geometry at paint time (the painter owns the page box).
 *
 *  Word's eight relativeFrom values collapse onto these four semantic axes
 *  per side; the axes the adapter cannot resolve yet (character/line anchor
 *  points, inside/outside under mirrored margins) land on their nearest
 *  axis, a registered gap. */
export interface LayoutDrawingAnchor {
  horizontal: {
    /** column = the content box; leftMargin/rightMargin = its edges; page =
     *  the page box. margin/insideMargin → column, character → column,
     *  outsideMargin → rightMargin (unmirrored). */
    relative: "column" | "leftMargin" | "rightMargin" | "page";
    /** px from the reference box's left edge (posOffset), or a fraction of
     *  the reference box's width (percentOffset / 1000). Exclusive with align. */
    offsetPx?: number;
    percent?: number;
    /** Alignment inside the reference box (inside→left, outside→right). */
    align?: "left" | "center" | "right";
  };
  vertical: {
    /** paragraph = the anchor paragraph's top; topMargin/bottomMargin = the
     *  content box's edges; page = the page box. margin/insideMargin →
     *  topMargin, line → paragraph, outsideMargin → bottomMargin. */
    relative: "paragraph" | "topMargin" | "bottomMargin" | "page";
    offsetPx?: number;
    percent?: number;
    align?: "top" | "center" | "bottom";
  };
}

export interface LayoutDrawing {
  anchor: LayoutDrawingAnchor;
  width: number;
  height: number;
  members: LayoutDrawingMember[];
  /** w:wrap — how text flows around the box. Absent (wrapNone) paints over
   *  or under the text without affecting the flow; "square"/"tight" shrink
   *  the lines the box overlaps (tight reduces to square's rectangle — a
   *  contour pass is a registered gap); "topAndBottom" clears the band. */
  wrap?: "square" | "tight" | "topAndBottom";
  /** ST_WrapSide (w:wrap @side): which side of the box takes text. "right"
   *  (or "largest" with the wider right side) moves the wrapped lines past
   *  the box's right edge; "left"/"both" pack from the left as usual. */
  wrapSide?: "both" | "left" | "right" | "largest";
  /** wrapTight/through contour (wp:wrapPolygon points) in px, relative to
   *  the drawing box's top-left — the adapter scaled them out of Word's
   *  21600×21600 polygon space. Registered on the zone; the line breaker
   *  slices it per line (a bounding-box approximation without it). */
  contour?: { x: number; y: number }[];
  /** w:behindDoc — painted under the text layer (text strokes stay visible). */
  behind?: boolean;
  /** w:anchor distL/distT/distR/distB in px — how far wrapping text keeps
   *  its distance from the box. square/tight zones widen by left+right,
   *  topAndBottom bands by top+bottom; wrapNone never reads them. */
  distances?: { left?: number; top?: number; right?: number; bottom?: number };
}

/** A drawing's wrap box: its extent grown by the anchor's wrap distances,
 *  clipped to the column. Both zone builders (the flow's post-anchor zones
 *  and the anchor paragraph's own self-zones) share this so the square/
 *  tight/topAndBottom paddings stay in lockstep. `boxTopPx` is the drawing's
 *  own top in the caller's Y space; the distances pad top and bottom around
 *  it. `inCell` applies Word's layoutInCell containment: a cell-anchored
 *  object never extends past its cell — an offset that overflows the cell's
 *  right edge shifts the whole box left to touch it (the wrap zone follows
 *  the shifted box). Body floats keep their raw offset (they may hang into
 *  the margins). Returns undefined when the padded box misses the column
 *  entirely. */
export function drawingWrapBox(
  d: LayoutDrawing,
  boxTopPx: number,
  columnWidth: number,
  inCell = false,
):
  | {
      widthPx: number;
      topPx: number;
      bottomPx: number;
      x0Px: number;
      textAfter?: boolean;
      contour?: { x: number; y: number }[];
    }
  | undefined {
  const left = d.distances?.left ?? 0;
  const offset = inCell
    ? Math.min(Math.max(d.anchor.horizontal.offsetPx ?? 0, 0), Math.max(0, columnWidth - d.width))
    : (d.anchor.horizontal.offsetPx ?? 0);
  const x0 = offset - left;
  const start = Math.max(x0, 0);
  const widthPx =
    Math.min(offset + d.width + left + (d.distances?.right ?? 0), columnWidth) - start;
  if (widthPx <= 0) return undefined;
  // Which side takes the text: "right" forces the right side, "left" keeps
  // the text left no matter what, and both/largest take the wider side —
  // Word's two-sided wrap approximated per line (a line packs on one side).
  const rightSpace = columnWidth - start - widthPx;
  const textAfter =
    d.wrapSide === "right" || (d.wrapSide !== "left" && rightSpace > start) || undefined;
  // The contour rides along in zone coordinates: the drawing box's own
  // top-left sits at (offset-start, distTop) inside the padded zone box.
  const contour = d.contour?.map((p) => ({
    x: p.x + offset - start,
    y: p.y + (d.distances?.top ?? 0),
  }));
  return {
    widthPx,
    topPx: boxTopPx - (d.distances?.top ?? 0),
    bottomPx: boxTopPx + d.height + (d.distances?.bottom ?? 0),
    x0Px: start,
    ...(textAfter ? { textAfter } : {}),
    ...(contour ? { contour } : {}),
  };
}

/** A paragraph's wrapping drawings as flow effects in the caller's Y space
 *  (`baseY` = the anchor paragraph's top; 0 for paragraph-relative zones).
 *  Zones shrink the lines they overlap; bands (topAndBottom, or a square box
 *  covering the full column) clear everything in their band — callers that
 *  cannot dodge a band mid-stack (a paragraph's own lines, a table cell)
 *  drop them. Column/margin/page-anchored and wrapNone drawings produce
 *  neither. */
export function wrapEffectsOf(
  drawings: readonly LayoutDrawing[] | undefined,
  baseY: number,
  columnWidth: number,
  inCell = false,
): { zones: LayoutFloatZone[]; bands: LayoutFloatZone[] } {
  const zones: LayoutFloatZone[] = [];
  const bands: LayoutFloatZone[] = [];
  for (const d of drawings ?? []) {
    if (!d.wrap) continue;
    const { horizontal: h, vertical: v } = d.anchor;
    if (v.relative !== "paragraph" || h.relative !== "column") continue;
    const box = drawingWrapBox(d, baseY + (v.offsetPx ?? 0), columnWidth, inCell);
    if (!box) continue;
    const zone: LayoutFloatZone = {
      widthPx: box.widthPx,
      topPx: box.topPx,
      bottomPx: box.bottomPx,
      x0Px: box.x0Px,
      ...(box.textAfter ? { textAfter: true } : {}),
      ...(box.contour ? { contour: box.contour } : {}),
    };
    if (d.wrap === "topAndBottom" || box.widthPx >= columnWidth - 1) bands.push(zone);
    else zones.push(zone);
  }
  return { zones, bands };
}

// ── blocks ──

/** OOXML w:spacing/@w:line + @w:lineRule, unit-resolved by the adapter. */
export type LayoutLineHeight =
  | { rule: "exact"; px: number }
  | { rule: "atLeast"; px: number }
  | { rule: "multiple"; factor: number }; // w:line in 240ths of a single line

/** Paragraph spacing (w:spacing), all px. `before`/`after` are data for the
 *  flow: vertical margin collapse between siblings is applied by the caller
 *  that stacks blocks (the flow for body paragraphs, the cell for table
 *  content — a cell is a BFC that eats the first `before` and last `after`). */
export interface LayoutSpacing {
  lineHeight?: LayoutLineHeight;
  beforePx: number;
  afterPx: number;
}

/** Paragraph indents (w:ind), px. `firstLinePx` shrinks only the first line's
 *  wrapping width (CSS text-indent); negative values are hanging indents (a
 *  numbering bullet's line reaches left of the body text — the first line is
 *  WIDER than the rest). The adapter resolved twips — and firstLineChars
 *  (chars/100 × font size) — into px. */
export interface LayoutIndent {
  leftPx?: number;
  rightPx?: number;
  firstLinePx?: number;
}

/** One w:tab stop, px from the content-box left edge. */
export interface LayoutTabStop {
  positionPx: number;
  type: "left" | "center" | "right";
  /** w:leader — the fill drawn across the tab's advance ("none" dropped). */
  leader?: "dot" | "heavy" | "hyphen" | "middleDot" | "underscore";
}

/** One w:pBdr edge: a border line drawn beside the paragraph. */
export interface LayoutParagraphBorderEdge extends LayoutBorderEdge {
  /** w:space — offset from the text to the line, pt→px. */
  spacePx?: number;
}

export interface LayoutParagraph {
  kind: "paragraph";
  inline: LayoutInline[];
  spacing?: LayoutSpacing;
  indent?: LayoutIndent;
  /** Explicit tab stops (w:tabs), px from the content-box left edge; tabs
   *  beyond the last stop fall to the default grid (720 twips). */
  tabStops?: LayoutTabStop[];
  /** Paragraph borders (w:pBdr) — the painter draws them beside the block. */
  borders?: Partial<Record<"top" | "right" | "bottom" | "left", LayoutParagraphBorderEdge>>;
  /** ¶-mark strut size in px (w:pPr/w:rPr/w:sz resolved): an ABSOLUTE line
   *  height for the paragraph-mark line — the sole content of an empty
   *  paragraph, and the minimum of a picture row shorter than a text line. */
  markSizePx?: number;
  /** Default run style (the style chain's run, resolved): the strut font when
   *  the paragraph has no text runs. */
  defaultTextStyle?: LayoutTextStyle;
  /** w:snapToGrid: absent/null = engine default (snap when a grid pitch is
   *  active); explicit false drops the grid pitch. */
  snapToGrid?: boolean | null;
  /** OOXML pagination controls, already resolved through the style cascade
   *  (a heading defaults keepNext=true — the adapter decides that). */
  keepLines?: boolean;
  /** Horizontal alignment (w:jc, resolved through the style cascade). "both"
   *  (justify) stretches every WRAPPED line's inter-character gaps to the
   *  full content width — the paragraph's last line and hard-break lines
   *  keep their natural width; "distribute" stretches every line, the last
   *  included. "center"/"right" shift each line's items by its slack
   *  (trailing whitespace hangs and never counts). */
  align?: "left" | "center" | "right" | "both" | "distribute";
  keepNext?: boolean;
  widowControl?: boolean;
  pageBreakBefore?: boolean;
  /** Floating drawings anchored to this paragraph: wrap-none boxes paint at
   *  their offset; a `wrap` on the drawing also shrinks the anchor
   *  paragraph's own lines around the box and registers a float zone the
   *  flow applies to later paragraphs (or a cleared band for topAndBottom). */
  drawings?: LayoutDrawing[];
}

/** Cell insets in px: per-cell w:tcMar, or the table's w:tblCellMar default
 *  (cells inherit per side — the adapter or the engine resolves cell ?? table). */
export interface LayoutCellInsets {
  top?: number;
  right?: number;
  bottom?: number;
  left?: number;
}

/** One border edge of a cell: nil/none/absent sides carry no width. The
 *  visual default border (the DOM route's 1px Table-Grid stamp) is a renderer
 *  decision injected by the adapter — the engine measures only declared edges. */
export interface LayoutBorderEdge {
  style?: string; // "nil"/"none" → no width
  px?: number; // resolved from w:sz (eighths of a point)
  /** Hex RRGGBB (OOXML w:color); absent/auto → the renderer's ink default. */
  color?: string;
}

/** CT_TblBorders: the table-level edges a cell's own w:tcBorders side falls
 *  back to — outer edges for the grid's rim, inside edges for shared lines. */
export interface LayoutTableBorders {
  top?: LayoutBorderEdge;
  bottom?: LayoutBorderEdge;
  left?: LayoutBorderEdge;
  right?: LayoutBorderEdge;
  insideHorizontal?: LayoutBorderEdge;
  insideVertical?: LayoutBorderEdge;
}

export interface LayoutTableCell {
  colspan?: number;
  rowspan?: number;
  /** Per-spanned-column widths in px (w:tcW resolved); absent → grid share. */
  widthPx?: number;
  insets?: LayoutCellInsets;
  borders?: Partial<Record<"top" | "right" | "bottom" | "left", LayoutBorderEdge>>;
  /** Cell shading (w:shd @w:fill), hex RRGGBB. */
  fill?: string;
  /** w:vAlign — the content's placement when the row is taller than it. */
  verticalAlign?: "top" | "center" | "bottom";
  blocks: LayoutBlock[];
}

export interface LayoutTableRow {
  cells: LayoutTableCell[];
  /** w:trHeight resolved: atLeast floors the row, exact fixes it (content
   *  overflows but the row does not grow). */
  height?: { rule: "atLeast" | "exact"; px: number };
  /** w:tblHeader — leading rows repeat on every page the table splits onto
   *  (only a contiguous prefix from the first row counts). */
  tableHeader?: boolean;
  /** w:cantSplit — the row moves whole to the next page instead of splitting
   *  mid-content (a row taller than a page still force-splits — Word clips
   *  nothing). */
  cantSplit?: boolean;
}

export type LayoutTableWidth =
  | { type: "percent"; percent: number } // 0-100
  | { type: "px"; px: number }; // w:tblW dxa resolved

export interface LayoutTable {
  kind: "table";
  /** Absent width = auto: fill the containing flow width. */
  width?: LayoutTableWidth;
  /** w:tblPr/w:jc — the table box's placement inside the flow column. A table
   *  wider than the column centers into the margins (negative offset). */
  align?: "left" | "center" | "right";
  /** tblGrid column widths in px, scaled proportionally to the effective
   *  table width (Word scales the grid to tblW, never to the raw sum). */
  columnWidthsPx?: number[];
  /** Table-level default insets (w:tblCellMar) a cell without its own w:tcMar
   *  inherits, per side. */
  cellInsets?: LayoutCellInsets;
  /** Table-level border defaults (w:tblBorders, style chain resolved):
   *  the renderer falls a cell's missing edge back to these per side. */
  borders?: LayoutTableBorders;
  rows: LayoutTableRow[];
}

/** A plain container (list body, blockquote) — laid out by recursion, no
 *  geometry of its own. */
export interface LayoutGroup {
  kind: "group";
  blocks: LayoutBlock[];
}

/** An opaque block the adapter cannot lay out (raw XML passthrough, TOC,
 *  altChunk…): the flow reserves an estimated box so the content keeps a
 *  visual presence instead of silently vanishing. The renderer draws the
 *  label; the engine only moves the box. */
export interface LayoutPlaceholder {
  kind: "placeholder";
  /** Estimated height in px (the adapter's guess — usually N default lines). */
  heightPx: number;
  /** What the box stands for, shown by the renderer (e.g. "toc", "rawXml"). */
  label?: string;
}

/** A page/column break atom: zero height, closes the current flow box after
 *  the preceding content (the break never opens a page — Word semantics). */
export interface LayoutPageBreak {
  kind: "pageBreak";
}

export type LayoutBlock =
  | LayoutParagraph
  | LayoutTable
  | LayoutGroup
  | LayoutPlaceholder
  | LayoutPageBreak;

// ── flow context ──

/** A floating picture's flow band: text lines overlapping [topPx, bottomPx)
 *  wrap beside it (usable width reduced by widthPx). wrapNone/page/margin
 *  anchors never produce a zone — they sit outside the text flow. */
export interface LayoutFloatZone {
  widthPx: number;
  topPx: number;
  bottomPx: number;
  /** Line-content x of the zone's left edge — pairs with `textAfter` to move
   *  the line's text right of the box (wrapSide right/largest). */
  x0Px?: number;
  /** Text packs RIGHT of the box: the line's usable interval starts past
   *  x0Px + widthPx instead of packing from the left margin. */
  textAfter?: boolean;
  /** Tight/through contour polygon in zone coordinates (the drawing box's
   *  own contour translated by the wrap distances) — the line breaker
   *  slices it at each line's mid-height instead of using the box. */
  contour?: { x: number; y: number }[];
}

/** Block-level layout context threaded by the caller that stacks blocks. */
export interface LayoutBlockContext {
  /** Section document-grid pitch in px (w:docGrid linePitch); 0/absent = no
   *  grid. Body CJK lines ceil to a whole pitch multiple; Latin lines floor
   *  at max(natural, pitch); table cells never ceil — the pitch scales a
   *  multiple rule and floors a snapped one (adjustLineHeightInTable). */
  linePitchPx?: number;
  /** The body flow centers a grid-height line's natural text box in the
   *  span (the docGrid lattice — Word-verified). Text-box stacks share the
   *  rule for their grid-snapped lines (half-leading like the body);
   *  header/footer stacks are laid with no grid context at all (natural
   *  line heights) and so never set this. */
  onGrid?: boolean;
  /** True inside a table cell: lines never ceil to whole grid rows (the row's
   *  trHeight floors separately — the w:adjustLineHeightInTable pitch is a
   *  floor, not a row count). Cells also clear float zones — a cell's width
   *  is its column, not the page flow. */
  inTable?: boolean;
  floatZones?: readonly LayoutFloatZone[];
  /** This block's top Y within the flow — pairs with floatZones to derive
   *  each line's band. */
  startY?: number;
}
