import {
  familyOfSlot,
  gridPadOf,
  isCjkCodeUnit,
  justifiedIntervals,
  justifyPerGrapheme,
  lineOriginXPx,
  lineSpaceGaps,
  vertAlignedSizePx,
  vertAlignBaselineShiftPx,
  type LaidOutLine,
  type LaidOutLineItem,
  type LaidOutParagraph,
  type LayoutInline,
  type LayoutParagraphBorderEdge,
  type LayoutTextStyle,
} from "@docen/layout";
import {
  Box,
  Ellipse,
  Group,
  Image as LeaferImage,
  Line,
  Path,
  Rect,
  Text,
  type IGroup,
} from "leafer-ui";

import type { PaintColumn, PaintContext } from "./context";
import { paintDrawing, paintMembers, recordDrawingHit } from "./drawing";
import { addCroppedImage, pinImage } from "./image";

/** OOXML ST_HighlightColor tokens → Word's highlight palette, #RRGGBB. */
const HIGHLIGHT_COLOR: Record<string, string> = {
  yellow: "#FFFF00",
  green: "#00FF00",
  cyan: "#00FFFF",
  magenta: "#FF00FF",
  blue: "#0000FF",
  red: "#FF0000",
  darkBlue: "#000080",
  darkYellow: "#808000",
  darkGreen: "#008000",
  darkRed: "#800000",
  darkCyan: "#008080",
  darkMagenta: "#800080",
  lightGray: "#C0C0C0",
  darkGray: "#808080",
  black: "#000000",
  white: "#FFFFFF",
};

/** The w:u pattern to hand-stroke — undefined for the plain single line (and
 *  for runs with no underline at all), which stay on Leafer's textDecoration. */
function underlinePatternOf(style: LayoutTextStyle): string | undefined {
  if (!style.underline) return undefined;
  const s = style.underlineStyle;
  return !s || s === "single" ? undefined : s;
}

/** w:u dash patterns in px (Leafer dashPattern stroke-gap pairs). */
const UNDERLINE_DASHES: Record<string, number[] | undefined> = {
  dotted: [1, 2],
  dottedHeavy: [2, 2],
  dash: [3, 2],
  dashedHeavy: [4, 2],
  dashLong: [6, 2],
  dashLongHeavy: [8, 2],
  dotDash: [1, 2, 3, 2],
  dashDotHeavy: [2, 2, 4, 2],
  dotDotDash: [1, 2, 1, 2, 3, 2],
  dashDotDotHeavy: [2, 2, 2, 2, 4, 2],
};

/** Stroke a patterned/colored w:u under one inline item. Coordinates ride the
 *  Text element's own frame (y is its box top); the baseline sits at the
 *  shared ~0.85 em ascent approximation (no ascent on the layout type). */
function paintUnderlinePattern(
  tree: IGroup,
  pattern: string,
  style: LayoutTextStyle,
  box: { x: number; y: number; width: number; emPx: number },
): void {
  const color = style.underlineColor
    ? `#${style.underlineColor}`
    : style.color
      ? `#${style.color}`
      : "#1b1b1b";
  const heavy = pattern.endsWith("Heavy") || pattern === "thick";
  const strokeWidth = heavy ? 2 : 1;
  const baseline = box.y + box.emPx * 0.85;
  const wave = pattern.startsWith("wave");
  if (!wave) {
    const dash = UNDERLINE_DASHES[pattern];
    tree.add(
      new Line({
        x: box.x,
        y: baseline + box.emPx * 0.08,
        width: box.width,
        stroke: color,
        strokeWidth,
        dashPattern: dash,
        hittable: false,
      }),
    );
    if (pattern === "double") {
      tree.add(
        new Line({
          x: box.x,
          y: baseline + box.emPx * 0.08 + 2,
          width: box.width,
          stroke: color,
          strokeWidth,
          hittable: false,
        }),
      );
    }
    return;
  }
  // Sine wave as quadratic half-waves (4px per lobe, ~1/16 em amplitude).
  const amp = heavy || pattern === "wavyDouble" ? 1.5 : 1;
  const lobe = 4;
  let d = "M 0 0";
  let px = 0;
  let dir = 1;
  while (px < box.width) {
    const step = Math.min(lobe, box.width - px);
    d += ` Q ${(px + step / 2).toFixed(2)} ${(dir * amp).toFixed(2)} ${(px + step).toFixed(2)} 0`;
    px += step;
    dir = -dir;
  }
  tree.add(
    new Path({
      x: box.x,
      y: baseline + box.emPx * 0.08,
      path: d,
      stroke: color,
      strokeWidth,
      fill: "none",
      hittable: false,
    }),
  );
  if (pattern === "wavyDouble") {
    tree.add(
      new Path({
        x: box.x,
        y: baseline + box.emPx * 0.08 + 3,
        path: d,
        stroke: color,
        strokeWidth,
        fill: "none",
        hittable: false,
      }),
    );
  }
}

export function paintParagraph(
  tree: IGroup,
  para: LaidOutParagraph,
  x: number,
  y: number,
  ctx: PaintContext,
  col?: PaintColumn,
): void {
  // The paragraph's right edge: the column/cell width when one was threaded
  // down (a table cell's content box, a w:cols column), else the full content
  // width. Shading, borders and the marks rules all stop here.
  const boxRight = x + (col?.width ?? ctx.flow.contentWidthPx);
  // The stage composes a page in two passes; this one paints only its own
  // layer's drawings. Behind-doc floats land beneath the furniture pass.
  const behind = ctx.layer === "behind";
  if (behind) {
    // Paragraph shading (w:shd) fills the block box beneath everything else
    // the paragraph paints — behind-doc floats included, matching Word.
    if (para.shadingFill) {
      tree.add(
        new Rect({
          x,
          y,
          width: Math.max(0, boxRight - x),
          height: para.heightPx,
          fill: `#${para.shadingFill}`,
        }),
      );
    }
    let index = 0;
    for (const drawing of para.drawings ?? []) {
      const host = { para, index: index++ };
      if (!drawing.behind) continue;
      // Deferred like the body band: the stage sorts the queue by
      // relativeHeight, so same-band stacking follows the z-order.
      const paint = (): void => paintDrawing(tree, drawing, x, y, ctx, col, host);
      if (ctx.deferredDrawings)
        ctx.deferredDrawings.push({ z: drawing.zIndex ?? 0, layer: "behind", paint });
      else paint();
    }
    return;
  }
  // w:pBdr: horizontal rules (the "Education" underline shape) span the
  // wrapping width between the text and `spacePx` of it; the vertical rails
  // run the paragraph's height, extended by the horizontal edges' space so a
  // four-side box reads closed. Added AFTER the line loop so Word's paint
  // order holds: borders ride above text and character highlights, below
  // the paragraph's own floating drawings.
  const paintBorders = (): void => {
    const live = (
      e: LayoutParagraphBorderEdge | undefined,
    ): e is LayoutParagraphBorderEdge & { px: number } =>
      !!e?.px && e.style !== "nil" && e.style !== "none";
    const { top, right, bottom, left } = para.borders ?? {};
    if (live(top)) {
      tree.add(
        new Rect({
          x,
          y: y - (top.spacePx ?? 0),
          width: Math.max(0, boxRight - x),
          height: top.px,
          fill: "#1b1b1b",
        }),
      );
    }
    if (live(bottom)) {
      tree.add(
        new Rect({
          x,
          y: y + para.heightPx + (bottom.spacePx ?? 0) - bottom.px,
          width: Math.max(0, boxRight - x),
          height: bottom.px,
          fill: "#1b1b1b",
        }),
      );
    }
    if (live(left) || live(right)) {
      const y0 = y - (live(top) ? (top.spacePx ?? 0) : 0);
      const height =
        para.heightPx +
        (live(top) ? (top.spacePx ?? 0) : 0) +
        (live(bottom) ? (bottom.spacePx ?? 0) : 0);
      if (live(left)) {
        tree.add(
          new Rect({ x: x - (left.spacePx ?? 0), y: y0, width: left.px, height, fill: "#1b1b1b" }),
        );
      }
      if (live(right)) {
        tree.add(
          new Rect({
            x: boxRight + (right.spacePx ?? 0),
            y: y0,
            width: right.px,
            height,
            fill: "#1b1b1b",
          }),
        );
      }
    }
  };
  let inlinePicIndex = 0;
  // Space-dot bookkeeping: the laid items are matched against the paragraph's
  // concatenated run text in order (pretext consumed the spaces between
  // them), so each gap's whitespace is known as the walk advances. Carried
  // across the paragraph's lines — a wrapped line's trailing spaces skip
  // before the next line's first item.
  const marks: ParagraphMarkState | null = ctx.showMarks ? paragraphMarkState(para) : null;
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
        // A squeezed line (advanceScale — Word's compressPunctuation)
        // compresses its pure-CJK glyphs to the item's already-scaled width:
        // the same both-letter uniform distribution the justified path
        // stretches with, run at negative slack (Leafer accepts it).
        const squeezePx =
          intervalPx == null && line.advanceScale != null ? item.widthPx : undefined;
        // Character highlight (w:highlight): the token's palette color fills
        // the run's box beneath the glyphs (Word paints it opaque). A run
        // shading (w:shd) fills the same box with an arbitrary color when no
        // highlight is present — OOXML precedence puts the highlight on top.
        const hl = inline.style.highlight ? HIGHLIGHT_COLOR[inline.style.highlight] : undefined;
        const runFill =
          hl ?? (inline.style.shadingFill ? `#${inline.style.shadingFill}` : undefined);
        if (runFill) {
          tree.add(
            new Rect({
              x: lineX + item.xPx,
              y: lineY + pad,
              width: intervalPx ?? item.widthPx,
              height: Math.max(1, line.naturalPx || line.heightPx),
              fill: runFill,
            }),
          );
        }
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
        // Two-lines-in-one (双行合一 / 合并字符): the atom packs its text
        // into two half-size lines within this line box, optionally wrapped
        // in bracket glyphs — Word's look; the item skips the normal run
        // paint below.
        if (item.combine) {
          const size = vertAlignedSizePx(inline.style) / 2;
          const boxH = Math.max(line.naturalPx || line.heightPx, size * 2);
          const rows = [item.combine.first, item.combine.second];
          const pair = item.combine.bracket
            ? { round: "()", square: "[]", angle: "<>", curly: "{}" }[item.combine.bracket]
            : undefined;
          const inset = pair ? size * 0.35 : 0;
          for (let row = 0; row < 2; row++) {
            if (!rows[row]) continue;
            tree.add(
              new Text({
                x: lineX + item.xPx + inset,
                y: lineY + pad + (row * boxH) / 2,
                width: Math.max(1, item.widthPx - inset * 2),
                textWrap: "none",
                height: Math.max(1, size),
                textAlign: "center",
                text: rows[row],
                fill: inline.style.color ? `#${inline.style.color}` : "#1b1b1b",
                fontFamily: family,
                fontSize: size,
                lineHeight: size,
                fontWeight: inline.style.bold ? 700 : 400,
                italic: inline.style.italic,
                hittable: false,
              }),
            );
          }
          if (pair) {
            const ink = inline.style.color ? `#${inline.style.color}` : "#1b1b1b";
            const brackets: [string, number][] = [
              [pair[0], lineX + item.xPx],
              [pair[1], lineX + item.xPx + item.widthPx - inset * 2],
            ];
            for (const [glyph, gx] of brackets) {
              tree.add(
                new Text({
                  x: gx,
                  y: lineY + pad,
                  width: Math.max(1, inset * 2),
                  textWrap: "none",
                  height: Math.max(1, boxH),
                  textAlign: "center",
                  text: glyph,
                  fill: ink,
                  fontFamily: family,
                  fontSize: vertAlignedSizePx(inline.style),
                  lineHeight: boxH,
                  hittable: false,
                }),
              );
            }
          }
          continue;
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
          // the shared vertAlignedSizePx so measure and paint agree. A ruby
          // base sinks below the annotation space reserved at the line top.
          y: lineY + pad + (item.rubyLiftPx ?? 0) + vertAlignBaselineShiftPx(inline.style),
          // width ONLY on justified/squeezed items (their stretch interval
          // or compressed width): a width on every line would let Leafer
          // wrap the slice again with its own metrics (a phantom second
          // line). textWrap "none" keeps the interval from wrapping; height
          // keeps the element paintable (height 0 is skipped by Leafer).
          width: intervalPx ?? squeezePx,
          textWrap: intervalPx != null || squeezePx != null ? "none" : undefined,
          // CJK items spread per glyph (both-letter); Latin items spread
          // per word gap (both-justify — Leafer's word mode, Word's Latin
          // justification). "both" keeps the single-row Text justifiable —
          // and compresses when the interval is narrower than the glyphs
          // (the squeeze path).
          textAlign: rights
            ? justifyPerGrapheme(item.text)
              ? "both-letter"
              : "both-justify"
            : squeezePx != null
              ? "both-letter"
              : undefined,
          height: Math.max(1, line.heightPx),
          text: label,
          fill: inline.style.color ? `#${inline.style.color}` : "#1b1b1b",
          // Leafer's textDecoration only knows the single line — a patterned
          // or colored w:u strokes its own path below (paintUnderlinePattern).
          // Single keeps the native path: the 91-page parity baseline rides on
          // its metrics.
          textDecoration: inline.style.strikethrough
            ? underlinePatternOf(inline.style)
              ? "delete"
              : inline.style.underline
                ? "under-delete"
                : "delete"
            : underlinePatternOf(inline.style)
              ? undefined
              : inline.style.underline
                ? "under"
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
        // The phonetic guide (w:ruby): the annotation fills the space
        // reserved above the base glyphs, centered within the base's width —
        // Word's default; the other ST_RubyAlign tokens shift the same box.
        if (item.ruby) {
          const size = item.ruby.fontSizePx;
          tree.add(
            new Text({
              x: lineX + item.xPx,
              y: lineY + pad,
              width: item.widthPx,
              textWrap: "none",
              height: Math.max(1, size),
              textAlign: "center",
              text: item.ruby.text,
              fill: inline.style.color ? `#${inline.style.color}` : "#1b1b1b",
              fontFamily: family,
              fontSize: size,
              lineHeight: size,
              fontWeight: inline.style.bold ? 700 : 400,
              italic: inline.style.italic,
              hittable: false,
            }),
          );
        }
        const pattern = underlinePatternOf(inline.style);
        if (pattern) {
          paintUnderlinePattern(tree, pattern, inline.style, {
            x: lineX + item.xPx,
            y: lineY + pad + (item.rubyLiftPx ?? 0) + vertAlignBaselineShiftPx(inline.style),
            width: intervalPx ?? item.widthPx,
            emPx: vertAlignedSizePx(inline.style),
          });
        }
      } else if (item.kind === "math" && inline.kind === "math") {
        // A formula the engine does not lay out yet: a dashed slot with the
        // structural label centered inside — Word's empty-argument look, an
        // honest stand-in until the math layout engine lands.
        tree.add(
          new Rect({
            x: lineX + item.xPx,
            y: lineY + pad,
            width: item.widthPx,
            height: item.heightPx,
            fill: "rgba(149, 166, 190, 0.12)",
            stroke: "#9aa6be",
            strokeWidth: 1,
            dashPattern: [3, 2],
            hittable: false,
          }),
        );
        tree.add(
          new Text({
            x: lineX + item.xPx,
            y: lineY + pad,
            width: item.widthPx,
            height: item.heightPx,
            text: item.label,
            fill: "#5b6675",
            fontFamily: "Inter, sans-serif",
            fontSize: item.heightPx * 0.62,
            italic: true,
            textAlign: "center",
            verticalAlign: "middle",
            hittable: false,
          }),
        );
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
          ...(inline.rotation ? { rotation: inline.rotation } : {}),
        });
        // A rotated picture spins about the extent's center — a group parked
        // at the center carries the angle (the floating drawing's pivot) and
        // the content re-offsets so the box stays centered under it.
        const picX = lineX + item.xPx;
        const picY = lineY + pad;
        let target: IGroup = tree;
        let ox = picX;
        let oy = picY;
        if (inline.rotation) {
          const spinner = new Group({
            x: picX + item.widthPx / 2,
            y: picY + item.heightPx / 2,
            rotation: inline.rotation,
          });
          tree.add(spinner);
          target = spinner;
          ox = -item.widthPx / 2;
          oy = -item.heightPx / 2;
        }
        if (inline.members) {
          // A metafile source replayed into members (WMF vector layers): the
          // structured scene paints in place of the flat image, clipped to
          // the extent — a srcRect leaves records reaching past the box and
          // GDI never lets metafile ink out of the playback rect. Leafer's
          // Group ignores `overflow` (a Box-only data getter clips children),
          // so the clip holder must be a Box.
          const holder = new Box({
            x: ox,
            y: oy,
            width: item.widthPx,
            height: item.heightPx,
            overflow: "hide",
          });
          paintMembers(holder, inline.members, 0, 0, ctx);
          target.add(holder);
        } else if (inline.src && inline.crop) {
          // A cropped flat source (a:srcRect): the visible remainder fills
          // the extent box — the whole source would stretch into it.
          addCroppedImage(
            target,
            inline.src,
            inline.crop,
            ox,
            oy,
            item.widthPx,
            item.heightPx,
            ctx,
          );
        } else if (inline.src) {
          pinImage(inline.src);
          target.add(
            new LeaferImage({
              url: inline.src,
              x: ox,
              y: oy,
              width: item.widthPx,
              height: item.heightPx,
            }),
          );
        } else {
          // Linked-only picture (no bytes in the package): an empty frame.
          target.add(
            new Rect({
              x: ox,
              y: oy,
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
          // metrics come from the line's dominant run — the tab atom carries
          // no style of its own.
          const { sizePx, color } = dominantRunOf(para, line, 0, "#1b1b1b");
          if (sizePx > 0) paintTabLeader(tree, item, lineX, lineY, pad, sizePx, color);
        }
      }
    }
    // Formatting marks (Word's ¶ toggle) — drawn per line after its content,
    // above the glyphs, in the text's own color.
    if (ctx.showMarks && marks)
      paintLineMarks(tree, para, line, lineX, lineY, pad, ctx, marks, boxRight);
  }
  paintBorders();
  // Floating drawings anchored to this paragraph: wrap-none boxes painted
  // over the text — the flow reserved them no height. (behindDoc ones went
  // first, above.) In-front floats defer to the queue: Word paints them
  // above every paragraph, not just the ones before their anchor.
  // The body pass collects EVERY drawing's hit box — behind-doc floats
  // painted by the earlier pass included (their boxes are just as clickable).
  let hitIndex = 0;
  for (const drawing of para.drawings ?? []) {
    const host = { para, index: hitIndex++ };
    if (drawing.behind) {
      if (ctx.hitBoxes) recordDrawingHit(drawing, x, y, ctx, ctx.hitBoxes, host);
      continue;
    }
    const paint = (): void => paintDrawing(tree, drawing, x, y, ctx, col, host);
    if (ctx.deferredDrawings)
      ctx.deferredDrawings.push({ z: drawing.zIndex ?? 0, layer: "body", paint });
    else paint();
  }
}

/** Formatting marks gray — the break rows' only paint (the character marks
 *  ride the text's own color, dimmed by opacity like Word's non-printing
 *  characters: slightly for the strokes, hard for the dense space dots). */
const MARK_COLOR = "#808080";
const MARK_OPACITY = 0.75;
const SPACE_DOT_OPACITY = 0.45;

/** Paragraph-wide state for the space-dot walk: the run texts concatenated,
 *  the walk cursor into that text and its drift flag — carried across the
 *  paragraph's lines. The walk itself is the layout's shared `lineSpaceGaps`
 *  (the caret map builds its character lattice from the same one). */
interface ParagraphMarkState {
  cursor: number;
  broken: boolean;
  fullText: string;
}

function paragraphMarkState(para: LaidOutParagraph): ParagraphMarkState {
  // The caret map's runTextOf — the same placeholder-marked text, so the two
  // lattices the shared gap walk builds stay in lockstep.
  let fullText = "";
  for (const inline of para.inline) {
    fullText += inline.kind === "text" ? inline.text : "￼";
  }
  return { cursor: 0, broken: false, fullText };
}

/** The line's dominant run — the largest text run by vert-align-scaled size,
 *  blank runs skipped; the tab leader's dots and the formatting marks both
 *  ride its size and color (the tab atom carries no style, the marks ride
 *  the text's own). `floorPx`/`fallback` are the caller's defaults — the ¶
 *  strut and ink — kept when no run beats them. */
function dominantRunOf(
  para: LaidOutParagraph,
  line: LaidOutLine,
  floorPx: number,
  fallback: string,
): { sizePx: number; color: string } {
  let sizePx = floorPx;
  let color = fallback;
  for (const other of line.items) {
    const src = para.inline[other.inlineIndex];
    if (src?.kind !== "text" || !(src.text ?? "").trim()) continue;
    // Raised/lowered runs count at their scaled size — a footnote reference
    // must not pull the leader dots up.
    const px = vertAlignedSizePx(src.style);
    if (px > sizePx) {
      sizePx = px;
      color = src.style.color ? `#${src.style.color}` : color;
    }
  }
  return { sizePx, color };
}

/** One line's formatting marks (Word's ¶ toggle): the bent arrow at every
 *  line end — a soft break or the paragraph's final line (the paragraph mark
 *  rides the project's arrow style, not Word's ¶) — while a section-end
 *  paragraph paints Word's "═══分节符(下一页)═══" double rule instead. Plus
 *  an arrow at each tab's start and a round dot centered in each space, all
 *  Leafer vectors (the glyphs behind ↵/→ vary by font fallback). Size follows
 *  the line's largest run (the tab leader's reference); a textless line falls
 *  back to the paragraph's mark strut (w:pPr/w:rPr/w:sz). Character marks
 *  ride the run color, dimmed by opacity (Word paints its non-printing
 *  characters secondary to the text). */
function paintLineMarks(
  tree: IGroup,
  para: LaidOutParagraph,
  line: LaidOutLine,
  lineX: number,
  lineY: number,
  pad: number,
  ctx: PaintContext,
  marks: ParagraphMarkState,
  rightX: number,
): void {
  // The line's dominant run style — the end-of-line marks ride it like the
  // tab leader; the color comes off that same run.
  const dominant = dominantRunOf(para, line, para.markSizePx ?? 0, "#000000");
  const sizePx = dominant.sizePx > 0 ? dominant.sizePx : 12;
  const color = dominant.color;

  // Space dots: Word centers a dim dot in each space. Pretext trims the
  // inter-word spaces out of the laid items — they live on as the gaps
  // between them — so the dots are placed from the source text: the shared
  // gap walk (the caret map's character lattice runs the same one) counts
  // the whitespace each item consumed, and the dots paint centered in the
  // gap between the two caret boundaries flanking the space (the previous
  // item's laid end and this item's x). Dot, caret and selection edges
  // therefore share one geometry on natural, justified and squeezed lines
  // alike, and a dot can never drift with the stretch. Radius clamps to the
  // gap's per-space share; an atom (tab/picture) between the items or a
  // line-leading trim paints nothing.
  const gaps = marks.broken ? null : lineSpaceGaps(line, marks.fullText, marks.cursor);
  if (!gaps || !gaps.matched) marks.broken = true;
  else {
    marks.cursor = gaps.next;
    let prevEndPx: number | null = null;
    let prevIndex = -2;
    for (const [itemIndex, item] of line.items.entries()) {
      if (item.kind !== "text") continue;
      const spaces = gaps.spaces[itemIndex]!;
      const src = para.inline[item.inlineIndex];
      if (spaces > 0 && prevEndPx != null && src?.kind === "text" && prevIndex === itemIndex - 1) {
        const span = item.xPx - prevEndPx;
        if (span >= 2) {
          const fill = src.style.color ? `#${src.style.color}` : color;
          for (let j = 0; j < spaces; j++) {
            const r = Math.min(sizePx * 0.075, (span / spaces) * 0.35);
            const cx = prevEndPx + (span * (j + 0.5)) / spaces;
            tree.add(
              new Ellipse({
                x: lineX + cx - r,
                y: lineY + pad + sizePx * 0.55 - r,
                width: r * 2,
                height: r * 2,
                fill,
                opacity: SPACE_DOT_OPACITY,
                hittable: false,
              }),
            );
          }
        }
      }
      prevEndPx = item.xPx + item.widthPx;
      prevIndex = itemIndex;
    }
  }
  // Tab arrows sit at the tab's start (Word draws the arrow leading the hop).
  for (const item of line.items) {
    if (item.kind !== "tab") continue;
    const w = sizePx * 0.9;
    const h = sizePx * 0.24;
    tree.add(
      new Path({
        x: lineX + item.xPx,
        y: lineY + pad + sizePx * 0.62,
        path: `M0 0 H${w} M${w - h} ${-h} L${w} 0 L${w - h} ${h}`,
        stroke: color,
        strokeWidth: Math.max(1, sizePx * 0.07),
        opacity: MARK_OPACITY,
        hittable: false,
      }),
    );
  }
  // Line-end marks: the bent arrow on every line end (a soft break or the
  // paragraph's end — a page-split tail only marks when the text really ends
  // there); a section-end paragraph paints the double rule instead. An empty
  // paragraph's lone strut line always marks.
  const endInline = para.inline[line.endInlineIndex];
  const paragraphEnd = para.inline.length === 0 || line.endInlineIndex >= para.inline.length - 1;
  if (!paragraphEnd && endInline?.kind !== "break") return;
  const last = line.items[line.items.length - 1];
  // An empty line paints no items — its ↵ mark rides the alignment's share of
  // the slack (center: half, right: all), like the caret stub in the editor.
  const emptySlack = line.maxWidthPx ?? 0;
  const emptyShift =
    para.align === "center" ? emptySlack / 2 : para.align === "right" ? emptySlack : 0;
  const markX = lineX + (last ? last.xPx + last.widthPx : emptyShift);
  if (paragraphEnd && para.sectionEnd) {
    const labels = ctx.marksLabels;
    const label =
      para.sectionEnd === "continuous"
        ? (labels?.sectionBreakContinuous ?? "Section Break (Continuous)")
        : para.sectionEnd === "evenPage"
          ? (labels?.sectionBreakEvenPage ?? "Section Break (Even Page)")
          : para.sectionEnd === "oddPage"
            ? (labels?.sectionBreakOddPage ?? "Section Break (Odd Page)")
            : (labels?.sectionBreak ?? "Section Break (Next Page)");
    paintSectionEndMark(tree, markX + sizePx * 0.25, rightX, lineY + pad + sizePx * 0.62, label);
    return;
  }
  // The bent-arrow mark (Word's ↵): a rod dropping into a foot that ends in a
  // leftward head — down, then left. ANCHORED BY ITS LEFTMOST POINT a small
  // gap past the last glyph, so the shape never leans back over the text at
  // any font size; it stays smaller than the text with the foot longer than
  // the rod. A vector, not the font's ↵ glyph (fallback faces vary). The foot
  // sits ON the text baseline (0.85s below the element top — the same
  // leaferBaselinePadPx anchor the glyphs paint at), like Word's paragraph
  // mark: an earlier 0.3s anchor left it floating 0.17em above the line.
  const rod = sizePx * 0.38;
  const arm = sizePx * 0.6;
  const head = sizePx * 0.18;
  tree.add(
    new Path({
      x: markX + sizePx * 0.25,
      y: lineY + pad + sizePx * (0.85 - 0.38),
      path: `M0 ${rod} H${arm} V0 ` + `M${head} ${rod - head} L0 ${rod} L${head} ${rod + head}`,
      stroke: color,
      strokeWidth: Math.max(1, sizePx * 0.06),
      opacity: MARK_OPACITY,
      hittable: false,
    }),
  );
}

/** Word's section-end mark row: "═══分节符(下一页)═══" — a double rule from
 *  the mark position to the column's right edge with the label centered in
 *  it, in the marks gray (the label rides the UI language via PaintContext). */
function paintSectionEndMark(
  tree: IGroup,
  x1: number,
  x2: number,
  midY: number,
  label: string,
): void {
  const px = 11;
  const cx = (x1 + x2) / 2;
  const gap = measureMarkText(label, px) / 2 + 8;
  for (const dy of [-1.5, 1.5]) {
    tree.add(
      new Line({
        points: [x1, midY + dy, cx - gap, midY + dy],
        stroke: MARK_COLOR,
        hittable: false,
      }),
    );
    tree.add(
      new Line({
        points: [cx + gap, midY + dy, x2, midY + dy],
        stroke: MARK_COLOR,
        hittable: false,
      }),
    );
  }
  tree.add(
    new Text({
      x: cx - gap,
      y: midY - px * 0.72,
      width: gap * 2,
      height: px * 1.5,
      text: label,
      fill: MARK_COLOR,
      fontSize: px,
      textAlign: "center",
      lineHeight: px * 1.5,
      hittable: false,
    }),
  );
}

/** The painted width of a marks label (the break rows center their text) —
 *  the same hidden canvas the dot placement measures with. */
function measureMarkText(text: string, px: number): number {
  markCtx ??=
    typeof document === "undefined" ? null : document.createElement("canvas").getContext("2d");
  if (!markCtx) return text.length * px;
  markCtx.font = `${px}px sans-serif`;
  return markCtx.measureText(text).width;
}

/** Word's page-break row: a dotted rule across the column with the label
 *  ("·····分页符·····") centered in it, in the marks gray. Painted only
 *  while marks are visible — the row's height is the layout's, always. */
export function paintBreakRow(
  tree: IGroup,
  block: { heightPx: number },
  x: number,
  y: number,
  width: number,
  ctx: PaintContext,
): void {
  const label = ctx.marksLabels?.pageBreak ?? "Page Break";
  const px = 11;
  const midY = y + block.heightPx / 2;
  const cx = x + width / 2;
  const gap = measureMarkText(label, px) / 2 + 8;
  tree.add(
    new Line({
      points: [x, midY, cx - gap, midY],
      stroke: MARK_COLOR,
      dashPattern: [1, 3],
      hittable: false,
    }),
  );
  tree.add(
    new Line({
      points: [cx + gap, midY, x + width, midY],
      stroke: MARK_COLOR,
      dashPattern: [1, 3],
      hittable: false,
    }),
  );
  tree.add(
    new Text({
      x: cx - gap,
      y: midY - px * 0.72,
      width: gap * 2,
      height: px * 1.5,
      text: label,
      fill: MARK_COLOR,
      fontSize: px,
      textAlign: "center",
      lineHeight: px * 1.5,
      hittable: false,
    }),
  );
}

/** A hidden 2d context for marks-label measurement — paint runs in the
 *  browser (Leafer), so the canvas is always available here. */
let markCtx: CanvasRenderingContext2D | null | undefined;

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

/** The font family a text slice paints in: the measurement side's slot pick
 *  (cssFontOf builds `, serif` on top of it — layout, caret map and paint
 *  must resolve empty slots to the same face or glyph advances drift from
 *  the boundaries the caret map computed). */
function familyOf(style: LayoutTextStyle, text: string): string {
  return familyOfSlot(style.family, isCjkCodeUnit(text, 0)) || "serif";
}
