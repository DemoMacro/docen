import {
  familyOfSlot,
  gridPadOf,
  isCjkCodeUnit,
  justifiedIntervals,
  justifyPerGrapheme,
  lineOriginXPx,
  vertAlignedSizePx,
  vertAlignBaselineShiftPx,
  type LaidOutLineItem,
  type LaidOutParagraph,
  type LayoutInline,
  type LayoutParagraphBorderEdge,
  type LayoutTextStyle,
} from "@docen/layout";
import { Box, Image as LeaferImage, Line, Rect, Text, type IGroup } from "leafer-ui";

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

export function paintParagraph(
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
    // Paragraph shading (w:shd) fills the block box beneath everything else
    // the paragraph paints — behind-doc floats included, matching Word.
    if (para.shadingFill) {
      const right = ctx.flow.contentLeftPx + ctx.flow.contentWidthPx;
      tree.add(
        new Rect({
          x,
          y,
          width: Math.max(0, right - x),
          height: para.heightPx,
          fill: `#${para.shadingFill}`,
        }),
      );
    }
    let index = 0;
    for (const drawing of para.drawings ?? []) {
      if (drawing.behind) paintDrawing(tree, drawing, x, y, ctx, col, { para, index: index++ });
      else index++;
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
    const boxRight = ctx.flow.contentLeftPx + ctx.flow.contentWidthPx;
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
    if (ctx.deferredDrawings) ctx.deferredDrawings.push(paint);
    else paint();
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

/** The font family a text slice paints in: the measurement side's slot pick
 *  (cssFontOf builds `, serif` on top of it — layout, caret map and paint
 *  must resolve empty slots to the same face or glyph advances drift from
 *  the boundaries the caret map computed). */
function familyOf(style: LayoutTextStyle, text: string): string {
  return familyOfSlot(style.family, isCjkCodeUnit(text, 0)) || "serif";
}
