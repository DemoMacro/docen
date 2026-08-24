// Paragraph layout — OOXML line-height semantics over the packer's output.
// Ported from the editor's measure.ts (its full feature list is P1's
// acceptance bar), with the PmNode/style-cascade inputs replaced by the
// already-resolved LayoutParagraph projection.
//
// Line-height model (ECMA-376, verified against Word in the measure.ts era):
//  1. spacing.line applies to every paragraph, table cells included: exact →
//     fixed; atLeast → max(natural, spec) — the DOM route pinned atLeast to
//     its spec value as a CSS approximation; the engine does the true max;
//     multiple → factor × single line (the docGrid pitch when a grid is
//     defined, else the font natural).
//  2. No spacing.line + snapToGrid on + a grid pitch:
//     - table cell → max(natural, pitch) (the row's trHeight governs)
//     - CJK line → ceil(natural / pitch) × pitch (chars snap to the grid)
//     - Latin line → max(natural, pitch)
//  3. Otherwise → natural.
//
// The ¶-mark strut: an empty paragraph's line (and a picture row's minimum)
// is the paragraph-mark line height — spacing.line first, then the ¶-mark
// size (an absolute height), then the default run's natural metric. The
// empty line takes NO grid pitch (verified vs Word).

import type {
  LayoutBlockContext,
  LayoutInline,
  LayoutLineHeight,
  LayoutParagraph,
} from "../layout-doc";
import type { LaidOutLine, LaidOutLineItem, LaidOutParagraph } from "../layout-result";
import { packLines, type PackedLine } from "../text/line-break";
import type { TextMeasurer } from "../text/measure";

/** Lay out a paragraph at `width` (its container's content width; indents
 *  shrink the usable width inside). */
export function layoutParagraph(
  para: LayoutParagraph,
  width: number,
  ctx: LayoutBlockContext | undefined,
  measurer: TextMeasurer,
): LaidOutParagraph {
  const spec = para.spacing?.lineHeight;
  const pitch = para.snapToGrid === false ? 0 : (ctx?.linePitchPx ?? 0);
  const inTable = ctx?.inTable ?? false;

  // The ¶-mark strut line: spacing wins, then the mark size, then the default
  // run's natural metric (no grid pitch — see the module doc).
  const strutNatural = para.defaultTextStyle ? measurer.naturalOf(para.defaultTextStyle) : 0;
  let strutPx: number;
  if (spec) {
    strutPx = resolveLine(spec, strutNatural, pitch);
  } else {
    strutPx = para.markSizePx ?? strutNatural;
  }

  const usable = Math.max(0, width - (para.indent?.leftPx ?? 0) - (para.indent?.rightPx ?? 0));

  const packed = packLines(para.inline, {
    measurer,
    width: usable,
    firstLineIndentPx: para.indent?.firstLinePx,
    tabStops: para.tabStops,
    strutPx,
    lineHeight: ({ naturalPx, hasCjk }) =>
      spec ? resolveLine(spec, naturalPx, pitch) : snapLine(naturalPx, hasCjk, pitch, inTable),
    startY: ctx?.startY,
    // A cell's width is its column, not the page flow — floats never bend it.
    floatZones: inTable ? undefined : ctx?.floatZones,
  });

  // An empty paragraph still occupies one strut line (the ¶ glyph's).
  if (packed.length === 0) {
    packed.push({
      items: [],
      endInlineIndex: 0,
      maxWidthPx: usable,
      heightPx: strutPx,
    });
  }

  // w:jc — horizontal alignment, one pass over the packed lines (the layout
  // owns the geometry; the painter places what it is given):
  //  - both/distribute stretch inter-character gaps to the line's content
  //    width. both skips the paragraph's last line and hard-break lines (they
  //    are that logical line's natural end); distribute stretches every line.
  //  - center/right shift the whole line's items by the slack after the
  //    content (trailing whitespace hangs and never counts).
  const justifyGapPx: (number | undefined)[] = packed.map(() => undefined);
  const stretchAll = para.align === "distribute";
  if (para.align === "both" || stretchAll) {
    for (let i = 0; i < packed.length; i++) {
      const line = packed[i];
      if (line.items.length === 0) continue;
      if (!stretchAll && i === packed.length - 1) continue;
      // A hard break ends this line — it is that logical line's last line.
      if (para.inline[line.endInlineIndex]?.kind === "break") continue;
      justifyGapPx[i] = justifyLine(line, para.inline, measurer);
    }
  } else if (para.align === "center" || para.align === "right") {
    for (const line of packed) {
      const items = line.items;
      if (items.length === 0) continue;
      const last = items[items.length - 1];
      const contentEnd = last.xPx + last.widthPx - trailingHang(items, para.inline, measurer);
      const slack = line.maxWidthPx - contentEnd;
      if (slack <= 0) continue;
      const shift = para.align === "center" ? slack / 2 : slack;
      for (const it of items) it.xPx += shift;
    }
  }

  const lines: LaidOutLine[] = [];
  let heightPx = 0;
  for (let i = 0; i < packed.length; i++) {
    const line = packed[i];
    lines.push({
      yPx: heightPx,
      heightPx: line.heightPx,
      endInlineIndex: line.endInlineIndex,
      items: line.items,
      maxWidthPx: justifyGapPx[i] !== undefined ? line.maxWidthPx : undefined,
      justifyGapPx: justifyGapPx[i],
    });
    heightPx += line.heightPx;
  }

  return {
    kind: "paragraph",
    heightPx,
    beforePx: para.spacing?.beforePx ?? 0,
    afterPx: para.spacing?.afterPx ?? 0,
    lines,
    inline: para.inline,
    keepLines: para.keepLines,
    keepNext: para.keepNext,
    widowControl: para.widowControl,
    borders: para.borders,
    indent: para.indent,
  };
}

/** The width of the last text item's trailing whitespace — it hangs past the
 *  line's right edge and never counts toward justification or alignment. */
function trailingHang(
  items: readonly LaidOutLineItem[],
  inline: readonly LayoutInline[],
  measurer: TextMeasurer,
): number {
  const last = items[items.length - 1];
  if (last.kind !== "text") return 0;
  const src = inline[last.inlineIndex];
  if (src?.kind !== "text") return 0;
  const trail = /\s+$/.exec(last.text)?.[0];
  return trail ? measurer.analyze(trail, src.style).naturalPx : 0;
}

/** Stretch one wrapped line to its packed width: the slack after the last
 *  item's content (trailing whitespace hangs past the right edge — excluded)
 *  spread evenly over the inter-character gaps. Re-spaces every item's x in
 *  place and returns the per-gap stretch (0 when there is nothing to
 *  spread). */
function justifyLine(
  line: PackedLine,
  inline: readonly LayoutInline[],
  measurer: TextMeasurer,
): number {
  const items = line.items;
  const last = items[items.length - 1];
  const hang = trailingHang(items, inline, measurer);
  let graphemes = 0;
  for (const it of items) graphemes += it.kind === "text" ? [...it.text].length : 1;
  const delta = Math.max(
    0,
    (line.maxWidthPx - (last.xPx + last.widthPx - hang)) / Math.max(1, graphemes - 1),
  );
  if (delta === 0) return 0;
  let before = 0;
  for (const it of items) {
    it.xPx += delta * before;
    before += it.kind === "text" ? [...it.text].length : 1;
  }
  return delta;
}

/** A spacing.line spec against a line's natural height. */
function resolveLine(spec: LayoutLineHeight, naturalPx: number, pitch: number): number {
  if (spec.rule === "exact") return spec.px;
  if (spec.rule === "atLeast") return Math.max(naturalPx, spec.px);
  // multiple: 240ths of a single line — the grid pitch when defined, else the
  // font natural (verified vs Word).
  const single = pitch > 0 ? pitch : naturalPx;
  return spec.factor * single;
}

/** No spacing.line: snap to the document grid. */
function snapLine(naturalPx: number, hasCjk: boolean, pitch: number, inTable: boolean): number {
  if (pitch <= 0) return naturalPx;
  if (inTable) return Math.max(naturalPx, pitch);
  if (hasCjk) return Math.ceil(naturalPx / pitch) * pitch;
  return Math.max(naturalPx, pitch);
}
