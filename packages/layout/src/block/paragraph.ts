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
//     - picture-sized line → same ceil (the box spans whole rows; the
//       painter half-leads it — gridPadOf)
//     - Latin line → max(natural, pitch)
//  3. Otherwise → natural.
//
// The ¶-mark strut: an empty paragraph's line (and a picture row's minimum)
// is the paragraph-mark line height — spacing.line first, then the ¶-mark
// size (an absolute height), then the default run's natural metric. The
// empty line takes NO grid pitch (verified vs Word).

import {
  type LayoutBlockContext,
  type LayoutFloatZone,
  type LayoutInline,
  type LayoutLineHeight,
  type LayoutParagraph,
  wrapEffectsOf,
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

  // The anchor paragraph wraps beside its own square floats: the drawings'
  // offsets are paragraph-relative by definition, so their zones feed the
  // packer in paragraph-relative Y (the flow's — or the table cell's —
  // absolute zones cover every later paragraph). The box grows by the
  // anchor's wrap distances first (distL/R/T/B), matching the flow's zone
  // padding. Full-column boxes and topAndBottom clears are bands — the
  // packer cannot skip a mid-paragraph band, so they stay flow-only.
  // Table cells qualify too (Word's layoutInCell): a cell-anchored float
  // wraps the cell's text, and inside a cell `column` IS the cell's column.
  const selfZones: LayoutFloatZone[] = wrapEffectsOf(para.drawings, 0, width, inTable).zones;

  const packed = packLines(para.inline, {
    measurer,
    width: usable,
    firstLineIndentPx: para.indent?.firstLinePx,
    tabStops: para.tabStops,
    strutPx,
    lineHeight: ({ naturalPx, hasCjk, hasPicture }) =>
      spec
        ? resolveLine(spec, naturalPx, pitch, hasCjk, inTable, hasPicture)
        : snapLine(naturalPx, hasCjk, pitch, inTable, hasPicture),
    startY: ctx?.startY,
    // Absolute zones come from whoever stacks the blocks: the flow passes the
    // page's float zones, the table cell stacker accumulates the cell's own
    // (a cell's width is its column — page floats never reach it because the
    // cellCtx starts empty, not because the paragraph drops them here).
    floatZones: ctx?.floatZones,
    selfZones: selfZones.length > 0 ? selfZones : undefined,
  });

  // An empty paragraph still occupies one strut line (the ¶ glyph's).
  if (packed.length === 0) {
    packed.push({
      items: [],
      endInlineIndex: 0,
      maxWidthPx: usable,
      heightPx: strutPx,
      naturalPx: strutNatural,
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
      const contentEnd =
        last.xPx +
        last.widthPx -
        Math.max(trailingHang(items, para.inline, measurer), line.hangPx ?? 0);
      const slack = line.maxWidthPx - contentEnd;
      if (slack <= 0) continue;
      const shift = para.align === "center" ? slack / 2 : slack;
      for (const it of items) it.xPx += shift;
    }
  }

  const lines: LaidOutLine[] = [];
  let heightPx = 0;
  // Grid lattice: only the body flow (ctx.onGrid) centers grid-height lines;
  // exact/atLeast spacing overrides the grid's height and with it the centering.
  const gridLine =
    pitch > 0 && ctx?.onGrid === true && spec?.rule !== "exact" && spec?.rule !== "atLeast";
  for (let i = 0; i < packed.length; i++) {
    const line = packed[i];
    lines.push({
      yPx: heightPx,
      heightPx: line.heightPx,
      naturalPx: line.naturalPx,
      textEmPx: line.textEmPx,
      grid: gridLine || undefined,
      pictureFloored: line.pictureFloored,
      firstLineIndentPx: i === 0 ? para.indent?.firstLinePx : undefined,
      endInlineIndex: line.endInlineIndex,
      items: line.items,
      // Carried on every line (not just justified ones) — the wrap width is
      // the selection highlight's right edge (Word highlights to the wrap
      // edge, not the last glyph) and hit-testing's line-end boundary.
      maxWidthPx: line.maxWidthPx,
      justifyGapPx: justifyGapPx[i],
      hangPx: line.hangPx,
      xOffsetPx: line.xOffsetPx,
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
    shadingFill: para.shadingFill,
    indent: para.indent,
    drawings: para.drawings,
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
 *  item's content (trailing whitespace AND an overflow-punct hang — both
 *  reach past the right edge, both excluded) spread over justify units —
 *  CJK items stretch per inter-grapheme gap (Word's CJK justification),
 *  Latin items per word gap (spaces absorb the slack; letter-spreading
 *  English is the letter-mode tell-tale). Re-spaces every item's x in place
 *  and returns the per-unit stretch (0 when there is nothing to spread). */
function justifyLine(
  line: PackedLine,
  inline: readonly LayoutInline[],
  measurer: TextMeasurer,
): number {
  const items = line.items;
  const last = items[items.length - 1];
  const hang = Math.max(trailingHang(items, inline, measurer), line.hangPx ?? 0);
  const itemUnits = items.map((it) => {
    if (it.kind !== "text") return 1;
    if (CJK_ITEM.test(it.text)) {
      let points = 0;
      for (const _ of it.text) points++;
      return points - 1;
    }
    const indices = leaferWordIndices(it.text);
    return indices[indices.length - 1] ?? 0;
  });
  let units = 0;
  for (const u of itemUnits) units += Math.max(u, 0);
  const delta =
    units > 0 ? Math.max(0, line.maxWidthPx - (last.xPx + last.widthPx - hang)) / units : 0;
  if (delta === 0) return 0;
  let before = 0;
  items.forEach((it, i) => {
    it.xPx += delta * before;
    before += itemUnits[i]!;
  });
  return delta;
}

/** Items stretch per grapheme when they hold any CJK glyph (the painter's
 *  "both-letter" trigger — painter and caret map consume this same test). */
const CJK_ITEM = /[一-鿿぀-ヿ가-힯]/;

/** Whether a text item justifies per grapheme (any CJK glyph) rather than
 *  per word gap — Leafer's both-letter vs both-justify choice, shared by
 *  the painter's textAlign and the caret map's boundary distribution. */
export function justifyPerGrapheme(text: string): boolean {
  return CJK_ITEM.test(text);
}

/** Leafer's word split (its justify denominator): a space or one of its
 *  break chars, or a CJK glyph, stands alone as one word; other runs
 *  coalesce. Maps each grapheme to its word index — the count is the last
 *  index + 1, and a grapheme's justify shift is its index × the per-gap
 *  stretch (the painter's "both-justify" Text applies exactly this). */
export function leaferWordIndices(text: string): number[] {
  const indices: number[] = [];
  let word = -1;
  let inRun = false;
  for (const ch of text) {
    if (ch === " " || LEAFER_BREAK_CHARS.has(ch) || CJK_ITEM.test(ch)) {
      word++;
      inRun = false;
    } else {
      if (!inRun) word++;
      inRun = true;
    }
    indices.push(word);
  }
  return indices;
}

const LEAFER_BREAK_CHARS = new Set(["-", "—", "／", "～", "｜", "┆", "·"]);

/** A spacing.line spec against a line's natural height. On a grid, a body CJK
 *  line's spec'd height never falls below the line's natural height and snaps
 *  up to whole rows; a cell line (w:adjustLineHeightInTable, on in every
 *  CJK-Word document) adds grid pitch only as far as the line's own natural
 *  height demands — the multiple's demand is factor × pitch, and the natural
 *  height snaps up to whole rows *before* the comparison. Corpus-verified
 *  (honor table, 340-twip grid): the 14pt header cell (natural 24.2px >
 *  pitch 22.7px) takes 2 rows (45.3px) while every 12pt/10.5pt row stays at
 *  1.5 × pitch (34px); shrinking the header run to 12pt moves its bottom
 *  border from y215 to y204 (= 1.5 × pitch), confirming natural — not factor
 *  — drives the snap. */
function resolveLine(
  spec: LayoutLineHeight,
  naturalPx: number,
  pitch: number,
  hasCjk = false,
  inTable = false,
  hasPicture = false,
): number {
  if (spec.rule === "exact") return spec.px;
  if (spec.rule === "atLeast") return Math.max(naturalPx, spec.px);
  // multiple: 240ths of a single line — the grid pitch when defined, else the
  // font natural (verified vs Word). On a grid, a body CJK line's spec'd
  // height never falls below the line's natural height and snaps up to whole
  // rows: a face taller than its grid rows (Microsoft YaHei runs ~1.7em while
  // a 340-twip pitch is ~1.2em of a 14pt line) takes the rows it needs. Latin
  // lines are exempt from the lattice — but a picture-sized line joins it
  // (its box spans whole rows and half-leads; pixel-verified against the
  // reference renders).
  if (pitch > 0 && (hasCjk || hasPicture)) {
    if (inTable) {
      // A cell line's natural height still snaps up to whole rows before the
      // comparison against the multiple's demand (see resolveLine doc).
      return Math.max(spec.factor * pitch, Math.ceil(naturalPx / pitch) * pitch);
    }
    const specH = spec.factor * pitch;
    return Math.ceil(Math.max(specH, naturalPx) / pitch) * pitch;
  }
  const single = pitch > 0 ? pitch : naturalPx;
  return spec.factor * single;
}

/** No spacing.line: snap to the document grid. */
function snapLine(
  naturalPx: number,
  hasCjk: boolean,
  pitch: number,
  inTable: boolean,
  hasPicture = false,
): number {
  if (pitch <= 0) return naturalPx;
  // A cell line's natural height snaps up to whole rows — the "add grid
  // pitch" compat raises it as far as the line demands, never just one pitch.
  // A picture-sized body line snaps the same way: the box spans whole rows
  // and gridPadOf half-leads the picture inside them.
  if (inTable || hasCjk || hasPicture) return Math.ceil(naturalPx / pitch) * pitch;
  return Math.max(naturalPx, pitch);
}
