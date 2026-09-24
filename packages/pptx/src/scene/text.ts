// Shape text projection: a bodyPr's paragraphs as the layout paragraphs the
// painter stacks, with Word-compatible run/paragraph defaults.

import { solidFillOf } from "@docen/core/geometry";
import {
  EMU_PER_PX,
  formatNumber,
  ptToPx,
  type LayoutInline,
  type LayoutParagraph,
  type LayoutTextStyle,
} from "@docen/layout";
import type {
  BulletOptions,
  RunFont,
  TextBodyOptions,
  TextCharacterPropertiesOptions,
  TextFont,
  TextParagraphPropertiesOptions,
} from "@office-open/core/drawing";

// TextAlignment → the layout's align tokens: justify's spellings land on
// "both", the distributed pair on "distribute".
const ALIGN_MAP = {
  left: "left",
  center: "center",
  right: "right",
  justify: "both",
  lowJustification: "both",
  distribute: "distribute",
  thaiDistributed: "distribute",
} as const;

// PowerPoint's default run size (a:sz 1800) and face (the theme minor font —
// master/lstStyle defaults are a batch gap; every run carries explicit props
// in common decks).
const DEFAULT_FAMILY = "Calibri";
const DEFAULT_SIZE_PX = ptToPx(18);

// A fresh bullet's indents (a:pPr @marL 342900 / @indent −342900 — 0.375"
// hanging); each @lvl level deepens both by 0.5". These also apply when the
// paragraph carries a bullet but no explicit marL/indent.
const BULLET_MARL_EMU = 342900;
const LEVEL_STEP_EMU = 457200;

// a:buAutoNum scheme → the layout formatter's (w:numFmt) word list + suffix.
// Unlisted schemes keep PowerPoint's plain "N." look through the default.
const AUTONUM_FORMATS: ReadonlyMap<string, [format: string, suffix: string]> = new Map([
  ["arabicPeriod", ["decimal", "."]],
  ["arabicPlain", ["decimal", ""]],
  ["arabicParenR", ["decimal", ")"]],
  ["arabicParenBoth", ["decimal", ")"]],
  ["alphaLcPeriod", ["lowerLetter", "."]],
  ["alphaUcPeriod", ["upperLetter", "."]],
  ["alphaLcParenR", ["lowerLetter", ")"]],
  ["alphaUcParenR", ["upperLetter", ")"]],
  ["alphaLcParenBoth", ["lowerLetter", ")"]],
  ["alphaUcParenBoth", ["upperLetter", ")"]],
  ["romanLcPeriod", ["lowerRoman", "."]],
  ["romanUcPeriod", ["upperRoman", "."]],
  ["romanLcParenR", ["lowerRoman", ")"]],
  ["romanUcParenBoth", ["upperRoman", ")"]],
  ["ea1ChsPeriod", ["chineseCounting", "."]],
  ["ea1ChsPlain", ["chineseCounting", ""]],
]);

const autonumMarkerOf = (format: string | undefined, n: number): string => {
  const [fmt, suffix] = AUTONUM_FORMATS.get(format ?? "arabicPeriod") ?? ["decimal", "."];
  return formatNumber(fmt, n) + suffix;
};

/** A body's paragraphs as projected blocks; the paragraph shape (not the
 *  LayoutBlock union) so consumers read align/spacing directly. */
export function textBlocks(body: TextBodyOptions): LayoutParagraph[] {
  const paragraphs = body.paragraphs ?? (body.text != null ? [body.text] : []);
  // Live autonumbering counters, keyed by scheme; a non-list paragraph
  // breaks every run (PowerPoint restarts the series after plain text).
  const counters = new Map<string, number>();
  return paragraphs.map((p) => {
    const para = typeof p === "string" ? { text: p } : p;
    const props = para.properties;
    const inline: LayoutInline[] = [];
    const kids = para.children ?? (para.text != null ? [para.text] : []);
    for (const kid of kids) {
      if (typeof kid === "string") {
        inline.push({ kind: "text", text: kid, style: runStyle(undefined) });
      } else if ("break" in kid) {
        inline.push({ kind: "break" });
      } else if ("type" in kid) {
        // a:fld paints its cached display text; live evaluation (slidenum,
        // datetime) lands with a field engine — batch gap.
        inline.push({ kind: "text", text: kid.text ?? "", style: runStyle(kid.properties) });
      } else {
        inline.push({ kind: "text", text: kid.text ?? "", style: runStyle(kid) });
      }
    }
    // The bullet/number marker rides the inline stream the docx projection's
    // numbering does — a synthetic glyph + tab hop to the body-text start,
    // with the hanging indent putting the marker left of that start. Its
    // style follows the paragraph's first text run.
    const firstRun = inline.find(
      (part): part is Extract<LayoutInline, { kind: "text" }> => part.kind === "text",
    );
    const marker = markerInline(props, firstRun?.style ?? runStyle(undefined), counters);
    if (marker) inline.unshift(...marker);
    const level = (props?.indentLevel ?? 0) * LEVEL_STEP_EMU;
    const listEmu = marker ? BULLET_MARL_EMU + level : null;
    const leftEmu = props?.marginIndent ?? listEmu;
    const firstEmu = props?.indent ?? (listEmu != null ? -listEmu : undefined);
    return {
      kind: "paragraph",
      inline,
      ...(props?.alignment ? { align: ALIGN_MAP[props.alignment] } : {}),
      ...(props?.spaceBefore != null ||
      props?.spaceAfter != null ||
      props?.lineSpacingPoints != null ||
      props?.lineSpacingPercent != null
        ? {
            spacing: {
              beforePx: ptToPx(props.spaceBefore ?? 0),
              afterPx: ptToPx(props.spaceAfter ?? 0),
              ...(props.lineSpacingPoints != null
                ? { lineHeight: { rule: "exact", px: ptToPx(props.lineSpacingPoints) } }
                : props.lineSpacingPercent != null
                  ? { lineHeight: { rule: "multiple", factor: props.lineSpacingPercent / 100 } }
                  : {}),
            },
          }
        : {}),
      ...(leftEmu != null || firstEmu != null
        ? {
            indent: {
              ...(leftEmu != null ? { leftPx: leftEmu / EMU_PER_PX } : {}),
              ...(firstEmu != null ? { firstLinePx: firstEmu / EMU_PER_PX } : {}),
            },
          }
        : {}),
      // An empty paragraph renders as a blank line — the strut keeps its
      // height (the default style supplies the measuring font).
      ...(inline.length === 0 ? { defaultTextStyle: runStyle(undefined) } : {}),
    };
  });
}

/** The bullet marker as a synthetic inline pair (glyph + tab hop), or null
 *  when the paragraph carries no live bullet. Non-list paragraphs reset the
 *  autonumber counters. The glyph takes its color/size from the shared
 *  bullet style where present, else from the paragraph's first run. */
function markerInline(
  props: TextParagraphPropertiesOptions | undefined,
  runStyle: LayoutTextStyle,
  counters: Map<string, number>,
): LayoutInline[] | null {
  const bullet: BulletOptions | undefined = props?.bullet;
  if (!bullet) {
    counters.clear();
    return null;
  }
  if (bullet.type === "none") return null;
  let text: string;
  if (bullet.type === "autoNum") {
    const key = bullet.format ?? "arabicPeriod";
    const n = (counters.get(key) ?? (bullet.startAt != null ? bullet.startAt - 1 : 0)) + 1;
    counters.set(key, n);
    text = autonumMarkerOf(bullet.format, n);
  } else if (bullet.type === "char") {
    text = bullet.char ?? "•";
  } else {
    // Picture bullets need the image relationship the projection can't name
    // (r:embed) — they degrade to no marker until a media-aware pass.
    return null;
  }
  // buFont/buClr/buSz win over the run when explicit; the "follows text"
  // toggles and an unset dimension keep the run's value.
  const style: LayoutTextStyle = {
    ...runStyle,
    ...(bullet.font ? { family: bullet.font } : {}),
    ...(bullet.sizePoints != null
      ? { sizePx: ptToPx(bullet.sizePoints) }
      : bullet.size != null
        ? { sizePx: (runStyle.sizePx * bullet.size) / 100 }
        : {}),
    ...(typeof bullet.color === "string" ? { color: bullet.color } : {}),
  };
  return [
    { kind: "text", text, style, synthetic: true },
    { kind: "tab", toPx: 0 },
  ];
}

function runStyle(rp: TextCharacterPropertiesOptions | undefined): LayoutTextStyle {
  const family = familyOf(rp?.font);
  // FillOptions (parse emits {type:"solid", color}) — colorOf alone only
  // reads the bare-string/flat-color shapes and would drop every parsed run.
  const color = solidFillOf(rp?.fill);
  return {
    family: family ?? DEFAULT_FAMILY,
    sizePx: rp?.size != null ? ptToPx(rp.size) : DEFAULT_SIZE_PX,
    ...(rp?.bold ? { bold: true } : {}),
    ...(rp?.italic ? { italic: true } : {}),
    ...(rp?.underline && rp.underline !== "none" ? { underline: true } : {}),
    ...(rp?.strike === "singleStrike" || rp?.strike === "doubleStrike"
      ? { strikethrough: true }
      : {}),
    ...(color ? { color } : {}),
  };
}

/** A bare string sets latin + ea to the same face; the object form reads
 *  latin first, then eastAsia (each slot is a face string or a full font). */
function familyOf(font: RunFont | undefined): string | undefined {
  if (typeof font === "string") return font || undefined;
  const face = (f: TextFont | undefined): string | undefined =>
    typeof f === "string" ? f || undefined : f?.typeface;
  return face(font?.latin) ?? face(font?.eastAsia);
}
