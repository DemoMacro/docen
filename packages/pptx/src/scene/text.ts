// Shape text projection: a bodyPr's paragraphs as the layout paragraphs the
// painter stacks, with Word-compatible run/paragraph defaults.

import {
  EMU_PER_PX,
  formatNumber,
  ptToPx,
  type FontMetrics,
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
import type { ColorMappingOptions } from "@office-open/core/theme";

import { shapeFillOf } from "./shape-fill";
import type { TableStyleRegions, ThemeColors } from "./table-style";

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

/** Slide-level values needed to evaluate a field when its cached text is
 *  stale; all values are optional so standalone bodies keep using their cache. */
export interface TextFieldContext {
  slideNumber?: number;
  slideCount?: number;
  now?: Date;
  /** Theme color slots (accent1…); table styles resolve schemeClr fills. */
  themeColors?: ThemeColors;
  /** The font scheme's faces; +mn/+mj run-font tokens resolve through it. */
  themeFonts?: ThemeFonts;
  /** The master's clrMap; shape fills resolve short scheme tokens through it. */
  colorMapping?: ColorMappingOptions;
  /** The presentation's custom table styles (p:tblStyleLst), keyed by
   * upper-cased style GUID. */
  tableStyles?: Record<string, TableStyleRegions>;
  /** The layout metrics table auto-height uses before painting. */
  metrics?: FontMetrics;
}

/** A body's paragraphs as projected blocks; the paragraph shape (not the
 *  LayoutBlock union) so consumers read align/spacing directly. */
export function textBlocks(
  body: TextBodyOptions,
  context: TextFieldContext = {},
): LayoutParagraph[] {
  const paragraphs = body.paragraphs ?? (body.text != null ? [body.text] : []);
  // Live autonumbering counters, keyed by scheme; a non-list paragraph
  // breaks every run (PowerPoint restarts the series after plain text).
  const counters = new Map<string, number>();
  return paragraphs.map((p) => {
    const para = typeof p === "string" ? { text: p } : p;
    const explicit = para.properties;
    const props = mergeParagraphProperties(
      explicit,
      inheritedParagraphProperties(body, explicit?.indentLevel ?? 0),
    );
    const inline: LayoutInline[] = [];
    const kids = para.children ?? (para.text != null ? [para.text] : []);
    for (const kid of kids) {
      if (typeof kid === "string") {
        inline.push({ kind: "text", text: kid, style: runStyle(undefined, context, props) });
      } else if ("break" in kid) {
        inline.push({ kind: "break" });
      } else if ("type" in kid) {
        inline.push({
          kind: "text",
          text: fieldTextOf(kid.type, kid.text, context),
          style: runStyle(kid.properties, context, props),
        });
      } else {
        inline.push({ kind: "text", text: kid.text ?? "", style: runStyle(kid, context, props) });
      }
    }
    // The bullet/number marker rides the inline stream the docx projection's
    // numbering does — a synthetic glyph + tab hop to the body-text start,
    // with the hanging indent putting the marker left of that start. Its
    // style follows the paragraph's first text run.
    const firstRun = inline.find(
      (part): part is Extract<LayoutInline, { kind: "text" }> => part.kind === "text",
    );
    const marker = markerInline(
      props,
      firstRun?.style ?? runStyle(undefined, context, props),
      counters,
    );
    if (marker) inline.unshift(...marker);
    const level = (props.indentLevel ?? 0) * LEVEL_STEP_EMU;
    const referenceSizePx = Math.max(
      0,
      ...inline.flatMap((part) =>
        part.kind === "text" && part.style.sizePx != null ? [part.style.sizePx] : [],
      ),
      firstRun?.style.sizePx ?? runStyle(undefined, context, props).sizePx ?? DEFAULT_SIZE_PX,
    );
    const lineSpacingPercent = props.lineSpacingPercent ?? 100;
    const listEmu = marker ? BULLET_MARL_EMU + level : null;
    const leftEmu = props.marginIndent ?? listEmu;
    const firstEmu = props.indent ?? (listEmu != null ? -listEmu : undefined);
    return {
      kind: "paragraph",
      inline,
      ...(props.alignment ? { align: ALIGN_MAP[props.alignment] } : {}),
      ...(props.spaceBefore != null ||
      props.spaceAfter != null ||
      props.lineSpacingPoints != null ||
      props.lineSpacingPercent != null ||
      inline.length > 0
        ? {
            spacing: {
              beforePx: ptToPx(props.spaceBefore ?? 0),
              afterPx: ptToPx(props.spaceAfter ?? 0),
              ...(props.lineSpacingPoints != null
                ? { lineHeight: { rule: "exact", px: ptToPx(props.lineSpacingPoints) } }
                : {
                    lineHeight: {
                      rule: "exact",
                      px: (referenceSizePx * 1.2 * lineSpacingPercent) / 100,
                    },
                  }),
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
      ...(inline.length === 0 ? { defaultTextStyle: runStyle(undefined, context, props) } : {}),
    };
  });
}

function inheritedParagraphProperties(
  body: TextBodyOptions,
  level: number,
): TextParagraphPropertiesOptions | undefined {
  return {
    ...body.listStyle?.defaultParagraph,
    ...body.listStyle?.levels?.[level ?? 0],
  };
}

function mergeParagraphProperties(
  explicit: TextParagraphPropertiesOptions | undefined,
  inherited: TextParagraphPropertiesOptions | undefined,
): TextParagraphPropertiesOptions {
  if (!inherited) return explicit ?? {};
  return {
    ...inherited,
    ...explicit,
    defaultRunProperties: {
      ...inherited.defaultRunProperties,
      ...explicit?.defaultRunProperties,
    },
  };
}

/** Evaluate the two footer fields PowerPoint puts in common decks; unknown or
 *  unevaluated fields keep their cached display text. */
function fieldTextOf(
  type: string,
  cachedText: string | undefined,
  context: TextFieldContext,
): string {
  if (type === "slidenum" && context.slideNumber != null) return String(context.slideNumber);
  if (type === "datetimeFigureOut" && context.now) {
    return `${context.now.getMonth() + 1}/${context.now.getDate()}/${context.now.getFullYear()}`;
  }
  return cachedText ?? "";
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

function runStyle(
  rp: TextCharacterPropertiesOptions | undefined,
  context: TextFieldContext,
  paragraph: TextParagraphPropertiesOptions = {},
): LayoutTextStyle {
  const defaults = paragraph.defaultRunProperties;
  const family = familyOf(rp?.font ?? defaults?.font, context);
  // FillOptions (parse emits {type:"solid", color}) — colorOf alone only
  // reads the bare-string/flat-color shapes and would drop every parsed run.
  const fillContext = { themeColors: context.themeColors, colorMapping: context.colorMapping };
  const paint = shapeFillOf(rp?.fill ?? defaults?.fill, 0, 0, fillContext).fill;
  const color = typeof paint === "string" ? paint.replace("#", "") : undefined;
  const size = rp?.size ?? defaults?.size;
  const underline = rp?.underline ?? defaults?.underline;
  return {
    family: family ?? DEFAULT_FAMILY,
    sizePx: size != null ? ptToPx(size) : DEFAULT_SIZE_PX,
    ...((rp?.bold ?? defaults?.bold) ? { bold: true } : {}),
    ...((rp?.italic ?? defaults?.italic) ? { italic: true } : {}),
    ...(underline && underline !== "none" ? { underline: true } : {}),
    ...(rp?.strike === "singleStrike" ||
    rp?.strike === "doubleStrike" ||
    defaults?.strike === "singleStrike" ||
    defaults?.strike === "doubleStrike"
      ? { strikethrough: true }
      : {}),
    ...(color ? { color } : {}),
  };
}

/** A bare string sets latin + ea to the same face; the object form reads
 *  latin first, then eastAsia (each slot is a face string or a full font). */
/** The theme font scheme's faces, keyed the way a:rPr's +mn/+mj tokens
 *  reference them. */
export interface ThemeFonts {
  major?: Partial<Record<"latin" | "eastAsian" | "complexScript", string>>;
  minor?: Partial<Record<"latin" | "eastAsian" | "complexScript", string>>;
}

/** A run's font reference as a real family: the theme tokens (+mn-lt and
 *  friends) resolve through the font scheme — the raw token is no family at
 *  all, and feeding it to CSS measure/paint collapses wrapped lines. */
function familyOf(font: RunFont | undefined, context: TextFieldContext): string | undefined {
  const face = (f: TextFont | undefined): string | undefined =>
    typeof f === "string" ? f || undefined : f?.typeface;
  const reference = typeof font === "string" ? font : (face(font?.latin) ?? face(font?.eastAsia));
  const token = /^\+(mn|mj)-(lt|ea|cs)$/.exec(reference ?? "");
  if (!token) return reference || undefined;
  const collection = token[1] === "mj" ? context.themeFonts?.major : context.themeFonts?.minor;
  const slot = token[2] === "lt" ? "latin" : token[2] === "ea" ? "eastAsian" : "complexScript";
  return collection?.[slot] || undefined;
}
