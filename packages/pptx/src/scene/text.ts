// Shape text projection: a bodyPr's paragraphs as the layout paragraphs the
// painter stacks, with Word-compatible run/paragraph defaults.

import { solidFillOf } from "@docen/core/geometry";
import {
  ptToPx,
  type LayoutInline,
  type LayoutParagraph,
  type LayoutTextStyle,
} from "@docen/layout";
import type {
  RunFont,
  TextBodyOptions,
  TextCharacterPropertiesOptions,
  TextFont,
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

/** A body's paragraphs as projected blocks; the paragraph shape (not the
 *  LayoutBlock union) so consumers read align/spacing directly. */
export function textBlocks(body: TextBodyOptions): LayoutParagraph[] {
  const paragraphs = body.paragraphs ?? (body.text != null ? [body.text] : []);
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
    return {
      kind: "paragraph",
      inline,
      ...(props?.alignment ? { align: ALIGN_MAP[props.alignment] } : {}),
      ...(props?.spaceBefore != null || props?.spaceAfter != null
        ? {
            spacing: {
              beforePx: ptToPx(props.spaceBefore ?? 0),
              afterPx: ptToPx(props.spaceAfter ?? 0),
            },
          }
        : {}),
      // An empty paragraph renders as a blank line — the strut keeps its
      // height (the default style supplies the measuring font).
      ...(inline.length === 0 ? { defaultTextStyle: runStyle(undefined) } : {}),
    };
  });
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
