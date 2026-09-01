// Style cascade — direct pPr → style chain → docDefaults — the same
// mergeStyleChain resolution the editor's measure side uses, plus the
// per-run rPr resolution the text runs and the ¶-mark strut build on.

import type { LayoutParagraph, LayoutTextStyle } from "@docen/layout";
import type { StylesOptions } from "@office-open/docx";

import {
  defaultParagraphStyleId,
  indexParagraphStyles,
  mergeStyleChain,
} from "../../style-cascade";
import { colorOf, isRecord, measureTwip, num, str, type Rec } from "./guards";

// ── style cascade (direct pPr → style chain → docDefaults) ──

/** `default.document` — the docDefaults object (run directly on it, paragraph
 *  props nested one level down under `paragraph`). */
export function docDefaultsOf(styles: StylesOptions | undefined): Rec {
  const doc = styles?.default?.document;
  return isRecord(doc) ? doc : {};
}

/** w:jc (AlignmentType, ST_Jc) → the engine's alignment semantics. The
 *  kashida/thai/numericTab variants (Arabic elongation, Thai word-break
 *  justification, list tab alignment) have no faithful canvas algorithm —
 *  they fall back to the left default until one lands. */
const ALIGN_TO_LAYOUT = {
  left: "left",
  start: "left",
  right: "right",
  end: "right",
  center: "center",
  both: "both",
  distribute: "distribute",
} as const;

/** The merged {run, paragraph} for a style id — the same mergeStyleChain the
 *  editor's measure.ts resolves, so projection and pagination share one
 *  cascade. A style-less paragraph resolves to the default paragraph style
 *  (usually Normal); docDefaults sits UNDER the chain, not in it. */
export function styleChainOf(
  styles: StylesOptions | undefined,
  styleId: string | null | undefined,
) {
  if (!styles) return { run: {}, paragraph: {} };
  return mergeStyleChain(indexParagraphStyles(styles), styleId || defaultParagraphStyleId(styles));
}

export const pick = (layers: Rec[], key: string): unknown => {
  for (const layer of layers) if (layer[key] != null) return layer[key];
  return undefined;
};

/** The cascaded w:jc value → the engine's alignment (undefined → left). */
export function alignOf(jc: unknown): LayoutParagraph["align"] {
  if (typeof jc !== "string") return undefined;
  return jc in ALIGN_TO_LAYOUT ? ALIGN_TO_LAYOUT[jc as keyof typeof ALIGN_TO_LAYOUT] : undefined;
}

// ── run/style resolution ──

/** OOXML font: string or rFonts {ascii, hAnsi, eastAsia} → engine slots. */
export type FontAttr = string | Rec | null | undefined;

/** An unknown font pick → the FontAttr domain (string or rFonts record). */
export function fontAttr(v: unknown): FontAttr {
  return isRecord(v) || typeof v === "string" ? v : undefined;
}

/** Resolve a font pick against its fallback: a record with no usable slot
 *  (an empty rFonts shell from a round-tripped run) counts as unspecified,
 *  so the chain's face survives instead of shadowing it with empty slots. */
export function toFamily(font: FontAttr, def: FontAttr): LayoutTextStyle["family"] | undefined {
  const f = font ?? def;
  if (typeof f === "string") return f || undefined;
  const latin = str(f?.ascii) ?? str(f?.hAnsi);
  const eastAsia = str(f?.eastAsia);
  return latin || eastAsia ? { latin, eastAsia } : undefined;
}

export interface RunStyle {
  sizePt?: number;
  font?: FontAttr;
  characterSpacingTw?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  highlight?: string;
  shadingFill?: string;
  underline?: boolean;
  strikethrough?: boolean;
  verticalAlign?: "superscript" | "subscript";
}

/** rPr (a run's own, or the ¶-mark/paragraph default) → resolved fields.
 *  Toggle fields stay three-state: an explicit `w:b w:val="0"` resolves to
 *  false so it BEATS an inherited style bold — folding it to undefined would
 *  let the style chain's bold bleed through (Word: direct > style > doc). */
export function runStyleOf(rPr: Rec): RunStyle {
  const underline = isRecord(rPr.underline) ? rPr.underline.type !== "none" : undefined;
  const tri = (v: unknown): boolean | undefined => (v === undefined ? undefined : v === true);
  return {
    sizePt: num(rPr.size),
    font: fontAttr(rPr.font),
    characterSpacingTw: measureTwip(rPr.characterSpacing),
    bold: tri(rPr.bold),
    italic: tri(rPr.italic),
    color: colorOf(rPr.color),
    highlight: str(rPr.highlight),
    // Direct hex fill only — a themeFill-bound shading needs the theme palette
    // resolved with the document context, which runStyleOf doesn't carry.
    shadingFill: isRecord(rPr.shading) ? str(rPr.shading.fill) : undefined,
    underline,
    strikethrough:
      rPr.strike === true || rPr.doubleStrike === true
        ? true
        : rPr.strike === false && rPr.doubleStrike === false
          ? false
          : undefined,
    verticalAlign:
      rPr.verticalAlign === "superscript" || rPr.verticalAlign === "subscript"
        ? rPr.verticalAlign
        : undefined,
  };
}
