// Word watermark gallery: the preset table, the header-paragraph stamps
// (text shape and floating picture), and the strip used by Remove Watermark.
// Shapes carry Word's WordPictureWatermark name so the strip can find them
// again.

import type { JSONContent } from "@docen/docx";

import { t } from "../ui";

/** Word's watermark shape name — how Remove Watermark finds the shape again
 *  (Word's own galleries stamp PowerPlusWaterMarkObject* / WordPictureWatermark
 *  on the header shape). */
export const WATERMARK_NAME = "WordPictureWatermark";

/** The gallery presets: Word's diagonal silver text watermarks. Box 7.5" x
 *  1.6" centered on the page, text 96 pt; the diagonal ones rotate -45°
 *  about the box center. The text itself is a translation key — Word's
 *  presets spell CONFIDENTIAL/DRAFT in English locales, 机密/草稿 in Chinese. */
export const WATERMARK_PRESETS: Record<
  string,
  { textKey: string; color: string; rotation: number }
> = {
  "confidential-1": { textKey: "watermark.text-confidential", color: "C0C0C0", rotation: -45 },
  "confidential-2": { textKey: "watermark.text-confidential", color: "C0C0C0", rotation: 0 },
  "confidential-3": { textKey: "watermark.text-confidential", color: "C00000", rotation: -45 },
  urgent: { textKey: "watermark.text-urgent", color: "C0C0C0", rotation: -45 },
  asap: { textKey: "watermark.text-asap", color: "C0C0C0", rotation: -45 },
  draft: { textKey: "watermark.text-draft", color: "C0C0C0", rotation: -45 },
  sample: { textKey: "watermark.text-sample", color: "C0C0C0", rotation: -45 },
};

/** The custom watermark spec (Word's dialog): text or picture. */
export interface WatermarkTextSpec {
  text: string;
  font?: string;
  /** Point size; "auto" scales the text to fill the 7.5"-wide box. */
  size?: number | "auto";
  /** Hex RRGGBB. */
  color: string;
  /** Word's 版式: 斜式 rotates -45°, 水平 keeps 0°. */
  diagonal: boolean;
  /** Word's 半透明 — modeled as a 50% blend toward white (the run color is
   *  plain hex, and a blended silver reads identically on the page). */
  semiTransparent: boolean;
}

export interface WatermarkPictureSpec {
  /** Data URL of the picked image. */
  src: string;
  /** Scale of the natural size; "auto" fits the 7.5" text width. */
  scale: number | "auto";
  /** Word's 冲蚀 — stamps a luminance effect (bright +50%, contrast −50%).
   *  The canvas painter doesn't consume blip effects yet, so the editor
   *  shows the raw image; Word renders the washed-out version. */
  washout: boolean;
}

/** Natural image size probe for the picture stamp ("auto" scaling needs the
 *  real aspect before the paragraph is built). */
export function probeImageSize(src: string): Promise<{ w: number; h: number }> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = (): void => resolve({ w: img.naturalWidth, h: img.naturalHeight });
    img.onerror = reject;
    img.src = src;
  });
}

/** Blend a hex color toward white by `amount` (0..1) — the visual stand-in
 *  for Word's wash-out/半透明, which the plain-hex run color can't carry. */
export function mixWhite(hex: string, amount: number): string {
  const n = hex.replace("#", "");
  const full =
    n.length === 3
      ? n
          .split("")
          .map((c) => c + c)
          .join("")
      : n;
  const ch = (i: number): number =>
    Math.round(parseInt(full.slice(i, i + 2), 16) * (1 - amount) + 255 * amount);
  return [ch(0), ch(2), ch(4)]
    .map((v) => v.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}

/** Word's 字号 "自动": fit the text to the 7.5"-wide watermark box. */
export function autoTextSize(text: string): number {
  // CJK glyphs are square, latin roughly half — approximate with a 0.62
  // em-width average, tuned so CONFIDENTIAL (11 chars) lands at 96 pt like
  // the presets.
  const emWidths = Array.from(text).reduce(
    (sum, c) => sum + (c.charCodeAt(0) > 0x2e7f ? 1 : 0.62),
    0,
  );
  const size = Math.floor((540 / Math.max(1, emWidths)) * 1.9);
  return Math.max(24, Math.min(120, size));
}

/** One preset's header paragraph: the watermark shape alone on its line. */
export function watermarkPara(spec: {
  textKey: string;
  color: string;
  rotation: number;
}): JSONContent {
  return customTextWatermarkPara({
    text: t(spec.textKey),
    color: spec.color,
    diagonal: spec.rotation < 0,
    semiTransparent: false,
    size: 96,
  });
}

/** The custom dialog's text stamp: a centered behind-document shape with the
 *  spec'd run (font/size/color, blended when 半透明). */
export function customTextWatermarkPara(spec: WatermarkTextSpec): JSONContent {
  const color = spec.semiTransparent ? mixWhite(spec.color, 0.5) : spec.color;
  const size = spec.size === "auto" || spec.size == null ? autoTextSize(spec.text) : spec.size;
  return {
    type: "paragraph",
    content: [
      {
        type: "wpsShape",
        attrs: {
          wpsShape: {
            name: WATERMARK_NAME,
            transformation: { width: 6858000, height: 1463040, rotation: spec.diagonal ? -45 : 0 },
            floating: {
              horizontalPosition: { relative: "page", align: "center" },
              verticalPosition: { relative: "page", align: "center" },
              wrap: { type: "none" },
              behindDocument: true,
              allowOverlap: true,
            },
            fill: { type: "none" },
            outline: { type: "none" },
          },
        },
        content: [
          {
            type: "paragraph",
            attrs: { alignment: "center" },
            content: [
              {
                type: "text",
                text: spec.text,
                marks: [
                  {
                    type: "textStyle",
                    attrs: {
                      color,
                      size,
                      ...(spec.font ? { font: spec.font } : {}),
                    },
                  },
                ],
              },
            ],
          },
        ],
      },
    ],
  };
}

/** The custom dialog's picture stamp: a behind-document floating picture
 *  centered on the page, scaled from the natural size, optionally washed
 *  out (luminance effect — rendered by Word, pending in the canvas painter). */
export function pictureWatermarkPara(
  spec: WatermarkPictureSpec,
  natural: { w: number; h: number },
): JSONContent {
  // "auto" fits the 7.5" box width; px = EMU / 9525 for the Tiptap attrs.
  const boxPx = 6858000 / 9525;
  const scale =
    spec.scale === "auto" || spec.scale == null ? Math.min(1, boxPx / natural.w) : spec.scale;
  const width = Math.round(natural.w * scale);
  const height = Math.round(natural.h * scale);
  return {
    type: "paragraph",
    content: [
      {
        type: "image",
        attrs: {
          src: spec.src,
          width,
          height,
          alt: WATERMARK_NAME,
          floating: {
            horizontalPosition: { relative: "page", align: "center" },
            verticalPosition: { relative: "page", align: "center" },
            wrap: { type: "none" },
            behindDocument: true,
            allowOverlap: true,
          },
          ...(spec.washout ? { blipEffects: { luminance: { bright: 50, contrast: -50 } } } : {}),
        },
      },
    ],
  };
}

/** The watermark marker on either stamp kind: the text shape's name or the
 *  picture's altText name (alt → w:name on export). */
export function isWatermarkNode(child: JSONContent): boolean {
  if (child.type === "wpsShape")
    return (child.attrs as { wpsShape?: { name?: string } })?.wpsShape?.name === WATERMARK_NAME;
  if (child.type === "image") return (child.attrs as { alt?: string })?.alt === WATERMARK_NAME;
  return false;
}

/** A header paragraph with its watermark shape stripped (null when nothing
 *  but the watermark rode the paragraph). */
function stripWatermarkPara(para: JSONContent | undefined): JSONContent | null {
  if (!para || para.type !== "paragraph") return para ?? null;
  const content = (para.content ?? []).filter((child) => !isWatermarkNode(child));
  if (content.length === 0 && (para.content ?? []).length > 0) return null;
  return { ...para, content };
}

/** One section's header slots after a watermark stamp: existing watermark
 *  shapes stripped (re-inserting replaces, Word's re-insert does the same),
 *  then — a stamp given — the shape appended to every slot, including ones
 *  the section never carried (Word stamps every header part; a slot with no
 *  paragraphs gets its strut). */
export function stampHeaderSlots(
  slots: Record<string, JSONContent[] | undefined> | undefined,
  para: JSONContent | null,
): Record<string, JSONContent[]> {
  const out: Record<string, JSONContent[]> = {};
  for (const [slot, paras] of Object.entries(slots ?? {})) {
    out[slot] = paras?.map(stripWatermarkPara).filter((p) => p != null) ?? [];
  }
  if (para) {
    for (const slot of Object.keys(out)) out[slot]!.push(para);
    for (const slot of ["default", "first", "even"]) {
      if (!out[slot]) out[slot] = [para];
    }
  }
  return out;
}
