// Word watermark gallery: the preset table, the header-paragraph stamp, and
// the strip used by Remove Watermark. Shapes carry Word's
// WordPictureWatermark name so the strip can find them again.

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

/** One preset's header paragraph: the watermark shape alone on its line. */
export function watermarkPara(spec: {
  textKey: string;
  color: string;
  rotation: number;
}): JSONContent {
  return {
    type: "paragraph",
    content: [
      {
        type: "wpsShape",
        attrs: {
          wpsShape: {
            name: WATERMARK_NAME,
            transformation: { width: 6858000, height: 1463040, rotation: spec.rotation },
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
                text: t(spec.textKey),
                marks: [{ type: "textStyle", attrs: { color: spec.color, size: 96 } }],
              },
            ],
          },
        ],
      },
    ],
  };
}

/** A header paragraph with its watermark shape stripped (null when nothing
 *  but the watermark rode the paragraph). */
export function stripWatermarkPara(para: JSONContent | undefined): JSONContent | null {
  if (!para || para.type !== "paragraph") return para ?? null;
  const content = (para.content ?? []).filter(
    (child) =>
      !(
        child.type === "wpsShape" &&
        (child.attrs as { wpsShape?: { name?: string } })?.wpsShape?.name === WATERMARK_NAME
      ),
  );
  if (content.length === 0 && (para.content ?? []).length > 0) return null;
  return { ...para, content };
}
