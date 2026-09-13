// Page-number presets: the Page Number split's insertion templates (Word's
// top-of-page / bottom-of-page / current-position building blocks). Story
// presets are whole paragraphs — they replace the story's content; the
// current-position presets are inline fragments that splice into the caret's
// paragraph. The literal words around the fields follow the UI language,
// matching Word's localized building blocks.

import type { JSONContent } from "@docen/docx";

import { resolveLang } from "../ui";

/** A field atom; the painter resolves PAGE/NUMPAGES to live values per page. */
const field = (instruction: string): JSONContent => ({
  type: "inlinePassthrough",
  attrs: { data: JSON.stringify({ simpleField: { instruction } }) },
});

const text = (value: string): JSONContent => ({ type: "text", text: value });

/** The story preset for a `page-top-*` / `page-bottom-*` value: one paragraph
 *  with the placement's alignment (`bar` adds the classic rule under the
 *  number). Left keeps the default alignment. */
export function pageNumberStoryPreset(value: string): JSONContent {
  const alignment = value.endsWith("-left")
    ? undefined
    : value.endsWith("-right")
      ? "right"
      : "center";
  return {
    type: "paragraph",
    ...(alignment || value.endsWith("-bar")
      ? {
          attrs: {
            ...(alignment ? { alignment } : {}),
            // The bottom rule rides the paragraph border, like Word's
            // "Accent Bar" building blocks (size in 1/8 pt, space in pt).
            ...(value.endsWith("-bar")
              ? { border: { bottom: { style: "single", size: 4, color: "000000", space: 4 } } }
              : {}),
          },
        }
      : {}),
    content: [field("PAGE")],
  };
}

/** The current-position preset for a `cur-*` value: inline nodes spliced at
 *  the caret. */
export function pageNumberInlinePreset(value: string): JSONContent[] {
  const zh = resolveLang().toLowerCase().startsWith("zh");
  switch (value) {
    case "cur-page-of":
      return zh ? [text("第 "), field("PAGE"), text(" 页")] : [text("Page "), field("PAGE")];
    case "cur-page-total":
      return zh
        ? [text("第 "), field("PAGE"), text(" 页，共 "), field("NUMPAGES"), text(" 页")]
        : [text("Page "), field("PAGE"), text(" of "), field("NUMPAGES")];
    case "cur-dash":
      return [text("— "), field("PAGE"), text(" —")];
    case "cur-slash":
      return [field("PAGE"), text(" / "), field("NUMPAGES")];
    default:
      return [field("PAGE")];
  }
}
