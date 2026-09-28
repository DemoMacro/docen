import type { DocumentOptions, ParagraphChild, SectionChild } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { projectDocument } from "./document";

// The field descriptor contract: every w:fldSimple / complexField atom carries
// its instruction verbatim plus the cached result, PAGE/NUMPAGES keep the
// dynamic page marker and measuring placeholder, and the Alt+F9 field-code
// display projects the instruction in place of the result.

const inlineOf = (children: SectionChild[], showFieldCodes?: boolean) => {
  const doc: DocumentOptions = { sections: [{ children }] };
  const blocks = projectDocument(doc, undefined, showFieldCodes).sections[0]!.blocks;
  const para = blocks[0];
  if (para?.kind !== "paragraph") throw new Error("expected paragraph");
  return para.inline;
};

const fields = (): SectionChild[] => [
  {
    paragraph: {
      children: [
        { complexField: { instruction: "PAGE \\* MERGEFORMAT", result: "3" } },
        { complexField: { instruction: "NUMPAGES", result: "9" } },
        { simpleField: { instruction: "AUTHOR", cachedValue: "作者" } },
        { simpleField: { instruction: "REF _Ref00000001 \\h", cachedValue: "" } },
      ],
    },
  },
];

describe("field descriptor projection", () => {
  it("carries instruction + cached result on every field atom", () => {
    const [page, numPages, author, ref] = inlineOf(fields());
    expect(page).toMatchObject({
      kind: "text",
      text: "0",
      field: "page",
      instruction: "PAGE \\* MERGEFORMAT",
      result: "3",
    });
    expect(numPages).toMatchObject({
      kind: "text",
      text: "0",
      field: "numPages",
      instruction: "NUMPAGES",
      result: "9",
    });
    expect(author).toMatchObject({
      kind: "text",
      text: "作者",
      instruction: "AUTHOR",
      result: "作者",
    });
    // An empty cache still projects the atom (empty text) so the render pass
    // and the update commands can resolve it.
    expect(ref).toMatchObject({
      kind: "text",
      text: "",
      instruction: "REF _Ref00000001 \\h",
      result: "",
    });
  });

  it("projects the instruction verbatim under the field-code display", () => {
    const [page, numPages, author, ref] = inlineOf(fields(), true);
    expect(page).toMatchObject({ kind: "text", text: "PAGE \\* MERGEFORMAT" });
    expect((page as { field?: string }).field).toBeUndefined();
    expect(numPages).toMatchObject({ kind: "text", text: "NUMPAGES" });
    expect(author).toMatchObject({ kind: "text", text: "AUTHOR" });
    expect(ref).toMatchObject({ kind: "text", text: "REF _Ref00000001 \\h" });
  });

  it("keeps non-field passthrough atoms untouched", () => {
    const inline = inlineOf([
      {
        paragraph: {
          children: [
            { bookmarkStart: { id: 1, name: "bm" } },
            { text: "文本" },
            { bookmarkEnd: { id: 1 } },
          ],
        },
      },
    ]);
    expect(inline).toHaveLength(1);
    expect(inline[0]).toMatchObject({ kind: "text", text: "文本" });
    expect((inline[0] as { instruction?: string }).instruction).toBeUndefined();
  });

  it("keeps section numbering atoms measurable with an empty cache", () => {
    const [sectionPages, section] = inlineOf([
      {
        paragraph: {
          children: [
            { simpleField: { instruction: "SECTIONPAGES", cachedValue: "" } },
            { simpleField: { instruction: "SECTION", cachedValue: "3" } },
          ],
        },
      },
    ]);
    expect(sectionPages).toMatchObject({
      kind: "text",
      text: "0",
      instruction: "SECTIONPAGES",
      result: "",
    });
    // A non-empty cache stays the measuring text (Word measures the result).
    expect(section).toMatchObject({ kind: "text", text: "3", instruction: "SECTION" });
  });

  it("prefers a SECTION field's structured result runs over the placeholder", () => {
    const children = [
      {
        complexField: {
          instruction: "SECTION",
          result: "3",
          resultRunsXml: "<w:r><w:t>II</w:t></w:r>",
        },
      },
    ] as unknown as ParagraphChild[];
    const inline = inlineOf([{ paragraph: { children } }]);
    // The structured run re-hydrates; the flat cache must not leak beside it.
    expect(inline).toHaveLength(1);
    expect(inline[0]).toMatchObject({ kind: "text", text: "II" });
  });
});
