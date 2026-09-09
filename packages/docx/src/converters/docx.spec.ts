import type { DocumentOptions } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";
import { describe, expect, it } from "vitest";

import { ORDERED_REFERENCE_PREFIX, buildListLevels } from "../extensions/list-numbering";
import { compileDocument, docxExtensions, resolveDocument } from "../index";

/**
 * Compile-side persistence guarantees for the document-level attrs:
 *  - `numbering` round-trips whole (pic bullets, cleanup id, per-definition
 *    instances/aliases) — only abstractNumberings is rebuilt, for regenerated
 *    editor lists that must merge with the source definitions.
 *  - `documentExtras.rawParts` entries whose bytes a JSON round-trip already
 *    corrupted drop instead of reaching office-open's media reader, and the
 *    compile input stays untouched.
 */

function docWithAttrs(attrs: Record<string, unknown>): JSONContent {
  return {
    type: "doc",
    attrs,
    content: [{ type: "paragraph", content: [{ type: "text", text: "x" }] }],
  };
}

describe("compileDocument numbering persistence", () => {
  it("carries the full NumberingOptions through resolve→compile", () => {
    const doc: DocumentOptions = {
      sections: [{ children: [{ paragraph: { text: "x" } }] }],
      numbering: {
        abstractNumberings: [
          {
            reference: "list_1",
            levels: buildListLevels(`${ORDERED_REFERENCE_PREFIX}-1`)!,
            instanceCount: 2,
            aliases: ["2", "3"],
          },
        ],
        numIdMacAtCleanup: 9,
        numPicBullets: [{ numPicBulletId: 1, drawing: "<w:drawing/>" }],
      },
    };
    const json = resolveDocument(doc, docxExtensions);
    const compiled = compileDocument(json, docxExtensions);
    expect(compiled.numbering).toEqual(doc.numbering);
  });

  it("merges regenerated list definitions without dropping source numbering keys", () => {
    const sourceDefinition = { reference: "list_1", levels: buildListLevels("docen-bullet")! };
    const numPicBullets = [{ numPicBulletId: 7, drawing: "<w:drawing/>" }];
    const json = docWithAttrs({
      numbering: {
        abstractNumberings: [sourceDefinition],
        numIdMacAtCleanup: 4,
        numPicBullets,
      },
    });
    (json.content![0] as JSONContent).attrs = {
      numbering: { reference: `${ORDERED_REFERENCE_PREFIX}-1` },
    };

    const compiled = compileDocument(json, docxExtensions);
    const numbering = compiled.numbering!;
    expect(numbering.numPicBullets).toEqual(numPicBullets);
    expect(numbering.numIdMacAtCleanup).toBe(4);
    const references = numbering.abstractNumberings.map((c) => c.reference);
    expect(references).toContain("list_1");
    expect(references.some((r) => r.startsWith(ORDERED_REFERENCE_PREFIX))).toBe(true);
  });

  it("omits numbering when the JSON carries none", () => {
    const compiled = compileDocument({ type: "doc", content: [{ type: "paragraph" }] });
    expect("numbering" in compiled).toBe(false);
  });
});

describe("compileDocument rawParts sanitation", () => {
  const good = { path: "word/theme/theme1.xml", data: new Uint8Array([1, 2, 3]) };
  const text = { path: "customXml/item1.xml", data: "<x/>" };
  const corrupted = { path: "word/embeddings/ole.bin", data: { 0: 80, 1: 75 } };

  it("drops JSON-corrupted entries and never mutates the input attrs", () => {
    const json = docWithAttrs({ documentExtras: { rawParts: [good, corrupted, text] } });
    const compiled = compileDocument(json, docxExtensions);
    expect(compiled.rawParts).toEqual([good, text]);
    expect(json.attrs!.documentExtras).toEqual({ rawParts: [good, corrupted, text] });
  });

  it("drops the rawParts key once every entry is corrupted", () => {
    const json = docWithAttrs({ documentExtras: { rawParts: [corrupted] } });
    const compiled = compileDocument(json, docxExtensions);
    expect("rawParts" in compiled).toBe(false);
  });

  it("keeps a legal rawParts list verbatim", () => {
    const rawParts = [good, text];
    const compiled = compileDocument(docWithAttrs({ documentExtras: { rawParts } }));
    expect(compiled.rawParts).toBe(rawParts);
  });
});
