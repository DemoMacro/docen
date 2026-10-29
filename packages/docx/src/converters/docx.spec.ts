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

describe("compileDocument inline safety", () => {
  it("rejects unknown inline nodes instead of dropping them", () => {
    const json = docWithAttrs({});
    (json.content![0] as JSONContent).content = [{ type: "unsupportedInline" }];
    expect(() => compileDocument(json, docxExtensions)).toThrow(TypeError);
    expect(() => compileDocument(json, docxExtensions)).toThrow("Unsupported inline node");
  });
});

describe("resolveDocument empty sections", () => {
  it("preserves document-level attrs when no sections exist", () => {
    const rawPart = { path: "customXml/item1.xml", data: "<x/>" };
    const json = resolveDocument(
      {
        sections: [],
        styles: { paragraphStyles: [] },
        numbering: { abstractNumberings: [] },
        rawParts: [rawPart],
      } as never,
      docxExtensions,
    );
    const attrs = json.attrs as Record<string, any>;
    expect(json.content).toEqual([{ type: "paragraph" }]);
    expect(attrs.styles).toEqual({ paragraphStyles: [] });
    expect(attrs.numbering).toEqual({ abstractNumberings: [] });
    expect(attrs.documentExtras.rawParts).toHaveLength(1);
    expect(compileDocument(json, docxExtensions).rawParts).toEqual([rawPart]);
  });
});

describe("document binary persistence", () => {
  const good = { path: "word/theme/theme1.xml", data: new Uint8Array([1, 2, 3]) };
  const text = { path: "customXml/item1.xml", data: "<x/>" };
  const corrupted = { path: "word/embeddings/ole.bin", data: { 0: 80, 1: 75 } };

  it("drops legacy JSON-corrupted entries and never mutates the input attrs", () => {
    const json = docWithAttrs({ documentExtras: { rawParts: [good, corrupted, text] } });
    const compiled = compileDocument(json, docxExtensions);
    expect(compiled.rawParts?.[0]).toEqual(good);
    expect(compiled.rawParts?.[1]).toEqual(text);
    expect(json.attrs!.documentExtras).toEqual({ rawParts: [good, corrupted, text] });
  });

  it("drops the rawParts key once every entry is corrupted", () => {
    const json = docWithAttrs({ documentExtras: { rawParts: [corrupted] } });
    const compiled = compileDocument(json, docxExtensions);
    expect("rawParts" in compiled).toBe(false);
  });

  it("keeps a legal rawParts list verbatim before JSON persistence", () => {
    const rawParts = structuredClone([good, text]);
    const resolved = resolveDocument(
      {
        sections: [{ children: [{ paragraph: { text: "x" } }] }],
        rawParts,
      },
      docxExtensions,
    );
    expect(resolved.attrs?.documentExtras).toEqual({
      rawParts: [{ ...good, data: "data:application/octet-stream;base64,AQID" }, text],
    });
  });

  it("survives a JSON round-trip without losing rawPart or altChunk bytes", () => {
    const bytes = new Uint8Array([0, 255, 8, 16]);
    const rawParts = [good, text];
    const source: DocumentOptions = {
      sections: [
        {
          children: [{ altChunk: { data: bytes, contentType: "text/html", extension: "html" } }],
        },
      ],
      rawParts,
    };
    const persisted = JSON.parse(
      JSON.stringify(resolveDocument(source, docxExtensions)),
    ) as JSONContent;
    const compiled = compileDocument(persisted, docxExtensions);
    const child = compiled.sections[0]?.children[0] as {
      altChunk?: { data?: Uint8Array };
    };
    expect(Array.from(child.altChunk?.data ?? [])).toEqual(Array.from(bytes));
    expect(compiled.rawParts?.[0]).toEqual({
      ...good,
      data: "data:application/octet-stream;base64,AQID",
    });
    expect(compiled.rawParts?.[1]).toEqual(text);
  });

  it("serializes document-level binaries for JSON persistence", () => {
    const source: DocumentOptions = {
      sections: [{ children: [{ paragraph: { text: "x" } }] }],
      background: {
        rawMedia: [{ fileName: "backing.bin", data: new Uint8Array([1, 2, 3]), type: "png" }],
      },
      fonts: [{ name: "Custom", embedRegular: { data: new Uint8Array([4, 5, 6]) } }],
    };
    const persisted = JSON.parse(
      JSON.stringify(resolveDocument(source, docxExtensions)),
    ) as JSONContent;
    const compiled = compileDocument(persisted, docxExtensions);
    const extras = compiled as typeof source;
    expect(Array.from(extras.background?.rawMedia?.[0]?.data as Uint8Array)).toEqual([1, 2, 3]);
    expect(Array.from(extras.fonts?.[0]?.embedRegular?.data as Uint8Array)).toEqual([4, 5, 6]);
  });

  it("keeps tracked hyperlink containers nested", () => {
    const json = docWithAttrs({});
    const paragraph = json.content![0]! as JSONContent;
    paragraph.content = [
      {
        type: "text",
        text: "linked",
        marks: [
          {
            type: "insertion",
            attrs: { id: 3, author: "Ada", date: "2026-01-01T00:00:00Z" },
          },
          { type: "link", attrs: { href: "https://example.com" } },
        ],
      },
    ];
    const compiled = compileDocument(json, docxExtensions);
    const child = compiled.sections[0]?.children[0] as {
      paragraph?: { children?: Array<Record<string, unknown>> };
    };
    const insertion = child.paragraph?.children?.[0]?.insertion as {
      children?: Array<{ hyperlink?: { url?: string } }>;
    };
    expect(insertion.children?.[0]?.hyperlink?.url).toBe("https://example.com");
  });
});
