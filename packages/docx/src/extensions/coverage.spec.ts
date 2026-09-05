import type { DocumentOptions, ParagraphChild, RunOptions, SectionChild } from "@office-open/docx";
import { generateDocumentSync, parseDocument } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";
import { describe, expect, it } from "vitest";

import { compileDocument, docxExtensions, resolveDocument } from "../index";
import {
  PARAGRAPH_CHILD_DISPOSITIONS,
  RUN_CHILDREN_DROPPED,
  SECTION_CHILD_DISPOSITIONS,
} from "./coverage";

/**
 * Round-trip proof of the coverage registry: coverage.ts claims a disposition
 * for every office-open union branch, this spec drives one fixture per branch
 * through resolve → compile and proves the claim.
 *
 *  - editable: the owning node/mark engages on resolve (NOT the passthrough
 *    atom), and the compiled output keeps the branch + a probe field.
 *  - passthrough: the verbatim atom carries it and compile deep-equals the
 *    original fixture (byte-faithful).
 *  - dropped (run children): the shape is really absent after compile — guards
 *    against the table drifting from reality in the forgiving direction.
 *
 * The fixture/probe records are keyed by the disposition tables, so a tag
 * moving between dispositions (or a new office-open branch) fails to compile
 * until the spec is taught the branch — the registry claims, this file proves.
 */

// 1×1 PNG header bytes — content is opaque to the pipeline (base64 round-trip).
const PNG_BYTES = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

/** Resolve a single-section document holding exactly these children, then
 *  compile it back. Returns both legs for route/field assertions. */
function roundTrip(children: SectionChild[]) {
  const doc: DocumentOptions = { sections: [{ children }] };
  const json = resolveDocument(doc, docxExtensions);
  const out = compileDocument(json, docxExtensions);
  return { json, compiled: out.sections[out.sections.length - 1].children };
}

/** Resolve a paragraph holding exactly one inline child, compile back. */
function roundTripInline(child: ParagraphChild) {
  return roundTrip([{ paragraph: { children: [child] } }]);
}

/** Node and mark type names present anywhere in a resolved document. */
function collectTypes(node: JSONContent, into = new Set<string>()): Set<string> {
  if (node?.type) into.add(node.type);
  for (const mark of node?.marks ?? []) if (mark.type) into.add(mark.type);
  for (const child of node?.content ?? []) collectTypes(child, into);
  return into;
}

/** First inline child of the single compiled paragraph. */
function firstInline(compiled: SectionChild[]): ParagraphChild | undefined {
  const para = compiled[0] as { paragraph: { children?: ParagraphChild[] } };
  return para.paragraph.children?.[0];
}

// ── SectionChild (block level) ──

const SECTION_FIXTURES: Record<keyof typeof SECTION_CHILD_DISPOSITIONS, () => SectionChild> = {
  paragraph: () => ({ paragraph: { text: "plain" } }),
  table: () => ({
    table: { rows: [{ cells: [{ children: [{ paragraph: { text: "cell" } }] }] }] },
  }),
  toc: () => ({
    toc: { captionLabel: "Table", entries: [{ paragraph: { text: "entry" } }] },
  }),
  sdt: () => ({
    sdt: { properties: { tag: "other" }, children: [{ paragraph: { text: "x" } }] },
  }),
  textbox: () => ({
    textbox: { text: "box", children: [{ paragraph: { text: "in-box" } }] },
  }),
  altChunk: () => ({
    altChunk: { data: "PGh0bWw+", contentType: "text/html", extension: "html" },
  }),
  customXml: () => ({ customXml: { element: "CX", children: [] } }),
  bookmarkStart: () => ({ bookmarkStart: { id: 1, name: "bm" } }),
  bookmarkEnd: () => ({ bookmarkEnd: { id: 1 } }),
  rawXml: () => ({ rawXml: "<w:p><w:r><w:t>raw</w:t></w:r></w:p>" }),
};

/** Editable claims: the owning node type on the resolve side + a narrow field
 *  probe on the compile side (editor-equivalent rewrites are allowed — the
 *  branch identity and a representative field must survive). */
type SectionEditable = {
  [K in keyof typeof SECTION_CHILD_DISPOSITIONS as (typeof SECTION_CHILD_DISPOSITIONS)[K] extends {
    editable: string;
  }
    ? K
    : never]: {
    node: string;
    probe: (compiled: SectionChild[]) => void;
  };
};

const SECTION_EDITABLE: SectionEditable = {
  paragraph: {
    node: "paragraph",
    probe: (compiled) => {
      // compile collapses a text-only paragraph to the string shorthand.
      expect((compiled[0] as { paragraph: string | { text?: string } }).paragraph).toBe("plain");
    },
  },
  table: {
    node: "table",
    probe: (compiled) => {
      const table = (
        compiled[0] as {
          table: { rows: { cells: { children: { paragraph: string | { text?: string } }[] }[] }[] };
        }
      ).table;
      expect(table.rows).toHaveLength(1);
      expect(table.rows[0].cells[0].children[0].paragraph).toBe("cell");
    },
  },
  toc: {
    node: "tocField",
    probe: (compiled) => {
      const toc = (
        compiled[0] as {
          toc: { captionLabel?: string; entries: { paragraph: string | { text?: string } }[] };
        }
      ).toc;
      expect(toc.captionLabel).toBe("Table");
      // entries collapse to the string shorthand like plain paragraphs.
      expect(toc.entries[0].paragraph).toBe("entry");
    },
  },
};

describe("SectionChild dispositions", () => {
  for (const [tag, disposition] of Object.entries(SECTION_CHILD_DISPOSITIONS)) {
    it(`${tag} → ${JSON.stringify(disposition)}`, () => {
      const fixture = SECTION_FIXTURES[tag as keyof typeof SECTION_FIXTURES]();
      const { json, compiled } = roundTrip([fixture]);
      const types = collectTypes(json);
      if ("editable" in disposition) {
        const claim = SECTION_EDITABLE[tag as keyof typeof SECTION_EDITABLE];
        expect(types.has(claim.node), `resolved doc must contain a ${claim.node} node`).toBe(true);
        claim.probe(compiled);
      } else {
        expect(types.has("passthrough"), "must ride the block passthrough atom").toBe(true);
        expect(compiled).toEqual([fixture]);
      }
    });
  }
});

// ── ParagraphChild (inline level) ──

const INLINE_FIXTURES: Record<keyof typeof PARAGRAPH_CHILD_DISPOSITIONS, () => ParagraphChild> = {
  // Editable branches.
  picture: () => ({
    picture: {
      type: "png",
      data: PNG_BYTES,
      transformation: { width: 100, height: 60 },
      altText: { name: "probe" },
    },
  }),
  hyperlink: () => ({
    hyperlink: { url: "https://example.com", tooltip: "probe", children: [{ text: "link" }] },
  }),
  insertion: () => ({
    insertion: {
      id: 1,
      author: "a",
      date: "2024-01-01T00:00:00Z",
      children: [{ text: "ins" }],
    },
  }),
  deletion: () => ({
    deletion: { id: 2, author: "a", date: "2024-01-01T00:00:00Z", children: [{ text: "del" }] },
  }),
  pageBreak: () => ({ pageBreak: true }),
  columnBreak: () => ({ columnBreak: true }),
  wpsShape: () => ({
    wpsShape: { children: [{ text: "box" }], transformation: { width: 100, height: 60 } },
  }),
  wpgGroup: () => ({
    wpgGroup: { children: [], transformation: { width: 100, height: 60 } },
  }),

  // Passthrough branches. Fixtures with children exercise the run-catch-all
  // tag gate: a tagged branch carrying text/children/break keys must still
  // ride the verbatim atom, not get unwrapped as a plain run.
  bookmarkStart: () => ({ bookmarkStart: { id: 1, name: "bm" } }),
  bookmarkEnd: () => ({ bookmarkEnd: { id: 1 } }),
  bookmark: () => ({ bookmark: { name: "bm" } }),
  chart: () => ({
    chart: {
      type: "bar",
      series: [{ values: [1, 2] }],
      transformation: { width: 100, height: 60 },
    },
  }),
  smartArt: () => ({ smartArt: { nodes: [], transformation: { width: 100, height: 60 } } }),
  math: () => ({ math: { display: true } }),
  symbolRun: () => ({ symbolRun: { char: "§" } }),
  footnoteReference: () => ({ footnoteReference: 1 }),
  endnoteReference: () => ({ endnoteReference: 1 }),
  commentRangeStart: () => ({ commentRangeStart: { id: 1 } }),
  commentRangeEnd: () => ({ commentRangeEnd: { id: 1 } }),
  commentReference: () => ({ commentReference: 1 }),
  comment: () => ({ comment: { author: "a", children: [{ text: "note" }] } }),
  object: () => ({ object: { shapeId: "_x0000_s1026" } }),
  pict: () => ({ pict: { vmlFallback: "<v:rect/>" } }),
  contentPart: () => ({ contentPart: { referenceId: "rId1" } }),
  proofErr: () => ({ proofErr: "spellStart" }),
  positionalTab: () => ({
    positionalTab: { alignment: "left", relativeTo: "margin", leader: "none" },
  }),
  permStart: () => ({ permStart: { id: 1 } }),
  permEnd: () => ({ permEnd: 1 }),
  moveFromRangeStart: () => ({ moveFromRangeStart: { id: 1, name: "mv" } }),
  moveFromRangeEnd: () => ({ moveFromRangeEnd: { id: 1 } }),
  moveToRangeStart: () => ({ moveToRangeStart: { id: 1, name: "mv" } }),
  moveToRangeEnd: () => ({ moveToRangeEnd: { id: 1 } }),
  movedFrom: () => ({
    movedFrom: { id: 1, author: "a", date: "2024-01-01T00:00:00Z", children: [] },
  }),
  movedTo: () => ({
    movedTo: { id: 1, author: "a", date: "2024-01-01T00:00:00Z", children: [] },
  }),
  moveFrom: () => ({
    moveFrom: { author: "a", date: "2024-01-01T00:00:00Z", name: "mv" },
  }),
  moveTo: () => ({ moveTo: { author: "a", date: "2024-01-01T00:00:00Z", name: "mv" } }),
  customXmlInsRangeStart: () => ({ customXmlInsRangeStart: { id: 1 } }),
  customXmlInsRangeEnd: () => ({ customXmlInsRangeEnd: 1 }),
  customXmlDelRangeStart: () => ({ customXmlDelRangeStart: { id: 1 } }),
  customXmlDelRangeEnd: () => ({ customXmlDelRangeEnd: 1 }),
  customXmlMoveFromRangeStart: () => ({ customXmlMoveFromRangeStart: { id: 1 } }),
  customXmlMoveFromRangeEnd: () => ({ customXmlMoveFromRangeEnd: 1 }),
  customXmlMoveToRangeStart: () => ({ customXmlMoveToRangeStart: { id: 1 } }),
  customXmlMoveToRangeEnd: () => ({ customXmlMoveToRangeEnd: 1 }),
  simpleField: () => ({ simpleField: { instruction: "PAGE" } }),
  formField: () => ({ formField: { name: "ff1" } }),
  complexField: () => ({ complexField: { instruction: "PAGE" } }),
  seqIdentifier: () => ({ seqIdentifier: "Figure" }),
  pageReference: () => ({ pageReference: { bookmarkId: "bm" } }),
  dir: () => ({ dir: { val: "ltr", children: ["body"] } }),
  bdo: () => ({ bdo: { val: "rtl", children: ["body"] } }),
  smartTag: () => ({ smartTag: { element: "ST", children: ["body"] } }),
  customXml: () => ({ customXml: { element: "CX", children: ["body"] } }),
  sdt: () => ({ sdt: { properties: { alias: "inline-sdt" }, children: ["body"] } }),
  subDoc: () => ({ subDoc: { data: "x" } }),
  rawXml: () => ({ rawXml: '<w:fldSimple instr="PAGE"/>' }),
};

type InlineEditable = {
  [K in keyof typeof PARAGRAPH_CHILD_DISPOSITIONS as (typeof PARAGRAPH_CHILD_DISPOSITIONS)[K] extends {
    editable: string;
  }
    ? K
    : never]: {
    marker: string;
    probe: (out: ParagraphChild | undefined) => void;
  };
};

const INLINE_EDITABLE: InlineEditable = {
  picture: {
    marker: "image",
    probe: (out) => {
      const picture = (
        out as { picture: { type?: string; data?: Uint8Array; altText?: { name?: string } } }
      ).picture;
      expect(picture.type).toBe("png");
      expect(Array.from(picture.data ?? [])).toEqual(Array.from(PNG_BYTES));
      expect(picture.altText?.name).toBe("probe");
    },
  },
  hyperlink: {
    marker: "link",
    probe: (out) => {
      const hyperlink = (
        out as {
          hyperlink: { url?: string; tooltip?: string; children?: { text?: string }[] };
        }
      ).hyperlink;
      expect(hyperlink.url).toBe("https://example.com");
      expect(hyperlink.tooltip).toBe("probe");
      expect(hyperlink.children?.[0]?.text).toBe("link");
    },
  },
  insertion: {
    marker: "insertion",
    probe: (out) => {
      const insertion = (
        out as { insertion: { id?: number; author?: string; children?: { text?: string }[] } }
      ).insertion;
      expect(insertion.id).toBe(1);
      expect(insertion.author).toBe("a");
      expect(insertion.children?.[0]?.text).toBe("ins");
    },
  },
  deletion: {
    marker: "deletion",
    probe: (out) => {
      const deletion = (
        out as { deletion: { id?: number; author?: string; children?: { text?: string }[] } }
      ).deletion;
      expect(deletion.id).toBe(2);
      expect(deletion.author).toBe("a");
      expect(deletion.children?.[0]?.text).toBe("del");
    },
  },
  pageBreak: {
    marker: "pageBreak",
    probe: (out) => expect(out).toEqual({ pageBreak: true }),
  },
  columnBreak: {
    marker: "columnBreak",
    probe: (out) => expect(out).toEqual({ columnBreak: true }),
  },
  wpsShape: {
    marker: "wpsShape",
    probe: (out) => {
      const shape = (
        out as {
          wpsShape: {
            transformation?: { width?: number };
            children?: (string | { text?: string })[];
          };
        }
      ).wpsShape;
      expect(shape.transformation?.width).toBe(100);
      // the body paragraph collapses to the string shorthand.
      expect(shape.children?.[0]).toBe("box");
    },
  },
  wpgGroup: {
    marker: "wpgGroup",
    probe: (out) =>
      expect(out).toEqual({
        wpgGroup: { children: [], transformation: { width: 100, height: 60 } },
      }),
  },
};

describe("ParagraphChild dispositions", () => {
  for (const [tag, disposition] of Object.entries(PARAGRAPH_CHILD_DISPOSITIONS)) {
    it(`${tag} → ${JSON.stringify(disposition)}`, () => {
      const fixture = INLINE_FIXTURES[tag as keyof typeof INLINE_FIXTURES]();
      const { json, compiled } = roundTripInline(fixture);
      const types = collectTypes(json);
      const out = firstInline(compiled);
      if ("editable" in disposition) {
        const claim = INLINE_EDITABLE[tag as keyof typeof INLINE_EDITABLE];
        expect(
          types.has(claim.marker),
          `resolved doc must contain a ${claim.marker} node/mark`,
        ).toBe(true);
        claim.probe(out);
      } else {
        expect(types.has("inlinePassthrough"), "must ride the inline passthrough atom").toBe(true);
        expect(out).toEqual(fixture);
      }
    });
  }
});

// ── Run children drops ──

describe("run children drops", () => {
  for (const { tag, reason } of RUN_CHILDREN_DROPPED) {
    it(`drops {${tag}} nested in a run (${reason})`, () => {
      const run: ParagraphChild = {
        text: "x",
        children: [{ [tag]: true } as NonNullable<RunOptions["children"]>[number]],
      };
      const { compiled } = roundTrip([{ paragraph: { children: [run] } }]);
      expect(JSON.stringify(compiled)).not.toContain(`"${tag}"`);
    });
  }

  it("keeps rule-owned children nested in a run (pageBreak inside children)", () => {
    const run: ParagraphChild = {
      text: "x",
      children: [{ pageBreak: true } as NonNullable<RunOptions["children"]>[number]],
    };
    const { compiled } = roundTrip([{ paragraph: { children: [run] } }]);
    expect(JSON.stringify(compiled)).toContain("pageBreak");
  });
});

// ── Real-XML auxiliary path ──

/**
 * The fixtures above are hand-written DocumentOptions — they prove the resolve
 * contract but can drift from what office-open's XML parse actually produces
 * (the hyperlink `url` field was read as `link` for exactly that reason). This
 * suite closes the loop through real XML: generate a binary, parse it back
 * with office-open, then run the docen resolve → compile legs on the PARSED
 * options — the shapes are canonical by construction.
 */
describe("real-XML round-trip (generateDocument → parseDocument)", () => {
  function throughXml(children: SectionChild[]) {
    const binary = generateDocumentSync({ sections: [{ children }] });
    const parsed = parseDocument(new Uint8Array(binary as Buffer));
    const json = resolveDocument(parsed, docxExtensions);
    return compileDocument(json, docxExtensions).sections[0].children;
  }

  it("heading paragraph survives real XML as a paragraph with heading attr", () => {
    const binary = generateDocumentSync({
      sections: [
        { children: [{ paragraph: { heading: "Heading2", children: [{ text: "chapter" }] } }] },
      ],
    });
    const json = resolveDocument(parseDocument(new Uint8Array(binary as Buffer)), docxExtensions);
    const block = json.content?.[0] as { type: string; attrs?: { heading?: string } };
    expect(block.type).toBe("paragraph");
    expect(block.attrs?.heading).toBe("Heading2");
    const compiled = compileDocument(json, docxExtensions).sections[0].children;
    const para = compiled[0] as {
      paragraph: { heading?: string; children?: { text?: string }[] };
    };
    // The heading stays a paragraph node (a heading IS a paragraph); the
    // HeadingLevel pStyle rides on the `heading` attr through both legs.
    expect(para.paragraph.heading).toBe("Heading2");
    expect(para.paragraph.children?.[0]?.text).toBe("chapter");
  });

  it("hyperlink survives real XML with the editable link route", () => {
    const compiled = throughXml([
      {
        paragraph: {
          children: [{ hyperlink: { url: "https://example.com", children: [{ text: "link" }] } }],
        },
      },
    ]);
    const hyperlink = firstInline(compiled) as {
      hyperlink: { url?: string; children?: { text?: string }[] };
    };
    expect(hyperlink.hyperlink.url).toBe("https://example.com");
    expect(hyperlink.hyperlink.children?.[0]?.text).toBe("link");
  });

  it("picture survives real XML with the editable image route", () => {
    const compiled = throughXml([
      {
        paragraph: {
          children: [
            {
              picture: {
                type: "png",
                data: PNG_BYTES,
                transformation: { width: 100, height: 60 },
                altText: { name: "probe" },
              },
            },
          ],
        },
      },
    ]);
    const picture = firstInline(compiled) as {
      picture: { type?: string; altText?: { name?: string } };
    };
    expect(picture.picture.type).toBe("png");
    expect(picture.picture.altText?.name).toBe("probe");
  });

  it("oversized picture skips the base64 round-trip via the media registry", () => {
    // Above MEDIA_INLINE_LIMIT the resolve leg must hand out a registered
    // blob: URL (no megabyte base64 string in the attrs) and the compile leg
    // must recover the original bytes from the registry, untouched.
    const oversized = new Uint8Array(256 * 1024 + 1).fill(0x5a);
    const json = resolveDocument(
      parseDocument(
        new Uint8Array(
          generateDocumentSync({
            sections: [
              {
                children: [
                  {
                    paragraph: {
                      children: [
                        {
                          picture: {
                            type: "png",
                            data: oversized,
                            transformation: { width: 100, height: 60 },
                          },
                        },
                      ],
                    },
                  },
                ],
              },
            ],
          }) as Buffer,
        ),
      ),
      docxExtensions,
    );
    const node = (
      json.content?.[0] as {
        content?: { type: string; attrs?: { src?: string } }[];
      }
    )?.content?.[0];
    expect(node?.type).toBe("image");
    expect(node?.attrs?.src?.startsWith("blob:")).toBe(true);

    const compiled = compileDocument(json, docxExtensions).sections[0].children;
    const picture = firstInline(compiled) as { picture: { data?: Uint8Array } };
    expect(picture.picture.data).toStrictEqual(oversized);
  });

  it("table survives real XML with the editable table route", () => {
    const compiled = throughXml([
      {
        table: { rows: [{ cells: [{ children: [{ paragraph: { text: "cell" } }] }] }] },
      },
    ]);
    const table = (
      compiled[0] as {
        table: { rows: { cells: { children: { paragraph: string | { text?: string } }[] }[] }[] };
      }
    ).table;
    expect(table.rows).toHaveLength(1);
    expect(table.rows[0].cells[0].children[0].paragraph).toBe("cell");
  });

  it("tracked-change insertion survives real XML with the editable mark route", () => {
    const compiled = throughXml([
      {
        paragraph: {
          children: [
            {
              insertion: {
                id: 1,
                author: "a",
                date: "2024-01-01T00:00:00Z",
                children: [{ text: "ins" }],
              },
            },
          ],
        },
      },
    ]);
    const insertion = firstInline(compiled) as {
      insertion: { id?: number; author?: string; children?: { text?: string }[] };
    };
    expect(insertion.insertion.author).toBe("a");
    expect(insertion.insertion.children?.[0]?.text).toBe("ins");
  });
});
