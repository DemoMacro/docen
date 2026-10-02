import type {
  FlowPage,
  LaidOutParagraph,
  LayoutInline,
  ProjectedPageNumbering,
} from "@docen/layout";
import { describe, expect, it } from "vitest";

import {
  resolvePageFields,
  resolvePageFieldsBounded,
  type PageFieldContext,
} from "./field-resolve";
import { evaluateField, instructionName, LIVE_FIELD_NAMES } from "./fields";

// The pagination-feedback walk: field atoms resolve against the page they
// landed on. The resolver mirrors the orchestrator's live policy — only the
// numbering fields resolve at render; everything else keeps its cache.

const style = { family: "serif", sizePx: 16 };

const fieldAtom = (
  instruction: string,
  extra: Partial<{ field: "page" | "numPages"; result: string; text: string }> = {},
): LayoutInline => ({
  kind: "text",
  text: extra.text ?? "0",
  style,
  instruction,
  ...extra,
});

const cachedAtom = (text: string, instruction: string): LayoutInline => ({
  kind: "text",
  text,
  style,
  instruction,
  result: text,
});

const para = (inline: LayoutInline[]): LaidOutParagraph => ({
  kind: "paragraph",
  heightPx: 20,
  beforePx: 0,
  afterPx: 0,
  inline,
  lines: [
    {
      yPx: 0,
      heightPx: 20,
      naturalPx: 20,
      endInlineIndex: inline.length - 1,
      items: inline.map((atom, index) => ({
        kind: "text" as const,
        inlineIndex: index,
        text: atom.kind === "text" ? atom.text : "",
        xPx: index * 10,
        widthPx: 10,
      })),
    },
  ],
});

const pageOf = (children: LayoutInline[][]): FlowPage => ({
  items: children.map((inline, index) => ({ yPx: index * 20, block: para(inline) })),
});

/** The resolver the editor wires: the walk's page context becomes the frame. */
const resolve = (instruction: string, page: PageFieldContext): string | null => {
  if (!LIVE_FIELD_NAMES.has(instructionName(instruction))) return null;
  return evaluateField(instruction, {
    frame: {
      page: page.pageNumber,
      pageCount: page.pageCount,
      section: page.section,
      sectionPages: page.sectionPages,
      ...(page.pageFormat ? { pageFormat: page.pageFormat } : {}),
    },
  });
};

const PAGE = (result = "1") => fieldAtom("PAGE", { field: "page", result });
const NUMPAGES = (result = "1") => fieldAtom("NUMPAGES", { field: "numPages", result });

type AtomView = { kind?: string; text?: string; resolved?: string; field?: string };

const atomsOf = (page: FlowPage, index = 0): readonly AtomView[] => {
  const block = page.items[index]!.block;
  if (block.kind !== "paragraph") throw new Error("expected paragraph");
  return block.inline as readonly AtomView[];
};

describe("resolvePageFields", () => {
  it("resolves PAGE/NUMPAGES to the real numbers in a multi-page document", () => {
    const pages = [
      pageOf([[PAGE(), NUMPAGES(), cachedAtom("作者甲", "AUTHOR")]]),
      pageOf([[PAGE(), NUMPAGES()]]),
      pageOf([[PAGE(), NUMPAGES()]]),
    ];
    const outcome = resolvePageFields(pages, [{}], [0, 0, 0], resolve);
    expect(atomsOf(pages[0]!)[0]).toMatchObject({ text: "0", resolved: "1", field: "page" });
    expect(atomsOf(pages[1]!)[0]).toMatchObject({ text: "0", resolved: "2" });
    expect(atomsOf(pages[2]!)[0]).toMatchObject({ text: "0", resolved: "3" });
    expect(atomsOf(pages[0]!)[1]).toMatchObject({ text: "0", resolved: "3" });
    expect(atomsOf(pages[2]!)[1]).toMatchObject({ text: "0", resolved: "3" });
    // Dynamic page numbers keep the measuring placeholder — only `resolved`
    // changes, so no re-layout is requested.
    expect(outcome.textChanged).toBe(false);
    expect(outcome.changedPages).toEqual([0, 1, 2]);
    // A non-live field keeps its cache untouched.
    expect(atomsOf(pages[0]!)[2].resolved).toBeUndefined();
    expect(atomsOf(pages[0]!)[2].text).toBe("作者甲");
  });

  it("rewrites measured text for non-numbering live fields (SECTION)", () => {
    const pages = [pageOf([[fieldAtom("SECTION")]]), pageOf([[fieldAtom("SECTION")]])];
    const outcome = resolvePageFields(pages, [{}, {}], [0, 1], resolve);
    expect(atomsOf(pages[0]!)[0]).toMatchObject({ text: "1", resolved: "1" });
    expect(atomsOf(pages[1]!)[0]).toMatchObject({ text: "2", resolved: "2" });
    expect(outcome.textChanged).toBe(true);
  });

  it("applies each section's page-number restart and numFmt", () => {
    const sections: { pageNumbering?: ProjectedPageNumbering }[] = [
      { pageNumbering: { start: 5, format: "lowerRoman" } },
      {},
    ];
    const pages = [
      pageOf([[PAGE(), NUMPAGES(), fieldAtom("SECTION"), fieldAtom("SECTIONPAGES")]]),
      pageOf([[PAGE()]]),
      pageOf([[PAGE(), NUMPAGES(), fieldAtom("SECTION")]]),
      pageOf([[PAGE()]]),
    ];
    resolvePageFields(pages, sections, [0, 0, 1, 1], resolve);
    const numbersAt = (page: number) =>
      atomsOf(pages[page]!).map((atom) => (atom.kind === "text" ? atom.resolved : undefined));
    expect(numbersAt(0)).toEqual(["v", "4", "i", "ii"]);
    expect(numbersAt(1)).toEqual(["vi"]);
    // Section 1 has no w:numFmt of its own — decimals continue the count.
    expect(numbersAt(2)).toEqual(["7", "4", "2"]);
    expect(numbersAt(3)).toEqual(["8"]);
  });

  it("walks table cells and footnote areas", () => {
    const cellPara = para([PAGE()]);
    const table = {
      kind: "table" as const,
      widthPx: 100,
      columnWidthsPx: [100],
      heightPx: 20,
      rows: [
        {
          heightPx: 20,
          cells: [
            {
              colspan: 1,
              rowspan: 1,
              insets: { top: 0, right: 0, bottom: 0, left: 0 },
              innerWidthPx: 100,
              stack: [{ yPx: 0, block: cellPara }],
            },
          ],
        },
      ],
    };
    const notePara = para([NUMPAGES()]);
    const pages: FlowPage[] = [
      {
        items: [{ yPx: 0, block: table }],
        footnotes: {
          yPx: 200,
          separatorWidthPx: 192,
          totalHeightPx: 20,
          items: [{ yPx: 0, block: notePara }],
          notes: [{ id: 1, ordinal: 1, stack: [{ yPx: 0, block: notePara }], heightPx: 20 }],
        },
      },
    ];
    resolvePageFields(pages, [{}], [0], resolve);
    expect(cellPara.inline[0]).toMatchObject({ resolved: "1" });
    expect(notePara.inline[0]).toMatchObject({ resolved: "1" });
  });

  it("resolves a shared inline array once per page", () => {
    let calls = 0;
    const pages = [pageOf([[PAGE()]])];
    resolvePageFields(pages, [{}], [0], (instruction, page) => {
      calls += 1;
      return resolve(instruction, page);
    });
    expect(calls).toBe(1);
  });
});

describe("resolvePageFieldsBounded", () => {
  it("stops after one pass when numbering placeholders keep the layout stable", () => {
    const pages = [pageOf([[PAGE(), NUMPAGES()]])];
    let relayouts = 0;
    const result = resolvePageFieldsBounded(pages, [{}], [0], resolve, () => {
      relayouts += 1;
      return { pages: [pageOf([[PAGE(), NUMPAGES()]])], sectionOfPage: [0] };
    });
    expect(result.passes).toBe(1);
    expect(relayouts).toBe(0);
    expect(atomsOf(result.pages[0]!)[0]).toMatchObject({ resolved: "1" });
    expect(result.dirty).toEqual([0]);
  });

  it("re-lays while a value rewrites measured text, then stops when stable", () => {
    // A SECTION-like field whose value the first pass writes into the
    // measured text. The re-layout re-lays the SAME projected atom (the real
    // pipeline's rewritten inline), so the second pass resolves the same
    // value — two passes, one re-layout.
    const atom = fieldAtom("SECTION");
    let pages = [pageOf([[atom]])];
    let relayouts = 0;
    const result = resolvePageFieldsBounded(pages, [{}], [0], resolve, () => {
      relayouts += 1;
      pages = [pageOf([[atom]])];
      return { pages, sectionOfPage: [0] };
    });
    expect(result.passes).toBe(2);
    expect(relayouts).toBe(1);
    expect(atomsOf(result.pages[0]!)[0]).toMatchObject({ text: "1", resolved: "1" });
  });

  it("terminates at the pass cap when a value never stabilizes", () => {
    let flip = 0;
    const atom = fieldAtom("SECTION");
    let pages = [pageOf([[atom]])];
    let relayouts = 0;
    const result = resolvePageFieldsBounded(
      pages,
      [{}],
      [0],
      () => String(++flip),
      () => {
        relayouts += 1;
        pages = [pageOf([[atom]])];
        return { pages, sectionOfPage: [0] };
      },
      3,
    );
    expect(result.passes).toBe(3);
    expect(relayouts).toBe(2);
    expect(atomsOf(result.pages[0]!)[0]).toMatchObject({ resolved: "3" });
  });
});
