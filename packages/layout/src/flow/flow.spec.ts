import { describe, expect, it } from "vitest";

import { fakeFontMetrics, installFakeCanvas } from "../../test/fake-canvas";
import type { LayoutBlock, LayoutParagraph, LayoutTable } from "../layout-doc";
import { TextMeasurer } from "../text/measure";
import { layoutFlow, layoutFlowSections } from "./flow";

installFakeCanvas();
const measurer = new TextMeasurer(fakeFontMetrics);

const latin = { family: "serif", sizePx: 16 };
// 20px lines: exact line height keeps every paragraph arithmetic simple.
const LINE = 20;
const exact20 = { lineHeight: { rule: "exact" as const, px: LINE }, beforePx: 0, afterPx: 0 };

/** A paragraph with exactly `n` lines: text on line 1, then (break + text)
 *  per further line — a trailing break opens no line, so every break here is
 *  followed by text. Hard breaks give exact line counts, immune to font
 *  metrics. widowControl defaults off so plain splits stay arithmetic; the
 *  widow/keep tests turn it on explicitly. */
const para = (n: number, extra: Partial<LayoutParagraph> = {}): LayoutParagraph => ({
  kind: "paragraph",
  inline: [
    { kind: "text", text: "x", style: latin },
    ...Array.from({ length: n - 1 }, () => [
      { kind: "break" as const },
      { kind: "text" as const, text: "x", style: latin },
    ]).flat(),
  ],
  spacing: exact20,
  defaultTextStyle: latin,
  widowControl: false,
  ...extra,
});

const flow = (blocks: LayoutBlock[], contentHeightPx: number) =>
  layoutFlow(blocks, { contentWidthPx: 300, contentHeightPx }, measurer);

const paras = (pages: ReturnType<typeof flow>) =>
  pages.map((p) =>
    p.items.map((i) => {
      if (i.block.kind !== "paragraph") throw new Error(`expected paragraph, got ${i.block.kind}`);
      return i.block;
    }),
  );

describe("layoutFlow", () => {
  it("returns exactly one page for an empty flow", () => {
    expect(flow([], 100)).toHaveLength(1);
    expect(flow([], 100)[0].items).toHaveLength(0);
  });

  it("stacks whole blocks that fit", () => {
    const pages = flow([para(1), para(1), para(1)], 100);
    expect(pages).toHaveLength(1);
    expect(pages[0].items.map((i) => i.block.kind)).toEqual([
      "paragraph",
      "paragraph",
      "paragraph",
    ]);
    expect(pages[0].items[1].yPx).toBe(20);
    expect(pages[0].items[2].yPx).toBe(40);
  });

  it("splits a paragraph at line boundaries across pages", () => {
    // 9 lines of 20px in a 100px page: page 1 gets 5 lines, page 2 the rest.
    const pages = flow([para(9)], 100);
    expect(pages).toHaveLength(2);
    const [head] = paras(pages)[0];
    const [tail] = paras(pages)[1];
    expect(head.lines).toHaveLength(5);
    expect(tail.lines).toHaveLength(4);
    // The tail's lines are re-based to the page top.
    expect(tail.lines[0].yPx).toBe(0);
    expect(tail.heightPx).toBe(80);
  });

  it("keeps the first-line indent on the head page only when a paragraph splits", () => {
    // A 6-line first-line-indented paragraph over a 100px page: the indent
    // rides on the paragraph's first LINE, so the split tail's leading line
    // (mid-paragraph) must carry none — Word does not re-indent page 2.
    const pages = flow([para(6, { indent: { firstLinePx: 24 } })], 100);
    const [head] = paras(pages)[0];
    const [tail] = paras(pages)[1];
    expect(head.lines[0].firstLineIndentPx).toBe(24);
    for (const line of head.lines.slice(1)) expect(line.firstLineIndentPx).toBeUndefined();
    for (const line of tail.lines) expect(line.firstLineIndentPx).toBeUndefined();
  });

  it("collapses spacing between blocks and contains page-edge margins", () => {
    const withMargins = (before: number, after: number): LayoutParagraph =>
      para(1, {
        spacing: { lineHeight: { rule: "exact", px: LINE }, beforePx: before, afterPx: after },
      });
    const pages = flow([withMargins(10, 30), withMargins(5, 40)], 300);
    expect(pages).toHaveLength(1);
    const [, second] = pages[0].items;
    // max(prevAfter=30, before=5) collapses to 30 → y = 10+20+30.
    expect(second.yPx).toBe(60);
  });

  it("keeps widows and orphans whole with widowControl on", () => {
    // Preceding block leaves 40px: a 3-line paragraph would split 2+1 (a
    // widow tail) → the whole paragraph moves to the next page instead.
    const pages = flow([para(3), para(3, { widowControl: true })], 100);
    expect(pages[0].items).toHaveLength(1);
    expect(paras(pages)[0][0].lines).toHaveLength(3);
    expect(paras(pages)[1][0].lines).toHaveLength(3);

    // 4 lines in the same 40px: 2+2 split satisfies both sides.
    const split = flow([para(3), para(4, { widowControl: true })], 100);
    expect(split).toHaveLength(2);
    expect(paras(split)[0]).toHaveLength(2);
    expect(paras(split)[0][1].lines).toHaveLength(2);
    expect(paras(split)[1][0].lines).toHaveLength(2);
  });

  it("never splits a keepLines paragraph that fits a page", () => {
    // 40px left; the 4-line (80px) keepLines paragraph moves whole to page 2.
    const pages = flow([para(3), para(4, { keepLines: true })], 100);
    expect(pages[0].items).toHaveLength(1);
    expect(paras(pages)[1][0].lines).toHaveLength(4);
  });

  it("force-splits a keepLines paragraph taller than a whole page", () => {
    // Progress beats clipping: a block no page can hold splits greedily.
    // 5 lines of 20px in a 40px page → 2 + 2 + 1.
    const pages = flow([para(5, { keepLines: true })], 40);
    expect(pages).toHaveLength(3);
    expect(paras(pages).map((ps) => ps[0].lines.length)).toEqual([2, 2, 1]);
  });

  it("pulls a keepNext heading to the next page with its paragraph", () => {
    // Page: para(2)=40 + heading=20 → 60 of 80; the next 2-line paragraph
    // would orphan its heading, so the heading moves along — BEFORE it.
    const pages = flow(
      [
        para(2, { widowControl: true }),
        para(1, { keepNext: true }),
        para(2, { widowControl: true }),
      ],
      80,
    );
    expect(pages[0].items).toHaveLength(1);
    expect(pages[1].items).toHaveLength(2);
    // The heading precedes the paragraph it keeps with on the next page.
    expect(paras(pages)[1].map((p) => p.keepNext)).toEqual([true, undefined]);
  });

  it("cascades keepNext through a chain of keepers", () => {
    const pages = flow(
      [
        para(2, { widowControl: true }),
        para(1, { keepNext: true }),
        para(1, { keepNext: true }),
        para(2, { widowControl: true }),
      ],
      80,
    );
    expect(pages[0].items).toHaveLength(1);
    expect(pages[1].items).toHaveLength(3);
  });

  it("closes the page at a pageBreak atom (never opens one)", () => {
    const pages = flow([para(1), { kind: "pageBreak" }, para(1)], 300);
    expect(pages).toHaveLength(2);
    expect(pages[0].items).toHaveLength(1);
    expect(pages[1].items).toHaveLength(1);
  });

  it("pushes the body down and shrinks the room under a tall header (pageInsets)", () => {
    // 100px content box; a 30px header push leaves 70px of room: three 20px
    // paragraphs fit, the fourth moves on.
    const insets = { default: { topPx: 30, bottomPx: 0 } };
    const parasIn = [para(1), para(1), para(1), para(1)];
    const pages = layoutFlow(
      parasIn,
      { contentWidthPx: 300, contentHeightPx: 100, pageInsets: insets },
      measurer,
    );
    expect(pages).toHaveLength(2);
    expect(pages[0].items[0].yPx).toBe(30);
    expect(pages[0].items).toHaveLength(3);
    expect(pages[1].items[0].yPx).toBe(30);
  });

  it("pushes the body up from the bottom under a tall footer (pageInsets)", () => {
    // 30px footer push: room 70px — three 20px paragraphs fit, four don't.
    const insets = { default: { topPx: 0, bottomPx: 30 } };
    const pages = layoutFlow(
      [para(1), para(1), para(1), para(1)],
      { contentWidthPx: 300, contentHeightPx: 100, pageInsets: insets },
      measurer,
    );
    expect(pages).toHaveLength(2);
    expect(pages[0].items[0].yPx).toBe(0);
    expect(pages[0].items).toHaveLength(3);
  });

  it("applies first-page insets to page 1 only (title page push)", () => {
    const insets = { default: { topPx: 0, bottomPx: 0 }, first: { topPx: 30, bottomPx: 0 } };
    const pages = layoutFlow(
      [para(1), para(1), para(1), para(1), para(1)],
      { contentWidthPx: 300, contentHeightPx: 100, pageInsets: insets },
      measurer,
    );
    expect(pages[0].items[0].yPx).toBe(30);
    expect(pages[1].items[0].yPx).toBe(0);
  });

  it("honors pageBreakBefore", () => {
    const pages = flow([para(1), para(1, { pageBreakBefore: true })], 300);
    expect(pages).toHaveLength(2);
    expect(pages[0].items).toHaveLength(1);
  });

  it("splits tables at row boundaries", () => {
    const row = (): LayoutTable["rows"][number] => ({
      cells: [{ blocks: [para(1)] }],
    });
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [300],
      rows: [row(), row(), row(), row(), row()],
    };
    // Each row = 20px (one line). Page fits 2.5 rows → 2 + 2 + 1.
    const pages = flow([table], 50);
    expect(pages).toHaveLength(3);
    const counts = pages.map((p) => {
      const t = p.items[0].block;
      if (t.kind !== "table") throw new Error(`expected table, got ${t.kind}`);
      return t.rows.length;
    });
    expect(counts).toEqual([2, 2, 1]);
  });

  it("splits groups at child boundaries", () => {
    const pages = flow([{ kind: "group", blocks: [para(1), para(1), para(1)] }], 50);
    expect(pages).toHaveLength(2);
    const groups = pages.map((p) => {
      const g = p.items[0].block;
      if (g.kind !== "group") throw new Error(`expected group, got ${g.kind}`);
      return g;
    });
    expect(groups[0].children).toHaveLength(2);
    expect(groups[1].children).toHaveLength(1);
    expect(groups[1].children[0].block.kind).toBe("paragraph");
  });

  it("splits a lone 2-line paragraph 1+1 when widowControl is off", () => {
    const pages = flow([para(2)], 20);
    expect(pages).toHaveLength(2);
    expect(paras(pages)[0][0].lines).toHaveLength(1);
    expect(paras(pages)[1][0].lines).toHaveLength(1);
  });

  it("moves a placeholder whole — never splits, keeps its estimated height", () => {
    const pages = flow([para(1), { kind: "placeholder", heightPx: 30, label: "toc" }], 25);
    // 20px paragraph fills page 1; the 30px placeholder cannot split, so it
    // moves whole to page 2.
    expect(pages).toHaveLength(2);
    expect(pages[0].items).toHaveLength(1);
    const moved = pages[1].items[0].block;
    expect(moved).toEqual({ kind: "placeholder", heightPx: 30, label: "toc" });
  });

  it("overflows a placeholder taller than a page instead of looping", () => {
    const pages = flow([{ kind: "placeholder", heightPx: 500 }], 100);
    expect(pages).toHaveLength(1);
    expect(pages[0].items).toHaveLength(1);
  });
});

/** A floating drawing box on a paragraph: column-anchored at (x, y) with the
 *  given wrap mode — the members stay empty, the flow only reads geometry. */
const drawing = (
  x: number,
  y: number,
  width: number,
  height: number,
  wrap: "square" | "tight" | "topAndBottom" | undefined,
): NonNullable<LayoutParagraph["drawings"]>[number] => ({
  anchor: {
    horizontal: { relative: "column", offsetPx: x },
    vertical: { relative: "paragraph", offsetPx: y },
  },
  width,
  height,
  members: [],
  wrap,
});

/** A wrapping-text paragraph: one long run of latin atoms (8px each under
 *  the fake metrics) so the line count follows the usable width. */
const wrapPara = (chars: number, extra: Partial<LayoutParagraph> = {}): LayoutParagraph => ({
  kind: "paragraph",
  inline: [{ kind: "text", text: "x".repeat(chars), style: latin }],
  spacing: exact20,
  defaultTextStyle: latin,
  widowControl: false,
  ...extra,
});

describe("layoutFlow float wraps", () => {
  it("shrinks the lines a square drawing overlaps (text wraps beside it)", () => {
    // 24 atoms = 192px: one line at the full 300px width, two beside a
    // 200px box (100px usable = 12 atoms per line). The drawing is 60px tall
    // so its zone reaches the second paragraph's first line (the anchor's
    // own two lines already end at y 40).
    const pages = flow(
      [wrapPara(24, { drawings: [drawing(0, 0, 200, 60, "square")] }), wrapPara(24)],
      300,
    );
    const [, body] = paras(pages)[0];
    expect(body.lines).toHaveLength(2);
    expect(body.heightPx).toBe(40);
  });

  it("wraps the anchor paragraph's own lines beside its square float", () => {
    // The drawings' offsets are paragraph-relative, so the anchor wraps its
    // own 24 atoms beside the 200px box — two 12-atom lines, not one.
    const pages = flow([wrapPara(24, { drawings: [drawing(0, 0, 200, 40, "square")] })], 300);
    const [anchor] = paras(pages)[0];
    expect(anchor.lines).toHaveLength(2);
    expect(anchor.heightPx).toBe(40);
  });

  it("keeps a float below the anchor paragraph's first lines out of them", () => {
    // The box starts 60px into the paragraph: three full-width lines (37
    // atoms each, exact 20px rows) pack above it, the fourth line — whose
    // top reaches the box — wraps beside it (100px = 12 atoms).
    const pages = flow([wrapPara(112, { drawings: [drawing(0, 60, 200, 60, "square")] })], 300);
    const [anchor] = paras(pages)[0];
    expect(anchor.lines).toHaveLength(4);
    expect(anchor.lines[2]!.maxWidthPx).toBe(300);
    expect(anchor.lines[3]!.maxWidthPx).toBe(100);
  });

  it("wraps a later line a zone's top opens INSIDE of (box overlaps the line box)", () => {
    // The box starts at y 30 — inside the NEXT paragraph's first line
    // (20-40). The line box overlaps the box, so that line already wraps
    // beside it (12-atom rows); a top-of-line-only check would let the text
    // run straight under the float.
    const pages = flow(
      [para(1, { drawings: [drawing(0, 30, 200, 60, "square")] }), wrapPara(24)],
      300,
    );
    const [, body] = paras(pages)[0];
    expect(body.lines).toHaveLength(2);
    expect(body.lines[0]!.maxWidthPx).toBe(100);
  });

  it("wraps the anchor's first line when its float starts inside that line", () => {
    // Same mid-line start, self-zone flavor: the box hangs 10px into the
    // anchor's own first line, so line 1 already packs beside it.
    const pages = flow([wrapPara(24, { drawings: [drawing(0, 10, 200, 40, "square")] })], 300);
    const [anchor] = paras(pages)[0];
    expect(anchor.lines).toHaveLength(2);
    expect(anchor.lines[0]!.maxWidthPx).toBe(100);
  });

  it("clears the band of a topAndBottom drawing (next block resumes below)", () => {
    // Anchor line at y 0-20; the band [20, 70] pushes the next paragraph
    // from y 20 down to y 70.
    const pages = flow(
      [para(1, { drawings: [drawing(0, 20, 200, 50, "topAndBottom")] }), para(1)],
      300,
    );
    expect(pages).toHaveLength(1);
    expect(pages[0].items[1].yPx).toBe(70);
  });

  it("treats a square box covering the whole column width as a cleared band", () => {
    const pages = flow(
      [para(1, { drawings: [drawing(-100, 0, 500, 40, "square")] }), para(1)],
      300,
    );
    expect(pages[0].items[1].yPx).toBe(40);
  });

  it("splits a long paragraph at a band top and resumes below the band", () => {
    // Page 300px: a 2-line paragraph (y 0-40), an anchor line at y 40 whose
    // drawing clears the band [100, 140] (offset 60 from its top — the band
    // opens below the anchor line), then a 5-line paragraph starting at
    // y 60: two lines fit above the band (60-100), the remaining three
    // resume at 140.
    const pages = flow(
      [para(2), para(1, { drawings: [drawing(0, 60, 500, 40, "topAndBottom")] }), para(5)],
      300,
    );
    expect(pages).toHaveLength(1);
    const laid = paras(pages)[0];
    expect(laid).toHaveLength(4);
    expect(laid[2].lines).toHaveLength(2);
    expect(laid[2].heightPx).toBe(40);
    expect(pages[0].items[3].yPx).toBe(140);
    expect(laid[3].lines).toHaveLength(3);
  });

  it("keeps a wrapNone drawing out of the flow entirely", () => {
    const pages = flow(
      [para(1, { drawings: [drawing(0, 0, 200, 40, undefined)] }), wrapPara(24)],
      300,
    );
    expect(pages[0].items[1].yPx).toBe(20);
    expect(paras(pages)[0][1].lines).toHaveLength(1);
  });

  it("forgets float zones at a page break (floats never cross pages)", () => {
    // The anchor paragraph sits at the page bottom (y 80-100) and its zone
    // [80, 160] reaches past the page edge; the wrapping paragraph lands on
    // page 2 with the full width back.
    const pages = flow(
      [para(4), wrapPara(24, { drawings: [drawing(0, 0, 200, 80, "square")] }), wrapPara(24)],
      100,
    );
    expect(pages).toHaveLength(2);
    expect(paras(pages)[1][0].lines).toHaveLength(1);
  });

  it("drops a split tail's drawings (the float paints on its anchor page)", () => {
    // The anchor paragraph splits 4+1 across an 80px page: the head keeps
    // the drawing, the page-2 tail paints nothing.
    const pages = flow([para(5, { drawings: [drawing(0, 0, 200, 40, "square")] })], 80);
    const [head, tail] = [paras(pages)[0][0], paras(pages)[1][0]];
    expect(head.drawings).toHaveLength(1);
    expect(tail.drawings).toBeUndefined();
  });
});

describe("layoutFlowSections", () => {
  const opts = (
    contentHeightPx: number,
    pageInsets?: Parameters<typeof layoutFlow>[1]["pageInsets"],
  ) => ({
    contentWidthPx: 300,
    contentHeightPx,
    ...(pageInsets ? { pageInsets } : {}),
  });

  it("starts each section on a fresh page and maps pages to sections", () => {
    // Section 0 fills 2 pages (6 paragraphs of 20px into a 100px page: 5 fit,
    // the 6th overflows), section 1 adds a third page.
    const run = layoutFlowSections(
      [
        {
          blocks: [para(1), para(1), para(1), para(1), para(1), para(1)],
          opts: opts(100),
        },
        { blocks: [para(2)], opts: opts(100) },
      ],
      measurer,
    );
    expect(run.pages).toHaveLength(3);
    expect(run.sectionOfPage).toEqual([0, 0, 1]);
  });

  it("keys the even inset slot off the PHYSICAL page across sections", () => {
    // Section 0 has 3 paragraphs of 20px on an 80px page — all 3 lines DO fit
    // (60 <= 80), so it is one page; the even slot of the DOCUMENT's page 2
    // (the second section's first page) must apply there. With a 30px top
    // push on even pages, the second section's first paragraph starts at y=30.
    const even = { topPx: 30, bottomPx: 0 };
    const run = layoutFlowSections(
      [
        { blocks: [para(3)], opts: opts(100, { default: { topPx: 0, bottomPx: 0 }, even }) },
        { blocks: [para(1)], opts: opts(100, { default: { topPx: 0, bottomPx: 0 }, even }) },
      ],
      measurer,
    );
    expect(run.pages).toHaveLength(2);
    expect(run.sectionOfPage).toEqual([0, 1]);
    // Global page 1 (physical page 2, odd 0-based index) → even slot.
    expect(run.pages[1].items[0].yPx).toBe(30);
    // Global page 0 (physical page 1) → default slot.
    expect(run.pages[0].items[0].yPx).toBe(0);
  });

  it("keeps the first-page inset local to each section", () => {
    // Both sections carry a first-page top push — each section's own first
    // page gets it, whichever global page it lands on.
    const first = { topPx: 30, bottomPx: 0 };
    const run = layoutFlowSections(
      [
        { blocks: [para(1)], opts: opts(100, { default: { topPx: 0, bottomPx: 0 }, first }) },
        { blocks: [para(1)], opts: opts(100, { default: { topPx: 0, bottomPx: 0 }, first }) },
      ],
      measurer,
    );
    expect(run.pages).toHaveLength(2);
    expect(run.pages[0].items[0].yPx).toBe(30);
    expect(run.pages[1].items[0].yPx).toBe(30);
  });
});
