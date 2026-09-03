import { describe, expect, it } from "vitest";

import { fakeFontMetrics, installFakeCanvas } from "../../test/fake-canvas";
import type { LayoutBlock, LayoutParagraph, LayoutTable, LayoutTableCell } from "../layout-doc";
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
    // The break's own row stays on the page it ends (Word's "····分页符····"
    // line), the closing paragraph opens the next page.
    expect(pages[0].items).toHaveLength(2);
    expect(pages[0].items[1]!.block.kind).toBe("pageBreak");
    expect(pages[1].items).toHaveLength(1);
  });

  it("collapses a pageBreak that cannot fit the page's last line (zero height)", () => {
    // 80px of content in a 95px page leaves 15px — short of the break row's
    // 20px line. The row collapses at the page bottom instead of spilling
    // past the page edge; the page still closes for the next paragraph.
    const pages = flow([para(4), { kind: "pageBreak" }, para(1)], 95);
    expect(pages).toHaveLength(2);
    const brk = pages[0].items[1]!;
    expect(brk.block.kind).toBe("pageBreak");
    expect(brk.block.heightPx).toBe(0);
    expect(brk.yPx).toBe(80);
  });

  it("collapses a section-break paragraph that cannot fit, never blanking a page", () => {
    // The section's last paragraph is a lone marker row: dropped onto a fresh
    // page it would leave that page blank under the old section (Word's
    // undeletable blank page). It collapses at the page bottom instead.
    const pages = flow([para(4), para(1, { sectionEnd: true })], 95);
    expect(pages).toHaveLength(1);
    const [, last] = pages[0].items;
    expect(last!.block.kind).toBe("paragraph");
    if (last!.block.kind === "paragraph") expect(last!.block.sectionEnd).toBe(true);
    expect(last!.block.heightPx).toBe(0);
    expect(last!.yPx).toBe(80);
  });

  it("still splits a multi-line section-break paragraph that cannot fit", () => {
    // Content paragraphs split at line boundaries as always — only the lone
    // marker row collapses.
    const pages = flow([para(4), para(3, { sectionEnd: true })], 95);
    expect(pages).toHaveLength(2);
    expect(paras(pages)[0]).toHaveLength(1);
    const [tail] = paras(pages)[1];
    expect(tail.lines).toHaveLength(3);
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

  it("re-opens row spans a page cut crosses so tail rows keep their columns", () => {
    // Row 0's first cell spans all three rows (a vertical merge): the cut
    // leaves its anchor behind, so the tail rows must re-open an empty
    // placeholder in its columns — Word re-opens the merge on every page —
    // otherwise the tail's cells pack one column left and the first column
    // appears to vanish (its content renders where column 2 should be).
    const cell = (over: Partial<LayoutTableCell> = {}): LayoutTableCell => ({
      blocks: [para(1)],
      ...over,
    });
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [100, 100, 100],
      rows: [
        { cells: [cell({ rowspan: 3 }), cell({ colspan: 2 })] },
        { cells: [cell({ colspan: 2 })] },
        { cells: [cell({ colspan: 2, blocks: [para(2)] })] },
      ],
    };
    // Row heights 20/20/40 in a 50px page: the cut lands after row 1.
    const pages = flow([table], 50);
    expect(pages).toHaveLength(2);
    const tail = pages[1].items[0].block;
    if (tail.kind !== "table") throw new Error(`expected table, got ${tail.kind}`);
    expect(tail.rows).toHaveLength(1);
    // The re-opened placeholder owns column 0 (empty, no content stack); the
    // content cell follows at column 1.
    expect(tail.rows[0]!.cells).toHaveLength(2);
    expect(tail.rows[0]!.cells[0]!.colspan).toBe(1);
    expect(tail.rows[0]!.cells[0]!.stack).toHaveLength(0);
    expect(tail.rows[0]!.cells[1]!.colspan).toBe(2);
  });

  it("re-opens row spans across a mid-row force-split too", () => {
    // Row 0's first cell spans rows 0-1; row 1 is taller than the whole page
    // (force-split mid-content). The tail's first row is row 1's lower half —
    // the merge anchor above still owns column 0 there.
    const cell = (over: Partial<LayoutTableCell> = {}): LayoutTableCell => ({
      blocks: [para(1)],
      ...over,
    });
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [100, 100, 100],
      rows: [
        { cells: [cell({ rowspan: 2 }), cell({ colspan: 2 })] },
        { cells: [cell({ colspan: 2, blocks: [para(4)] })] },
      ],
    };
    // Row 0 = 20px, row 1 = 80px in a 50px page → row 1 force-splits at 30px.
    const pages = flow([table], 50);
    expect(pages).toHaveLength(2);
    const tail = pages[1].items[0].block;
    if (tail.kind !== "table") throw new Error(`expected table, got ${tail.kind}`);
    expect(tail.rows[0]!.cells[0]!.colspan).toBe(1);
    expect(tail.rows[0]!.cells[0]!.stack).toHaveLength(0);
    expect(tail.rows[0]!.cells[1]!.colspan).toBe(2);
  });

  it("repeats the tblHeader row on every continuation page", () => {
    const row = (header = false): LayoutTable["rows"][number] => ({
      cells: [{ blocks: [para(1)] }],
      ...(header ? { tableHeader: true } : {}),
    });
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [300],
      rows: [row(true), row(), row(), row(), row()],
    };
    // 20px rows in a 50px page: a cut must keep the band + a body row, so the
    // header rides along and re-opens EVERY continuation — the copy keeps its
    // mark (Word repeats it again on a further page) while eating one row's
    // worth of space per page (4 body rows → 4 band+body pages).
    const pages = flow([table], 50);
    expect(pages).toHaveLength(4);
    const tables = pages.map((p) => {
      const t = p.items[0].block;
      if (t.kind !== "table") throw new Error(`expected table, got ${t.kind}`);
      return t;
    });
    expect(tables.map((t) => t.rows.length)).toEqual([2, 2, 2, 2]);
    for (const page of tables) {
      expect(page.rows[0].tableHeader).toBe(true);
      expect(page.rows.slice(1).every((r) => !r.tableHeader)).toBe(true);
    }
  });

  it("treats only the contiguous tblHeader prefix as the repeat band", () => {
    const row = (header = false): LayoutTable["rows"][number] => ({
      cells: [{ blocks: [para(1)] }],
      ...(header ? { tableHeader: true } : {}),
    });
    // A mark that doesn't start at the top row is not part of the band (Word
    // ignores mid-table tblHeader rows).
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [300],
      rows: [row(true), row(true), row(), row(true), row()],
    };
    // Band = the 2-row prefix (40px). A 60px page fits band + 1 body row.
    const pages = flow([table], 60);
    const tables = pages.map((p) => {
      const t = p.items[0].block;
      if (t.kind !== "table") throw new Error(`expected table, got ${t.kind}`);
      return t;
    });
    // Page 1 [band + body1]; page 2 [copies + the mid marked row as a plain
    // body row]; page 3 [copies + body3]. The copies keep the band alive on
    // every page, but the mid-table mark is stripped — it never widens the
    // band or opens one of its own.
    expect(tables.map((t) => t.rows.length)).toEqual([3, 3, 3]);
    for (const page of tables.slice(1)) {
      expect(page.rows[0].tableHeader).toBe(true);
      expect(page.rows[1].tableHeader).toBe(true);
    }
    expect(tables[1].rows[2].tableHeader).toBeUndefined();
    expect(tables[2].rows[2].tableHeader).toBeUndefined();
  });

  it("moves the table whole when the band + first body row don't fit", () => {
    const row = (header = false): LayoutTable["rows"][number] => ({
      cells: [{ blocks: [para(1)] }],
      ...(header ? { tableHeader: true } : {}),
    });
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [300],
      rows: [row(true), row(), row(), row()],
    };
    // The 40px paragraph leaves 10px: not enough for band (20) + a body row —
    // the table moves whole to page 2 instead of splitting off a band-less
    // fragment, then splits normally (band + 1 body row per page, the band
    // repeated on every one).
    const pages = flow([para(2), table], 50);
    expect(pages[0].items.map((i) => i.block.kind)).toEqual(["paragraph"]);
    expect(pages).toHaveLength(4);
    const tables = pages.slice(1).map((p) => {
      const t = p.items[0].block;
      if (t.kind !== "table") throw new Error(`expected table, got ${t.kind}`);
      return t;
    });
    expect(tables.map((t) => t.rows.length)).toEqual([2, 2, 2]);
    expect(tables[0].rows[0].tableHeader).toBe(true);
    expect(tables[1].rows[0].tableHeader).toBe(true);
  });

  it("gives up repeating when the band is taller than the page body area", () => {
    // Band = 2 × 40px rows = 80px; the page body is 60px — Word's anti-loop
    // rule drops the repetition and the table splits as if unmarked.
    const tall = (header = false): LayoutTable["rows"][number] => ({
      cells: [{ blocks: [para(2)] }],
      ...(header ? { tableHeader: true } : {}),
    });
    const tallTable: LayoutTable = {
      kind: "table",
      columnWidthsPx: [300],
      rows: [tall(true), tall(true), tall(), tall()],
    };
    const pages = flow([tallTable], 60);
    const tables = pages.map((p) => {
      const t = p.items[0].block;
      if (t.kind !== "table") throw new Error(`expected table, got ${t.kind}`);
      return t;
    });
    // Plain split, no copies anywhere: every 40px row fits the 60px page
    // whole, so each page holds one row and none ever splits mid-content
    // (Word only force-splits rows no page could hold). Every tail row was
    // stripped, so no later page re-derives a band from a mid-table marked
    // row.
    expect(tables.map((t) => t.rows.length)).toEqual([1, 1, 1, 1]);
    expect(tables[0].rows[0].tableHeader).toBe(true);
    for (const t of tables.slice(1)) expect(t.rows.every((r) => !r.tableHeader)).toBe(true);
  });

  it("splits a tall row mid-content when the page runs out (mid-row split)", () => {
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [300],
      rows: [
        { cells: [{ blocks: [para(1)] }] },
        { cells: [{ blocks: [para(1)] }] },
        { cells: [{ blocks: [para(6)] }] }, // 120px — taller than the 100px page
        { cells: [{ blocks: [para(1)] }] },
      ],
    };
    // 40px of paragraphs + 20 + 20 fills 80; the over-page 120px row can't
    // fit the last 20px — it splits there: 1 line stays in the head half, 5
    // re-open on page 2 above the trailing row (the tail rebased to its own
    // top). A row that fits a fresh page moves whole instead — Word never
    // splits pageable rows (COM-verified).
    const pages = flow([para(2), table], 100);
    expect(pages).toHaveLength(3);
    const tables = pages.map((p) => {
      const t = p.items.find((i) => i.block.kind === "table")?.block;
      if (t?.kind !== "table") throw new Error("expected table");
      return t;
    });
    const [head, tail] = tables;
    expect(head.rows.map((r) => r.heightPx)).toEqual([20, 20, 20]);
    const headPara = head.rows[2].cells[0].stack[0].block;
    if (headPara.kind !== "paragraph") throw new Error("expected paragraph");
    expect(headPara.lines).toHaveLength(1);
    expect(tail.rows.map((r) => r.heightPx)).toEqual([100]);
    expect(tables[2].rows.map((r) => r.heightPx)).toEqual([20]);
    const tailCell = tail.rows[0].cells[0];
    const tailPara = tailCell.stack[0].block;
    if (tailPara.kind !== "paragraph") throw new Error("expected paragraph");
    expect(tailPara.lines).toHaveLength(5);
    expect(tailPara.lines[0].yPx).toBe(0);
    // The continuation is a fresh row slice — the head's vAlign slack is void.
    expect(tailCell.contentOffsetYPx).toBeUndefined();
  });

  it("moves a pageable tall row whole — Word never splits it", () => {
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [300],
      rows: [
        { cells: [{ blocks: [para(1)] }] },
        { cells: [{ blocks: [para(1)] }] },
        { cells: [{ blocks: [para(4)] }] }, // 80px — fits the 100px page whole
        { cells: [{ blocks: [para(1)] }] },
      ],
    };
    // 40px of paragraphs + 20 + 20 fills 80; the 80px row can't fit the last
    // 20px, but a fresh page holds it whole — Word moves it (COM-verified:
    // wrapped/multi-paragraph/pageable rows all refuse the mid cut).
    const pages = flow([para(2), table], 100);
    expect(pages).toHaveLength(2);
    const [head, tail] = pages.map((p) => {
      const t = p.items.find((i) => i.block.kind === "table")?.block;
      if (t?.kind !== "table") throw new Error("expected table");
      return t;
    });
    expect(head.rows.map((r) => r.heightPx)).toEqual([20, 20]);
    expect(tail.rows.map((r) => r.heightPx)).toEqual([80, 20]);
  });

  it("moves a cantSplit row whole instead of splitting it mid-content", () => {
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [300],
      rows: [
        { cells: [{ blocks: [para(1)] }] },
        { cells: [{ blocks: [para(4)] }], cantSplit: true }, // 80px, won't split
      ],
    };
    // 40 + 20 = 60 fills; the 80px cantSplit row leaves the head as a row-
    // boundary slice and re-opens whole on page 2.
    const pages = flow([para(2), table], 100);
    expect(pages).toHaveLength(2);
    expect(pages[0].items.map((i) => i.block.kind)).toEqual(["paragraph", "table"]);
    const t = pages[1].items[0].block;
    if (t.kind !== "table") throw new Error("expected table");
    expect(t.rows.map((r) => r.heightPx)).toEqual([80]);
    const moved = t.rows[0].cells[0].stack[0].block;
    if (moved.kind !== "paragraph") throw new Error("expected paragraph");
    expect(moved.lines).toHaveLength(4);
  });

  it("force-splits a row taller than a whole page (progress beats clipping)", () => {
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [300],
      rows: [{ cells: [{ blocks: [para(5)] }] }], // 100px into a 60px page
    };
    const pages = flow([table], 60);
    expect(pages).toHaveLength(2);
    const [head, tail] = pages.map((p) => {
      const b = p.items[0].block;
      if (b.kind !== "table") throw new Error("expected table");
      return b;
    });
    // 3 lines fit the 60px page; the remaining 2 re-open on page 2.
    expect(head.rows).toHaveLength(1);
    expect(head.rows[0].heightPx).toBe(60);
    const headPara = head.rows[0].cells[0].stack[0].block;
    if (headPara.kind !== "paragraph") throw new Error("expected paragraph");
    expect(headPara.lines).toHaveLength(3);
    expect(tail.rows[0].heightPx).toBe(40);
    const tailPara = tail.rows[0].cells[0].stack[0].block;
    if (tailPara.kind !== "paragraph") throw new Error("expected paragraph");
    expect(tailPara.lines).toHaveLength(2);
  });

  it("never splits an exact-height row — it moves whole instead", () => {
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [300],
      rows: [{ cells: [{ blocks: [para(4)] }], height: { rule: "exact", px: 80 } }],
    };
    // 40px of paragraphs leave 60; the exact 80px row cannot split mid-content
    // (its height is fixed — overflow clips), so it moves whole.
    const pages = flow([para(2), table], 100);
    expect(pages).toHaveLength(2);
    expect(pages[0].items.map((i) => i.block.kind)).toEqual(["paragraph"]);
    const t = pages[1].items[0].block;
    if (t.kind !== "table") throw new Error("expected table");
    expect(t.rows).toHaveLength(1);
    expect(t.rows[0].heightPx).toBe(80);
    const cellPara = t.rows[0].cells[0].stack[0].block;
    if (cellPara.kind !== "paragraph") throw new Error("expected paragraph");
    expect(cellPara.lines).toHaveLength(4);
  });

  it("repeats the header band when a mid-row split continues the table", () => {
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [300],
      rows: [
        { cells: [{ blocks: [para(1)] }], tableHeader: true },
        { cells: [{ blocks: [para(1)] }] },
        { cells: [{ blocks: [para(5)] }] }, // 100px — over the band's 80px room
        { cells: [{ blocks: [para(1)] }] },
      ],
    };
    // Page 1 fits band(20) + body1(20) then 60px of the over-room 100px row;
    // page 2 re-opens with a fresh band copy + the row's lower 40px + the
    // trailing row. A row the band's fresh-page room (100 − 20) holds whole
    // would move instead — Word only force-splits what no page could hold.
    const pages = flow([table], 100);
    expect(pages).toHaveLength(2);
    const [head, tail] = pages.map((p) => {
      const t = p.items.find((i) => i.block.kind === "table")?.block;
      if (t?.kind !== "table") throw new Error("expected table");
      return t;
    });
    expect(head.rows.map((r) => r.heightPx)).toEqual([20, 20, 60]);
    expect(head.rows[0].tableHeader).toBe(true);
    expect(tail.rows.map((r) => r.heightPx)).toEqual([20, 40, 20]);
    expect(tail.rows[0].tableHeader).toBe(true);
    expect(tail.rows[1].tableHeader).toBeUndefined();
    const tailPara = tail.rows[1].cells[0].stack[0].block;
    if (tailPara.kind !== "paragraph") throw new Error("expected paragraph");
    expect(tailPara.lines).toHaveLength(2);
  });

  it("moves the row whole when no cell has content above the cut (guard)", () => {
    // The only content is a nested table — a non-paragraph block cannot split
    // at a line boundary, so the cut would leave an empty-shell head: the row
    // moves whole instead.
    const nested: LayoutTable = {
      kind: "table",
      columnWidthsPx: [300],
      rows: [
        { cells: [{ blocks: [para(1)] }] },
        { cells: [{ blocks: [para(1)] }] },
        { cells: [{ blocks: [para(1)] }] },
      ],
    };
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [300],
      rows: [{ cells: [{ blocks: [nested] }] }], // 60px row
    };
    const pages = flow([para(3), table], 100);
    expect(pages).toHaveLength(2);
    expect(pages[0].items.map((i) => i.block.kind)).toEqual(["paragraph"]);
    const t = pages[1].items[0].block;
    if (t.kind !== "table") throw new Error("expected table");
    expect(t.rows).toHaveLength(1);
    expect(t.rows[0].heightPx).toBe(60);
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

  it("pads a square zone by the anchor's wrap distances (distL/distR)", () => {
    // Without distances the 200px box leaves 100px (12 atoms per line); with
    // 10px pads each side the zone widens to 220px — 80px usable, 10 atoms,
    // so 24 atoms need three lines.
    const d = drawing(0, 0, 200, 40, "square");
    d.distances = { left: 10, right: 10 };
    const pages = flow([wrapPara(24, { drawings: [d] })], 300);
    const [anchor] = paras(pages)[0];
    expect(anchor.lines).toHaveLength(3);
  });

  it("pads a topAndBottom band by distT/distB", () => {
    // The plain band [20, 70] resumes the next paragraph at y 70; 5px pads
    // each side widen it to [15, 75].
    const d = drawing(0, 20, 200, 50, "topAndBottom");
    d.distances = { top: 5, bottom: 5 };
    const pages = flow([para(1, { drawings: [d] }), para(1)], 300);
    expect(pages[0].items[1].yPx).toBe(75);
  });

  it("packs lines right of a wrapSide-right float instead of shrinking", () => {
    // The 100px float sits at x 0: side "right" moves the text past its
    // edge — a 200px line shifted by 100 (192px of atoms fit one line),
    // not the left 200px strip.
    const d = drawing(0, 0, 100, 40, "square");
    d.wrapSide = "right";
    const pages = flow([wrapPara(24, { drawings: [d] })], 300);
    const [anchor] = paras(pages)[0];
    expect(anchor.lines).toHaveLength(1);
    expect(anchor.lines[0]!.xOffsetPx).toBe(100);
    expect(anchor.lines[0]!.maxWidthPx).toBe(200);
  });

  it("wraps a tight contour by slicing the polygon per line", () => {
    // A right triangle (0,0 → 200,0 → 0,40): at the next paragraph's first
    // line (mid-height y 30 relative to the zone) the hypotenuse slice ends
    // at x 50 — narrower than the 200px box, so the text shifts right of
    // the slice's far edge with 250px of room.
    const d = drawing(0, 0, 200, 60, "tight");
    d.contour = [
      { x: 0, y: 0 },
      { x: 200, y: 0 },
      { x: 0, y: 40 },
    ];
    const pages = flow([para(1, { drawings: [d] }), wrapPara(24)], 300);
    const [, body] = paras(pages)[0];
    expect(body.lines[0]!.xOffsetPx).toBe(50);
    expect(body.lines[0]!.maxWidthPx).toBe(250);
  });

  it("lets a line above the contour's slice use the full width", () => {
    // The triangle only occupies the zone's bottom strip (y 30-40): line 0
    // (mid-height 10) slices nothing and keeps the full width; line 1
    // (mid-height 30) cuts the full-width top edge and shifts right of it.
    const d = drawing(0, 0, 200, 40, "tight");
    d.contour = [
      { x: 0, y: 30 },
      { x: 200, y: 30 },
      { x: 0, y: 40 },
    ];
    const pages = flow([wrapPara(48, { drawings: [d] })], 300);
    const [anchor] = paras(pages)[0];
    expect(anchor.lines[0]!.maxWidthPx).toBe(300);
    expect(anchor.lines[0]!.xOffsetPx).toBeUndefined();
    expect(anchor.lines[1]!.xOffsetPx).toBe(200);
    expect(anchor.lines[1]!.maxWidthPx).toBe(100);
  });

  it("picks the wider side for wrapSide largest", () => {
    // Float at x 180-280: the left side (180px) is wider than the right
    // (20px), so the lines shrink from the left with no shift.
    const left = drawing(180, 0, 100, 40, "square");
    left.wrapSide = "largest";
    const pages = flow([wrapPara(24, { drawings: [left] })], 300);
    const [a] = paras(pages)[0];
    expect(a.lines[0]!.xOffsetPx).toBeUndefined();
    expect(a.lines[0]!.maxWidthPx).toBe(180);

    // Float at x 0-100: the right side (200px) wins — the text shifts past it.
    const right = drawing(0, 0, 100, 40, "square");
    right.wrapSide = "largest";
    const shifted = flow([wrapPara(24, { drawings: [right] })], 300);
    const [b] = paras(shifted)[0];
    expect(b.lines[0]!.xOffsetPx).toBe(100);
    expect(b.lines[0]!.maxWidthPx).toBe(200);
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

describe("layoutFlow columns", () => {
  const twoCol = { count: 2, spacePx: 20, separate: false, equalWidth: true };
  const colFlow = (blocks: LayoutBlock[], contentHeightPx: number) =>
    layoutFlow(blocks, { contentWidthPx: 300, contentHeightPx, columns: twoCol }, measurer);

  it("splits the box into equal columns and lays against the column width", () => {
    // 300px box, 20px gap → two 140px columns.
    const pages = colFlow([para(1)], 100);
    expect(pages).toHaveLength(1);
    expect(pages[0].items[0].xPx).toBe(0);
    const laid = pages[0].items[0].block;
    if (laid.kind === "paragraph") expect(laid.lines[0].maxWidthPx).toBe(140);
  });

  it("fills the left column before the right one", () => {
    // Five 20px lines fill a 100px column; the 6th continues in the right
    // column (left edge 140 + 20).
    const pages = colFlow([para(1), para(1), para(1), para(1), para(1), para(1), para(1)], 100);
    expect(pages).toHaveLength(1);
    expect(pages[0].items.map((i) => i.xPx)).toEqual([0, 0, 0, 0, 0, 160, 160]);
  });

  it("pages after the last column fills", () => {
    const pages = colFlow(
      Array.from({ length: 11 }, () => para(1)),
      100,
    );
    expect(pages).toHaveLength(2);
    expect(pages[0].items.map((i) => i.xPx).slice(-1)).toEqual([160]);
    expect(pages[1].items[0].xPx).toBe(0);
  });

  it("a split tail continues in the next column before paging", () => {
    // A 12-line paragraph: 5 lines fill the left column, the tail continues
    // at the top of the right one (5 more fit there), and the remaining 2
    // lines page — a column boundary never skips the sibling column.
    const pages = colFlow([para(12)], 100);
    expect(pages).toHaveLength(2);
    expect(pages[0].items.map((i) => i.xPx)).toEqual([0, 160]);
    expect(pages[1].items[0].xPx).toBe(0);
    const lineCounts = pages.map((p) =>
      p.items.map((i) => (i.block.kind === "paragraph" ? i.block.lines.length : 0)),
    );
    expect(lineCounts).toEqual([[5, 5], [2]]);
  });

  it("columnBreak moves the following content to the next column", () => {
    const blocks: LayoutBlock[] = [para(1), { kind: "columnBreak" }, para(1)];
    const pages = colFlow(blocks, 100);
    expect(pages).toHaveLength(1);
    expect(pages[0].items.map((i) => i.xPx)).toEqual([0, 160]);
  });

  it("explicit widths keep their own boxes", () => {
    const cols = {
      count: 2,
      spacePx: 20,
      separate: false,
      equalWidth: false,
      columnsPx: [100, 180],
    };
    const pages = layoutFlow(
      [{ kind: "columnBreak" }, para(1)],
      { contentWidthPx: 300, contentHeightPx: 100, columns: cols },
      measurer,
    );
    // The break commits nothing — the paragraph is the only item, in column 2.
    expect(pages[0].items).toHaveLength(1);
    expect(pages[0].items[0].xPx).toBe(120);
  });

  it("single-column flows stay unstamped", () => {
    const pages = flow([para(1)], 100);
    expect(pages[0].items[0].xPx).toBeUndefined();
  });
});

describe("layoutFlow footnotes", () => {
  const footnotePara = (id: number, ordinal: number): LayoutParagraph => ({
    kind: "paragraph",
    inline: [
      { kind: "text", text: "text", style: latin },
      {
        kind: "text",
        text: String(ordinal),
        style: latin,
        noteRef: { kind: "footnote", id, ordinal },
      },
    ],
    spacing: exact20,
    defaultTextStyle: latin,
    widowControl: false,
  });

  const noteBody = (text: string): LayoutBlock[] => [
    {
      kind: "paragraph",
      inline: [{ kind: "text", text, style: latin }],
      spacing: exact20,
      defaultTextStyle: latin,
      widowControl: false,
    },
  ];

  it("places footnote at the bottom of the page where reference lands", () => {
    const fnDefs = new Map<number, readonly LayoutBlock[]>([[1, noteBody("Footnote 1 text")]]);
    const pages = layoutFlow(
      [footnotePara(1, 1)],
      { contentWidthPx: 300, contentHeightPx: 200, footnoteDefinitions: fnDefs },
      measurer,
    );
    expect(pages).toHaveLength(1);
    expect(pages[0].footnotes).toBeDefined();
    expect(pages[0].footnotes?.separatorWidthPx).toBe(192);
    expect(pages[0].footnotes?.notes).toHaveLength(1);
    expect(pages[0].footnotes?.notes[0].id).toBe(1);
    expect(pages[0].footnotes?.notes[0].ordinal).toBe(1);
    // Note height: 20px, separator: 17px, total: 37px
    // yPx: 200 - 37 = 163
    expect(pages[0].footnotes?.totalHeightPx).toBe(37);
    expect(pages[0].footnotes?.yPx).toBe(163);
    expect(pages[0].footnotes?.items).toHaveLength(1);
  });

  it("reserves vertical space for footnote so body wraps to next page earlier", () => {
    // Page height = 100px (normally fits 5 lines of 20px).
    // Footnote consumes 17px (separator) + 20px (note) = 37px.
    // Available body height = 100 - 37 = 63px (fits only 3 lines of 20px).
    // 4 lines of 20px will not fit on page 0 with the footnote, so the 4th line must move to page 1.
    const fnDefs = new Map<number, readonly LayoutBlock[]>([[1, noteBody("Note 1")]]);
    const p4 = para(4, {
      inline: [
        {
          kind: "text",
          text: "line 1",
          style: latin,
          noteRef: { kind: "footnote", id: 1, ordinal: 1 },
        },
        { kind: "break" },
        { kind: "text", text: "line 2", style: latin },
        { kind: "break" },
        { kind: "text", text: "line 3", style: latin },
        { kind: "break" },
        { kind: "text", text: "line 4", style: latin },
      ],
    });
    const pages = layoutFlow(
      [p4],
      { contentWidthPx: 300, contentHeightPx: 100, footnoteDefinitions: fnDefs },
      measurer,
    );
    expect(pages).toHaveLength(2);
    // Page 0 holds 3 lines and has the footnote
    expect(pages[0].items[0].block.kind).toBe("paragraph");
    if (pages[0].items[0].block.kind === "paragraph") {
      expect(pages[0].items[0].block.lines).toHaveLength(3);
    }
    expect(pages[0].footnotes).toBeDefined();
    // Page 1 holds the 4th line and has NO footnotes
    expect(pages[1].footnotes).toBeUndefined();
    if (pages[1].items[0].block.kind === "paragraph") {
      expect(pages[1].items[0].block.lines).toHaveLength(1);
    }
  });

  it("accumulates multiple footnotes on the same page", () => {
    const fnDefs = new Map<number, readonly LayoutBlock[]>([
      [1, noteBody("Note 1")],
      [2, noteBody("Note 2")],
    ]);
    const pages = layoutFlow(
      [footnotePara(1, 1), footnotePara(2, 2)],
      { contentWidthPx: 300, contentHeightPx: 200, footnoteDefinitions: fnDefs },
      measurer,
    );
    expect(pages).toHaveLength(1);
    expect(pages[0].footnotes?.notes).toHaveLength(2);
    expect(pages[0].footnotes?.notes[0].id).toBe(1);
    expect(pages[0].footnotes?.notes[1].id).toBe(2);
    // Total height = 17 (separator once) + 20 + 20 = 57px
    expect(pages[0].footnotes?.totalHeightPx).toBe(57);
    expect(pages[0].footnotes?.yPx).toBe(200 - 57);
  });
});
