import { describe, expect, it } from "vitest";

import { fakeFontMetrics, installFakeCanvas } from "../../test/fake-canvas";
import type { LayoutBlock, LayoutParagraph, LayoutTable } from "../layout-doc";
import { TextMeasurer } from "../text/measure";
import { layoutFlow } from "./flow";

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
