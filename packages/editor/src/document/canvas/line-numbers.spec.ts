// @vitest-environment node
import type { FlowPage, ProjectedFlowBox, ProjectedLineNumbers } from "@docen/layout";
import { describe, expect, it } from "vitest";

import type { LineNumberSection } from "./line-numbers";
import { computeLineNumbers } from "./line-numbers";

/** A page whose items sit at fixed strides: paragraphs get one line each
 *  (plus a configurable count), tables and other kinds are passed through
 *  as opaque blocks. */
const pageOf = (
  items: { kind: string; lines?: number; suppress?: boolean; sectionEnd?: boolean; xPx?: number }[],
): FlowPage[] =>
  [
    {
      items: items.map((item, i) => ({
        yPx: i * 100,
        xPx: item.xPx,
        block:
          item.kind === "paragraph"
            ? {
                kind: "paragraph",
                suppressLineNumbers: item.suppress,
                sectionEnd: item.sectionEnd,
                markSizePx: 14,
                lines: Array.from({ length: item.lines ?? 1 }, (_, l) => ({
                  yPx: l * 20,
                  heightPx: 20,
                  naturalPx: 14,
                })),
              }
            : { kind: item.kind },
      })),
    },
  ] as unknown as FlowPage[];

const FLOW: ProjectedFlowBox = {
  pageWidthPx: 794,
  pageHeightPx: 1123,
  contentWidthPx: 554,
  contentHeightPx: 883,
  contentLeftPx: 120,
  contentTopPx: 120,
};

const section = (extra: Partial<LineNumberSection> = {}): LineNumberSection => ({
  flow: FLOW,
  lineNumbers: CONFIG,
  ...extra,
});

const CONFIG: ProjectedLineNumbers = {
  countBy: 1,
  start: 1,
  restart: "newPage",
  distancePx: 12,
};

describe("computeLineNumbers", () => {
  it("numbers every text line, tables skip, suppressed paragraphs render uncounted", () => {
    const pages = pageOf([
      { kind: "paragraph", lines: 2 },
      { kind: "table" },
      { kind: "paragraph", suppress: true, lines: 2 },
      { kind: "paragraph" },
    ]);
    const out = computeLineNumbers(pages, [section()], [0]);
    const marks = out.get(0)!;
    // Table lines don't count; the suppressed paragraph's two lines render
    // but don't either — the last paragraph's line is count 3.
    expect(marks.map((m) => m.num)).toEqual([1, 2, 3]);
    // The mark rides the line's box (item y + line y + no grid pad).
    expect(marks[2]!.yPx).toBe(300);
    expect(marks[2]!.sizePx).toBe(14);
  });

  it("shows only every countBy-th number, starting at `start`", () => {
    const pages = pageOf([
      { kind: "paragraph" },
      { kind: "paragraph" },
      { kind: "paragraph" },
      { kind: "paragraph" },
      { kind: "paragraph" },
    ]);
    const out = computeLineNumbers(
      pages,
      [section({ lineNumbers: { ...CONFIG, countBy: 2, start: 10 } })],
      [0],
    );
    expect(out.get(0)!.map((m) => m.num)).toEqual([11, 13]);
  });

  it("restart resets per page (newPage) or runs on (continuous)", () => {
    const two = (config: ProjectedLineNumbers): Map<number, { num: number }[]> =>
      computeLineNumbers(
        [
          ...pageOf([{ kind: "paragraph" }, { kind: "paragraph" }]),
          ...pageOf([{ kind: "paragraph" }]),
        ],
        [section({ lineNumbers: config })],
        [0, 0],
      );
    const newPage = two({ ...CONFIG, restart: "newPage" });
    expect(newPage.get(0)!.map((m) => m.num)).toEqual([1, 2]);
    expect(newPage.get(1)!.map((m) => m.num)).toEqual([1]);
    const continuous = two({ ...CONFIG, restart: "continuous" });
    expect(continuous.get(1)!.map((m) => m.num)).toEqual([3]);
  });

  it("continuous carries the count across pages and section breaks", () => {
    const pages = [
      ...pageOf([{ kind: "paragraph" }]),
      ...pageOf([{ kind: "paragraph" }]),
      ...pageOf([{ kind: "paragraph" }]),
    ];
    const config = { ...CONFIG, restart: "continuous" } as const;
    const out = computeLineNumbers(
      pages,
      [section({ lineNumbers: config }), section({ lineNumbers: config })],
      [0, 0, 1],
    );
    expect(out.get(0)!.map((m) => m.num)).toEqual([1]);
    // Across the page break…
    expect(out.get(1)!.map((m) => m.num)).toEqual([2]);
    // …and across the section break (a section break forces a new page, so
    // this is the same continuous run).
    expect(out.get(2)!.map((m) => m.num)).toEqual([3]);
  });

  it("pages of unnumbered sections stay empty", () => {
    const pages = [
      ...pageOf([{ kind: "paragraph" }]),
      ...pageOf([{ kind: "paragraph" }, { kind: "paragraph" }]),
    ];
    const out = computeLineNumbers(pages, [section(), { flow: FLOW }], [0, 1]);
    expect(out.get(0)!.map((m) => m.num)).toEqual([1]);
    expect(out.get(1)!).toEqual([]);
  });

  it("newSection resets when the next numbered section begins", () => {
    const pages = [
      ...pageOf([{ kind: "paragraph" }, { kind: "paragraph" }]),
      ...pageOf([{ kind: "paragraph" }]),
    ];
    const out = computeLineNumbers(
      pages,
      [
        section({ lineNumbers: { ...CONFIG, restart: "newSection" } }),
        section({ lineNumbers: { ...CONFIG, restart: "newSection" } }),
      ],
      [0, 1],
    );
    expect(out.get(0)!.map((m) => m.num)).toEqual([1, 2]);
    expect(out.get(1)!.map((m) => m.num)).toEqual([1]);
  });

  it("a section-break paragraph neither counts nor paints", () => {
    const pages = pageOf([
      { kind: "paragraph" },
      { kind: "paragraph", sectionEnd: true },
      { kind: "paragraph" },
    ]);
    const out = computeLineNumbers(pages, [section()], [0]);
    expect(out.get(0)!.map((m) => m.num)).toEqual([1, 2]);
  });

  it("later columns count but paint only beside the first column", () => {
    const twoCol = {
      count: 2,
      spacePx: 31,
      separate: false,
      equalWidth: true,
    } as const;
    const pages = pageOf([
      { kind: "paragraph" },
      { kind: "paragraph", xPx: 0 },
      { kind: "paragraph", xPx: 293 },
    ]);
    const out = computeLineNumbers(pages, [section({ columns: twoCol })], [0]);
    // Column 2's line (count 3) keeps the count running but paints nothing.
    expect(out.get(0)!.map((m) => m.num)).toEqual([1, 2]);
  });
});
