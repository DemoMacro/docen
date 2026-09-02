import { describe, expect, it } from "vitest";

import { fakeFontMetrics, installFakeCanvas } from "../../test/fake-canvas";
import type { LayoutParagraph, LayoutTable } from "../layout-doc";
import { TextMeasurer } from "../text/measure";
import { layoutBlock, stackBlocks } from "./block";

installFakeCanvas();
const measurer = new TextMeasurer(fakeFontMetrics);

const latin = { family: "serif", sizePx: 16 };
// The fake world: normal ratio 1.2 x 16px.
const NATURAL = 1.2 * 16;
const NOTO_NATURAL = 1.2 * 16;

const para = (over: Partial<LayoutParagraph> = {}): LayoutParagraph => ({
  kind: "paragraph",
  inline: [{ kind: "text", text: "word", style: latin }],
  defaultTextStyle: latin,
  ...over,
});

describe("layoutParagraph line-height semantics", () => {
  it("uses the natural metric with no spacing and no grid", () => {
    const out = layoutBlock(para(), 500, undefined, measurer);
    expect(out.kind).toBe("paragraph");
    if (out.kind === "paragraph") {
      expect(out.heightPx).toBeCloseTo(NATURAL, 4);
      expect(out.beforePx).toBe(0);
      expect(out.afterPx).toBe(0);
    }
  });

  it("exact line height pins every line", () => {
    const out = layoutBlock(
      para({ spacing: { lineHeight: { rule: "exact", px: 40 }, beforePx: 0, afterPx: 0 } }),
      500,
      undefined,
      measurer,
    );
    if (out.kind === "paragraph") expect(out.heightPx).toBe(40);
  });

  it("atLeast takes the max of natural and spec (true semantics)", () => {
    const short = layoutBlock(
      para({ spacing: { lineHeight: { rule: "atLeast", px: 10 }, beforePx: 0, afterPx: 0 } }),
      500,
      undefined,
      measurer,
    );
    const tall = layoutBlock(
      para({ spacing: { lineHeight: { rule: "atLeast", px: 60 }, beforePx: 0, afterPx: 0 } }),
      500,
      undefined,
      measurer,
    );
    if (short.kind === "paragraph") expect(short.heightPx).toBeCloseTo(NATURAL, 4);
    if (tall.kind === "paragraph") expect(tall.heightPx).toBe(60);
  });

  it("multiple resolves against the grid pitch when one is defined", () => {
    const pitch = 25;
    const out = layoutBlock(
      para({ spacing: { lineHeight: { rule: "multiple", factor: 1.5 }, beforePx: 0, afterPx: 0 } }),
      500,
      { linePitchPx: pitch },
      measurer,
    );
    if (out.kind === "paragraph") expect(out.heightPx).toBeCloseTo(37.5, 4);
  });

  it("flags grid lines only for grid contexts (onGrid) with non-overriding spacing", () => {
    const spec = { lineHeight: { rule: "multiple" as const, factor: 2 }, beforePx: 0, afterPx: 0 };
    const body = layoutBlock(
      para({ spacing: spec }),
      500,
      { linePitchPx: 25, onGrid: true },
      measurer,
    );
    if (body.kind === "paragraph") {
      expect(body.lines[0].grid).toBe(true);
      expect(body.lines[0].naturalPx).toBeGreaterThan(0);
    }
    // Pitch without onGrid (the text-box stack caller: pitch-scaled heights,
    // glyphs at the line-box top — user-verified vs Word): heights still
    // scale, and the line never carries the lattice-placement flag.
    const stack = layoutBlock(para({ spacing: spec }), 500, { linePitchPx: 25 }, measurer);
    if (stack.kind === "paragraph") {
      expect(stack.heightPx).toBeCloseTo(50, 4);
      expect(stack.lines[0].grid).toBeUndefined();
    }
    // exact/atLeast own the height outright — never a grid line.
    const exact = layoutBlock(
      para({ spacing: { lineHeight: { rule: "exact", px: 40 }, beforePx: 0, afterPx: 0 } }),
      500,
      { linePitchPx: 25, onGrid: true },
      measurer,
    );
    if (exact.kind === "paragraph") expect(exact.lines[0].grid).toBeUndefined();
  });

  it("snaps a CJK line up to a whole pitch multiple, Latin to max(natural, pitch)", () => {
    const pitch = Math.ceil(NOTO_NATURAL) + 3; // CJK natural < pitch < 2×pitch
    const cjkStyle = { family: { latin: "Inter", eastAsia: "Noto Sans SC" }, sizePx: 16 };
    const cjk = layoutBlock(
      para({ inline: [{ kind: "text", text: "中文", style: cjkStyle }] }),
      500,
      { linePitchPx: pitch },
      measurer,
    );
    if (cjk.kind === "paragraph") expect(cjk.heightPx).toBeCloseTo(pitch, 4);

    const latinOut = layoutBlock(para(), 500, { linePitchPx: Math.ceil(NATURAL) + 10 }, measurer);
    if (latinOut.kind === "paragraph") {
      expect(latinOut.heightPx).toBeCloseTo(Math.ceil(NATURAL) + 10, 4);
    }
  });

  it("floors a table cell's snapped line at max(natural, pitch)", () => {
    const pitch = Math.ceil(NATURAL) + 5;
    const out = layoutBlock(para(), 500, { linePitchPx: pitch, inTable: true }, measurer);
    if (out.kind === "paragraph") expect(out.heightPx).toBeCloseTo(pitch, 4);
  });

  it("bases a table cell's CJK multiple line on the pitch without row snapping", () => {
    // Corpus-verified (honor table, 340-twip grid): a 1.5× cell line measures
    // 1.5 × pitch — the cell floors at its demand, never rounding up to whole
    // grid rows (the row's trHeight floors separately). The earlier E14
    // "renders TWO rows" reading mistook 1.5×pitch in px for 2 rows in pt.
    // Body (non-cell) keeps its whole-row rounding.
    const pitch = 25;
    const cjkStyle = { family: { latin: "serif", eastAsia: "SimSun" }, sizePx: 16 };
    const cjk = (inTable: boolean) =>
      layoutBlock(
        para({
          inline: [{ kind: "text", text: "中文", style: cjkStyle }],
          spacing: { lineHeight: { rule: "multiple", factor: 1.5 }, beforePx: 0, afterPx: 0 },
        }),
        500,
        { linePitchPx: pitch, inTable },
        measurer,
      );
    if (cjk(true).kind === "paragraph") expect(cjk(true).heightPx).toBeCloseTo(37.5, 4);
    if (cjk(false).kind === "paragraph") expect(cjk(false).heightPx).toBeCloseTo(50, 4);
  });

  it("honors snapToGrid=false by dropping the pitch", () => {
    const out = layoutBlock(para({ snapToGrid: false }), 500, { linePitchPx: 100 }, measurer);
    if (out.kind === "paragraph") expect(out.heightPx).toBeCloseTo(NATURAL, 4);
  });

  it("sizes an empty paragraph at the ¶-mark strut (no grid pitch)", () => {
    const out = layoutBlock(
      para({ inline: [], markSizePx: 24 }),
      500,
      { linePitchPx: 100 },
      measurer,
    );
    if (out.kind === "paragraph") {
      expect(out.heightPx).toBe(24);
      expect(out.lines).toHaveLength(1);
      expect(out.lines[0].items).toHaveLength(0);
    }
    // Without a mark size: the default run's natural metric.
    const natural = layoutBlock(para({ inline: [] }), 500, { linePitchPx: 100 }, measurer);
    if (natural.kind === "paragraph") expect(natural.heightPx).toBeCloseTo(NATURAL, 4);
  });

  it("carries w:suppressLineNumbers to the laid block", () => {
    const out = layoutBlock(para({ suppressLineNumbers: true }), 500, undefined, measurer);
    if (out.kind === "paragraph") expect(out.suppressLineNumbers).toBe(true);
  });

  it("emits per-line y positions and split points", () => {
    const out = layoutBlock(
      para({ inline: [{ kind: "text", text: "aaa bbb ccc ddd", style: latin }] }),
      60,
      undefined,
      measurer,
    );
    if (out.kind === "paragraph") {
      expect(out.lines.length).toBeGreaterThan(1);
      let y = 0;
      for (const line of out.lines) {
        expect(line.yPx).toBeCloseTo(y, 6);
        y += line.heightPx;
      }
      expect(out.heightPx).toBeCloseTo(y, 6);
      // The last line's end is a split point at the paragraph's end — the
      // inclusive marker: the line's content ends at its final inline.
      const last = out.lines[out.lines.length - 1];
      expect(last.endInlineIndex).toBe(0);
    }
  });

  it("justifies wrapped lines to the full width, not the last line", () => {
    // 7 CJK chars = 112px in a 90px width: line 0 takes 5 (80px), leaving
    // 10px of slack over 4 gaps → 2.5px per gap; the last line stays natural.
    const cjkStyle = { family: { latin: "serif", eastAsia: "SimSun" }, sizePx: 16 };
    const out = layoutBlock(
      para({
        inline: [{ kind: "text", text: "中文中文中中中", style: cjkStyle }],
        align: "both",
      }),
      90,
      undefined,
      measurer,
    );
    if (out.kind !== "paragraph") return;
    expect(out.lines).toHaveLength(2);
    expect(out.lines[0].justifyGapPx).toBeCloseTo(2.5, 5);
    expect(out.lines[0].items[0].xPx).toBe(0);
    expect(out.lines[1].justifyGapPx).toBeUndefined();
    // Without align, nothing stretches.
    const plain = layoutBlock(
      para({ inline: [{ kind: "text", text: "中文中文中中中", style: cjkStyle }] }),
      90,
      undefined,
      measurer,
    );
    if (plain.kind === "paragraph") expect(plain.lines[0].justifyGapPx).toBeUndefined();
  });

  it("keeps a hard-break line natural under justify", () => {
    // "中中" = 32px then a hard break: the line ends at the break, so no
    // stretch despite align both and a following line.
    const cjkStyle = { family: { latin: "serif", eastAsia: "SimSun" }, sizePx: 16 };
    const out = layoutBlock(
      para({
        inline: [
          { kind: "text", text: "中中", style: cjkStyle },
          { kind: "break" },
          { kind: "text", text: "中中", style: cjkStyle },
        ],
        align: "both",
      }),
      100,
      undefined,
      measurer,
    );
    if (out.kind === "paragraph") {
      expect(out.lines).toHaveLength(2);
      expect(out.lines[0].justifyGapPx).toBeUndefined();
    }
  });

  it("centers and right-aligns a line by its slack", () => {
    // "中文中" = 48px in a 100px width: center shifts by 26, right by 52.
    const cjkStyle = { family: { latin: "serif", eastAsia: "SimSun" }, sizePx: 16 };
    const make = (align: "center" | "right") =>
      layoutBlock(
        para({ inline: [{ kind: "text", text: "中文中", style: cjkStyle }], align }),
        100,
        undefined,
        measurer,
      );
    const centered = make("center");
    const righted = make("right");
    if (centered.kind === "paragraph" && righted.kind === "paragraph") {
      expect(centered.lines[0].items[0].xPx).toBeCloseTo(26, 5);
      expect(righted.lines[0].items[0].xPx).toBeCloseTo(52, 5);
    }
  });

  it("distribute stretches the last line too, both does not", () => {
    // 7 CJK chars in a 90px width: two lines (5 + 2 chars). Under "both" the
    // last line keeps its natural width; under "distribute" it stretches —
    // 32px of content in 90px leaves 58px over the single gap.
    const cjkStyle = { family: { latin: "serif", eastAsia: "SimSun" }, sizePx: 16 };
    const make = (align: "both" | "distribute") =>
      layoutBlock(
        para({ inline: [{ kind: "text", text: "中文中文中中中", style: cjkStyle }], align }),
        90,
        undefined,
        measurer,
      );
    const both = make("both");
    const dist = make("distribute");
    if (both.kind === "paragraph" && dist.kind === "paragraph") {
      expect(both.lines[1].justifyGapPx).toBeUndefined();
      expect(dist.lines[1].justifyGapPx).toBeCloseTo(58, 5);
    }
  });
});

describe("stackBlocks", () => {
  it("collapses middles and contains edge margins (BFC model)", () => {
    const stacked = stackBlocks(
      [
        para({ spacing: { beforePx: 10, afterPx: 6 } }),
        para({ spacing: { beforePx: 4, afterPx: 8 } }),
        para({ spacing: { beforePx: 2, afterPx: 5 } }),
      ],
      500,
      undefined,
      measurer,
    );
    // 10 + N + max(6,4) + N + max(8,2) + N + 5
    const expected = 10 + NATURAL + 6 + NATURAL + 8 + NATURAL + 5;
    expect(stacked.heightPx).toBeCloseTo(expected, 4);
  });
});

describe("layoutTable", () => {
  const cellPara = (t = "word"): LayoutParagraph => ({
    kind: "paragraph",
    inline: [{ kind: "text", text: t, style: latin }],
    defaultTextStyle: latin,
  });

  it("scales the grid to the effective width and measures the tallest row", () => {
    const table: LayoutTable = {
      kind: "table",
      width: { type: "percent", percent: 50 },
      columnWidthsPx: [300, 100],
      rows: [
        {
          cells: [{ blocks: [cellPara()] }, { blocks: [cellPara()] }],
        },
      ],
    };
    const out = layoutBlock(table, 400, undefined, measurer);
    if (out.kind !== "table") throw new Error("expected table");
    expect(out.widthPx).toBe(200);
    expect(out.columnWidthsPx[0]).toBeCloseTo(150, 4);
    expect(out.columnWidthsPx[1]).toBeCloseTo(50, 4);
    // One line of natural height, no insets/borders declared.
    expect(out.heightPx).toBeCloseTo(NATURAL, 4);
  });

  it("subtracts cell insets and side borders from the wrapping width", () => {
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [200],
      rows: [
        {
          cells: [
            {
              insets: { left: 10, right: 10, top: 4, bottom: 6 },
              borders: {
                left: { style: "single", px: 2 },
                right: { style: "single", px: 2 },
              },
              blocks: [cellPara("aaaaaaaaaaaaaaaaaaa")],
            },
          ],
        },
      ],
    };
    const out = layoutBlock(table, 200, undefined, measurer);
    if (out.kind !== "table") throw new Error("expected table");
    expect(out.rows[0].cells[0].innerWidthPx).toBe(200 - 20 - 4);
    // height = natural + top/bottom insets (no vertical borders declared)
    expect(out.rows[0].heightPx).toBeCloseTo(NATURAL + 10, 4);
  });

  it("uses only the max vertical border and inherits table insets per side", () => {
    const table: LayoutTable = {
      kind: "table",
      cellInsets: { top: 3, left: 8 },
      columnWidthsPx: [200],
      rows: [
        {
          cells: [
            {
              insets: { bottom: 2 },
              borders: {
                top: { style: "single", px: 1 },
                bottom: { style: "single", px: 3 },
              },
              blocks: [cellPara()],
            },
          ],
        },
      ],
    };
    const out = layoutBlock(table, 200, undefined, measurer);
    if (out.kind !== "table") throw new Error("expected table");
    const cell = out.rows[0].cells[0];
    expect(cell.insets).toEqual({ top: 3, left: 8, bottom: 2 });
    expect(cell.innerWidthPx).toBeCloseTo(192, 4);
    // natural + 3 + 2 insets + max(1, 3) border
    expect(out.rows[0].heightPx).toBeCloseTo(NATURAL + 5 + 3, 4);
  });

  it("treats nil borders as zero width and applies trHeight rules", () => {
    const nil: LayoutTable = {
      kind: "table",
      columnWidthsPx: [200],
      rows: [
        {
          cells: [
            { borders: { left: { style: "nil" }, right: { style: "none" } }, blocks: [cellPara()] },
          ],
        },
      ],
    };
    const nilOut = layoutBlock(nil, 200, undefined, measurer);
    if (nilOut.kind !== "table") throw new Error("expected table");
    expect(nilOut.rows[0].cells[0].innerWidthPx).toBe(200);

    const atLeast: LayoutTable = {
      kind: "table",
      columnWidthsPx: [200],
      rows: [{ height: { rule: "atLeast", px: 99 }, cells: [{ blocks: [cellPara()] }] }],
    };
    const floorOut = layoutBlock(atLeast, 200, undefined, measurer);
    if (floorOut.kind !== "table") throw new Error("expected table");
    // atLeast floors the CONTENT box; the cell's margins ride on top (the
    // Word E0/E5 experiment: 18.7pt trHeight + 8.5pt margins + border).
    expect(floorOut.rows[0].heightPx).toBe(99);

    const exact: LayoutTable = {
      kind: "table",
      columnWidthsPx: [200],
      rows: [{ height: { rule: "exact", px: 20 }, cells: [{ blocks: [cellPara()] }] }],
    };
    const exactOut = layoutBlock(exact, 200, undefined, measurer);
    if (exactOut.kind !== "table") throw new Error("expected table");
    expect(exactOut.rows[0].heightPx).toBe(20);
  });

  it("adds cell insets on top of a won trHeight, not into the contest (Word E0)", () => {
    // Content 19.2px, trHeight atLeast 24.93, insets 5.67×2: Word renders
    // 24.93 + 11.33 — the corpus table Word-measured at 27.72pt per row.
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [300],
      rows: [
        {
          height: { rule: "atLeast", px: 24.93 },
          cells: [{ insets: { top: 5.67, bottom: 5.67 }, blocks: [cellPara()] }],
        },
      ],
    };
    const out = layoutBlock(table, 300, undefined, measurer);
    if (out.kind !== "table") throw new Error("expected table");
    expect(out.rows[0].heightPx).toBeCloseTo(24.93 + 11.34, 2);
  });

  it("sums colspan columns and wraps cell text at the combined width", () => {
    const wide = "word word word word word word word word word word";
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [50, 50],
      rows: [
        { cells: [{ blocks: [cellPara(wide)] }, { blocks: [] }] },
        { cells: [{ colspan: 2, blocks: [cellPara(wide)] }] },
      ],
    };
    const out = layoutBlock(table, 100, undefined, measurer);
    if (out.kind !== "table") throw new Error("expected table");
    const narrow = out.rows[0].cells[0].stack[0].block;
    const spanned = out.rows[1].cells[0].stack[0].block;
    if (narrow.kind === "paragraph" && spanned.kind === "paragraph") {
      expect(narrow.lines.length).toBeGreaterThan(spanned.lines.length);
    }
  });

  it("stacks cell paragraphs with collapsing margins inside the cell", () => {
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [200],
      rows: [
        {
          cells: [
            {
              blocks: [
                para({ spacing: { beforePx: 10, afterPx: 6 } }),
                para({ spacing: { beforePx: 4, afterPx: 8 } }),
              ],
            },
          ],
        },
      ],
    };
    const out = layoutBlock(table, 200, undefined, measurer);
    if (out.kind !== "table") {
      throw new Error("expected table");
    }
    expect(out.rows[0].heightPx).toBeCloseTo(10 + NATURAL + 6 + NATURAL + 8, 4);
  });

  it("falls back to first-row cell widths when no grid is given", () => {
    const table: LayoutTable = {
      kind: "table",
      rows: [
        { cells: [{ widthPx: 60, blocks: [] }, { widthPx: 140, blocks: [] }, { blocks: [] }] },
      ],
    };
    const out = layoutBlock(table, 200, undefined, measurer);
    if (out.kind !== "table") throw new Error("expected table");
    expect(out.columnWidthsPx).toHaveLength(3);
    const total = out.columnWidthsPx.reduce((a, b) => a + b, 0);
    expect(total).toBeCloseTo(200, 4);
    // 60:140:0 scaled to 200.
    expect(out.columnWidthsPx[1]).toBeCloseTo(140, 4);
  });
});

describe("layoutTable cell float wraps", () => {
  // A wrapping cell paragraph: latin atoms are 8px under the fake metrics,
  // exact 20px rows keep the arithmetic simple (flow.spec's conventions).
  const wrapPara = (chars: number, over: Partial<LayoutParagraph> = {}): LayoutParagraph => ({
    kind: "paragraph",
    inline: [{ kind: "text", text: "x".repeat(chars), style: latin }],
    spacing: { lineHeight: { rule: "exact", px: 20 }, beforePx: 0, afterPx: 0 },
    defaultTextStyle: latin,
    widowControl: false,
    ...over,
  });
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
  const oneCell = (blocks: LayoutParagraph[]): LayoutTable => ({
    kind: "table",
    columnWidthsPx: [300],
    rows: [{ cells: [{ blocks }] }],
  });
  const cellOf = (out: ReturnType<typeof layoutBlock>) => {
    if (out.kind !== "table") throw new Error("expected table");
    return out.rows[0].cells[0].stack.map((i) => i.block);
  };

  it("wraps the anchor cell paragraph's own lines beside its square float", () => {
    // 24 atoms = 192px: one 300px line without the float, two 96px lines
    // (12 atoms beside the 200px box) with it — the cell paragraph wraps
    // beside its own drawing just like a body paragraph.
    const out = layoutBlock(
      oneCell([wrapPara(24, { drawings: [drawing(0, 0, 200, 40, "square")] })]),
      300,
      undefined,
      measurer,
    );
    const [anchor] = cellOf(out);
    if (anchor.kind !== "paragraph") throw new Error("expected paragraph");
    expect(anchor.lines).toHaveLength(2);
  });

  it("wraps later paragraphs in the cell beside the anchor's float", () => {
    // The 40px-tall box spans the anchor's line and the next paragraph's
    // first line: that line packs beside it (12 atoms), the second is past
    // the zone (full width).
    const out = layoutBlock(
      oneCell([wrapPara(1, { drawings: [drawing(0, 0, 200, 40, "square")] }), wrapPara(24)]),
      300,
      undefined,
      measurer,
    );
    const [anchor, body] = cellOf(out);
    if (anchor.kind !== "paragraph" || body.kind !== "paragraph")
      throw new Error("expected paragraphs");
    expect(anchor.lines).toHaveLength(1);
    expect(body.lines).toHaveLength(2);
    expect(body.lines[0]!.maxWidthPx).toBe(100);
    expect(body.lines[1]!.maxWidthPx).toBe(300);
  });

  it("clamps a cell float that overflows the cell's right edge (layoutInCell)", () => {
    // offset 250 + 200 wide overflows the 300px cell by 150: the box shifts
    // left to touch the cell's right edge (x0 = 300-200 = 100) and the wrap
    // zone follows — the next paragraph's first line packs beside the SHIFTED
    // box (100px), not beside the requested one (250px).
    const out = layoutBlock(
      oneCell([wrapPara(1, { drawings: [drawing(250, 0, 200, 40, "square")] }), wrapPara(24)]),
      300,
      undefined,
      measurer,
    );
    const [anchor, body] = cellOf(out);
    if (anchor.kind !== "paragraph" || body.kind !== "paragraph")
      throw new Error("expected paragraphs");
    expect(body.lines[0]!.maxWidthPx).toBe(100);
    expect(body.lines[1]!.maxWidthPx).toBe(300);
  });

  it("keeps a sibling cell's text unwrapped by another cell's float", () => {
    // The float lives in cell A; cell B's paragraph packs the full 300px.
    const table: LayoutTable = {
      kind: "table",
      columnWidthsPx: [300, 300],
      rows: [
        {
          cells: [
            { blocks: [wrapPara(1, { drawings: [drawing(0, 0, 200, 80, "square")] })] },
            { blocks: [wrapPara(24)] },
          ],
        },
      ],
    };
    const out = layoutBlock(table, 600, undefined, measurer);
    if (out.kind !== "table") throw new Error("expected table");
    const b = out.rows[0].cells[1].stack[0].block;
    if (b.kind !== "paragraph") throw new Error("expected paragraph");
    expect(b.lines).toHaveLength(1);
  });
});
