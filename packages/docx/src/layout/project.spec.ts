import type {
  DocumentOptions,
  ParagraphStyleOptions,
  SectionChild,
  SectionPropertiesOptions,
  StylesOptions,
} from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { projectDocumentOptions, projectFlowBox } from "./project";

// The adapter's contract mirrors the persistence model it consumes:
// per-field cascade (direct pPr → style chain → docDefaults), unit
// conversions (twips/pt/eighths-of-pt/EMU → px), table geometry, placeholder
// boxes for shapes the engine cannot lay out yet.
// Reference values: 1pt = 4/3px, 1tw = 1/15px, 1EMU = 1/9525px, A4 = 11906×16838tw.

// `default` (w:default="1") rides the runtime shape but is not part of the
// public ParagraphStyleOptions type — spell it via intersection.
type DefaultFlaggedStyle = ParagraphStyleOptions & { default?: boolean };

const styles: StylesOptions = {
  paragraphStyles: [
    {
      id: "Normal",
      default: true,
      paragraph: { spacing: { line: 360, lineRule: "auto" } },
      run: { size: 12, font: { ascii: "Times", hAnsi: "Times", eastAsia: "SimSun" } },
    },
    {
      id: "Heading1",
      basedOn: "Normal",
      paragraph: { spacing: { before: 240 }, keepNext: true },
      run: { size: 32, bold: true },
    },
  ] satisfies DefaultFlaggedStyle[],
  default: {
    document: {
      paragraph: { spacing: { after: 160 }, indent: { firstLineChars: 200 } },
      run: { size: 10.5 },
    },
  },
};

const doc = (
  children: SectionChild[],
  sectionProps?: SectionPropertiesOptions,
): DocumentOptions => ({
  styles,
  sections: [{ children, ...(sectionProps ? { properties: sectionProps } : {}) }],
});

describe("projectDocumentOptions style cascade", () => {
  it("resolves spacing/indent per field through the whole cascade", () => {
    const { blocks } = projectDocumentOptions(doc([{ paragraph: { children: ["hi"] } }]));
    const para = blocks[0];
    expect(para?.kind).toBe("paragraph");
    if (para?.kind !== "paragraph") return;
    // line=360 auto from Normal (style chain); after=160 from docDefaults;
    // before unset → 0; firstLineChars=200 → 2em × Normal's 12pt.
    expect(para.spacing!.lineHeight).toEqual({ rule: "multiple", factor: 1.5 });
    expect(para.spacing!.afterPx).toBeCloseTo(160 / 15, 5);
    expect(para.spacing!.beforePx).toBe(0);
    expect(para.indent?.firstLinePx).toBeCloseTo(2 * 12 * (4 / 3), 5);
    // Default run: the style chain's Normal run (12pt, slot fonts).
    expect(para.defaultTextStyle?.sizePx).toBeCloseTo(12 * (4 / 3), 5);
    expect(para.defaultTextStyle?.family).toEqual({ latin: "Times", eastAsia: "SimSun" });
  });

  it("direct attrs win and basedOn ancestors fill gaps (Heading1 → Normal)", () => {
    const { blocks } = projectDocumentOptions(
      doc([
        {
          paragraph: {
            heading: "Heading1",
            spacing: { before: 480 },
            children: ["H"],
          },
        },
      ]),
    );
    const h = blocks[0];
    if (h?.kind !== "paragraph") throw new Error("expected paragraph");
    // Direct before=480 wins over Heading1's 240; line=360 comes through the
    // basedOn chain from Normal; keepNext comes from the style chain.
    expect(h.spacing!.beforePx).toBeCloseTo(480 / 15, 5);
    expect(h.spacing!.lineHeight).toEqual({ rule: "multiple", factor: 1.5 });
    expect(h.defaultTextStyle?.sizePx).toBeCloseTo(32 * (4 / 3), 5);
    expect(h.defaultTextStyle?.bold).toBe(true);
    expect(h.keepNext).toBe(true);
    expect(h.widowControl).toBe(true);
  });

  it("resolves run rPr over the defaults and converts breaks/pictures", () => {
    const { blocks } = projectDocumentOptions(
      doc([
        {
          paragraph: {
            children: [
              {
                text: "abc",
                size: 24,
                bold: true,
                characterSpacing: 20,
              },
              { text: "def" },
              { break: 1 },
              {
                children: [
                  {
                    picture: {
                      type: "png",
                      data: "x",
                      transformation: { width: 609600, height: 457200 },
                    },
                  },
                ],
              },
              {
                picture: {
                  type: "png",
                  data: "eA",
                  transformation: { width: 609600, height: 457200 },
                },
              },
            ],
          },
        },
      ]),
    );
    const para = blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected paragraph");
    expect(para.inline).toHaveLength(5);
    const [marked, plain, br, pic, topLevelPic] = para.inline;
    if (marked.kind === "text") {
      expect(marked.style.sizePx).toBeCloseTo(24 * (4 / 3), 5);
      expect(marked.style.bold).toBe(true);
      expect(marked.style.letterSpacingPx).toBeCloseTo(20 / 15, 5);
    } else throw new Error("expected text");
    if (plain.kind === "text") {
      // An unstyled run falls back to the paragraph default (Normal 12pt).
      expect(plain.style.sizePx).toBeCloseTo(12 * (4 / 3), 5);
      expect(plain.style.bold).toBeUndefined();
    }
    expect(br).toEqual({ kind: "break" });
    // 609600×457200 EMU = 64×48 px at 96 dpi — nested and paragraph-child slot.
    expect(pic).toEqual({
      kind: "picture",
      widthPx: 64,
      heightPx: 48,
      src: "data:image/png;base64,x",
    });
    expect(topLevelPic).toEqual({
      kind: "picture",
      widthPx: 64,
      heightPx: 48,
      src: "data:image/png;base64,eA",
    });
  });

  it("projects run color and underline/strike decorations", () => {
    const { blocks } = projectDocumentOptions(
      doc([
        {
          paragraph: {
            children: [
              { text: "link", color: "0563C1", underline: { type: "single" } },
              { text: "plain" },
              { text: "cut", strike: true },
              { text: "none", color: "auto", underline: { type: "none" } },
            ],
          },
        },
      ]),
    );
    const para = blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected paragraph");
    const [link, plain, cut, none] = para.inline.map((i) => (i.kind === "text" ? i.style : null));
    expect(link?.color).toBe("0563C1");
    expect(link?.underline).toBe(true);
    expect(plain?.color).toBeUndefined();
    expect(plain?.underline).toBeUndefined();
    expect(cut?.strikethrough).toBe(true);
    // "auto" color and underline type none carry no decoration.
    expect(none?.color).toBeUndefined();
    expect(none?.underline).toBeUndefined();
  });

  it("numbers list markers in document order with per-format composition", () => {
    const numbered = (reference: string, level = 0): SectionChild => ({
      paragraph: { numbering: { reference, level }, children: ["x"] },
    });
    const { blocks } = projectDocumentOptions({
      styles,
      numbering: {
        abstractNumberings: [
          {
            reference: "L",
            levels: [
              { level: 0, format: "decimal", text: "%1." },
              { level: 1, format: "lowerLetter", text: "%2)" },
              { level: 2, format: "chineseCounting", text: "%3、" },
            ],
          },
          { reference: "R", levels: [{ level: 0, format: "upperRoman", text: "(%1)" }] },
        ],
      },
      sections: [
        {
          children: [
            numbered("L"),
            numbered("L", 1),
            numbered("L", 2),
            numbered("L", 1),
            // A deeper level re-entered resets nothing above it; a new
            // reference counts independently.
            numbered("R"),
            numbered("L"),
          ],
        },
      ],
    });
    const markers = blocks.map((b) => {
      if (b?.kind !== "paragraph") throw new Error("expected paragraph");
      const first = b.inline[0];
      return first?.kind === "text" ? first.text : null;
    });
    expect(markers).toEqual(["1.", "a)", "一、", "b)", "(I)", "2."]);
  });

  it("converts exact/atLeast line rules and twip indents to px", () => {
    const { blocks } = projectDocumentOptions(
      doc([
        {
          paragraph: {
            spacing: { line: 400, lineRule: "exact", before: 100, after: 100 },
            indent: { left: 720, right: 360, firstLine: 480 },
            children: ["x"],
          },
        },
      ]),
    );
    const p = blocks[0];
    if (p?.kind !== "paragraph") throw new Error("expected paragraph");
    expect(p.spacing!.lineHeight).toEqual({ rule: "exact", px: 400 / 15 });
    expect(p.spacing!.beforePx).toBeCloseTo(100 / 15, 5);
    expect(p.indent).toEqual({ leftPx: 48, rightPx: 24, firstLinePx: 32 });
  });
});

describe("projectDocumentOptions blocks", () => {
  it("projects tables with converted geometry and threaded styles", () => {
    const { blocks } = projectDocumentOptions(
      doc([
        {
          table: {
            width: { size: 50, type: "percent" },
            columnWidths: [3000, 1500],
            margins: { left: { size: 108, type: "twips" }, right: { size: 108, type: "twips" } },
            rows: [
              {
                height: { value: 600, rule: "atLeast" },
                cells: [
                  {
                    columnSpan: 2,
                    borders: { left: { style: "single", size: 8 } },
                    margins: { left: { size: 60, type: "twips" } },
                    children: [{ paragraph: { children: ["c"] } }],
                  },
                ],
              },
            ],
          },
        },
      ]),
    );
    const table = blocks[0];
    if (table?.kind !== "table") throw new Error("expected table");
    expect(table.width).toEqual({ type: "percent", percent: 50 });
    expect(table.columnWidthsPx).toEqual([200, 100]); // twips ÷ 15
    expect(table.cellInsets).toEqual({ left: 108 / 15, right: 108 / 15 });
    const row = table.rows[0];
    expect(row.height).toEqual({ rule: "atLeast", px: 40 });
    const cell = row.cells[0];
    expect(cell.colspan).toBe(2);
    expect(cell.insets).toEqual({ left: 4 }); // cell's own tcMar wins per side
    expect(cell.borders?.left).toEqual({ style: "single", px: (8 / 8) * (4 / 3) });
    // Cell content projected as a paragraph through the same cascade.
    const para = cell.blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected paragraph in cell");
    expect(para.inline[0]).toMatchObject({ kind: "text", text: "c" });
  });

  it("projects cell border colors, shading fills, and table-level borders", () => {
    const stylesWithTable: StylesOptions = {
      ...styles,
      tableStyles: [
        {
          id: "TableGrid",
          table: {
            borders: {
              top: { style: "single", color: "000000", size: 4 },
              insideHorizontal: { style: "single", color: "000000", size: 4 },
              insideVertical: { style: "single", color: "000000", size: 4 },
            },
          },
        },
      ],
    };
    const { blocks } = projectDocumentOptions({
      styles: stylesWithTable,
      sections: [
        {
          children: [
            {
              table: {
                style: "TableGrid",
                borders: { bottom: { style: "double", color: "FF0000", size: 4 } },
                rows: [
                  {
                    cells: [
                      {
                        borders: { top: { style: "single", color: "00FF00", size: 8 } },
                        shading: { type: "clear", fill: "D9D9D9" },
                        children: [{ paragraph: { children: ["c"] } }],
                      },
                    ],
                  },
                ],
              },
            },
          ],
        },
      ],
    });
    const table = blocks[0];
    if (table?.kind !== "table") throw new Error("expected table");
    // Cell edge: color rides along; style-chain inside edges and the direct
    // bottom merge per side into the table-level defaults.
    const cell = table.rows[0].cells[0];
    expect(cell.borders?.top).toEqual({ style: "single", px: (8 / 8) * (4 / 3), color: "00FF00" });
    expect(cell.fill).toBe("D9D9D9");
    expect(table.borders?.top).toEqual({ style: "single", px: (4 / 8) * (4 / 3), color: "000000" });
    expect(table.borders?.bottom).toEqual({
      style: "double",
      px: (4 / 8) * (4 / 3),
      color: "FF0000",
    });
    expect(table.borders?.insideVertical).toEqual({
      style: "single",
      px: (4 / 8) * (4 / 3),
      color: "000000",
    });
    expect(table.borders?.left).toBeUndefined();
  });

  it("turns unprojectable body shapes into labeled placeholders and drops bookmarks", () => {
    const { blocks } = projectDocumentOptions(
      doc([
        { bookmarkStart: { id: 1, name: "n" } },
        { rawXml: "<w:customWrap/>" },
        { toc: {} },
        { bookmarkEnd: { id: 1 } },
        { paragraph: { text: "after" } },
      ]),
    );
    expect(blocks).toHaveLength(3);
    expect(blocks[0]).toEqual({ kind: "placeholder", heightPx: 48, label: "rawXml" });
    expect(blocks[1]).toEqual({ kind: "placeholder", heightPx: 48, label: "toc" });
    expect(blocks[2]).toMatchObject({ kind: "paragraph" });
  });
});

describe("projectDocumentOptions fields and furniture", () => {
  it("projects PAGE/NUMPAGES fields as dynamic atoms and other fields as cached text", () => {
    const { blocks } = projectDocumentOptions(
      doc([
        {
          paragraph: {
            children: [
              { complexField: { instruction: "PAGE \\* MERGEFORMAT", result: "1" } },
              { complexField: { instruction: "NUMPAGES", result: "9" } },
              { simpleField: { instruction: "CREATEDATE", cachedValue: "2024-01-01" } },
            ],
          },
        },
      ]),
    );
    const para = blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected paragraph");
    const [page, numPages, created] = para.inline;
    // `text` is only a measuring placeholder — the painter swaps the live
    // number in per page; the cached "1"/"9" must NOT leak into the layout.
    expect(page).toMatchObject({ kind: "text", text: "0", field: "page" });
    expect(numPages).toMatchObject({ kind: "text", text: "0", field: "numPages" });
    expect(created).toMatchObject({ kind: "text", text: "2024-01-01" });
  });

  it("projects the first section's header/footer slots with placement flags", () => {
    const { furniture } = projectDocumentOptions({
      styles,
      settings: { evenAndOddHeaders: true },
      sections: [
        {
          children: [],
          headers: { default: [{ paragraph: { children: ["head"] } }] },
          footers: { even: [{ paragraph: { children: ["foot"] } }] },
          properties: {
            titlePage: true,
            pageMargin: { header: 900, footer: 850 },
          },
        },
      ],
    });
    expect(furniture.header).toHaveLength(1);
    expect(furniture.header?.[0]).toMatchObject({ kind: "paragraph" });
    expect(furniture.evenFooter).toHaveLength(1);
    // Slots the document does not define stay undefined (paint-time fallback).
    expect(furniture.firstHeader).toBeUndefined();
    expect(furniture.evenHeader).toBeUndefined();
    expect(furniture.footer).toBeUndefined();
    expect(furniture.titlePage).toBe(true);
    expect(furniture.evenAndOddHeaders).toBe(true);
    expect(furniture.headerDistancePx).toBeCloseTo(900 / 15, 5);
    expect(furniture.footerDistancePx).toBeCloseTo(850 / 15, 5);
  });
});

describe("projectFlowBox", () => {
  it("derives the content box from A4 paper and margins", () => {
    const flow = projectFlowBox({
      pageSize: { width: 11906, height: 16838 },
      pageMargin: { top: 1440, bottom: 1440, left: 1800, right: 1800 },
    });
    expect(flow.pageWidthPx).toBeCloseTo(11906 / 15, 5);
    expect(flow.contentWidthPx).toBeCloseTo((11906 - 3600) / 15, 5);
    expect(flow.contentHeightPx).toBeCloseTo((16838 - 2880) / 15, 5);
    expect(flow.linePitchPx).toBeUndefined();
  });

  it("defaults to A4 portrait with zh-CN Normal margins when absent", () => {
    const flow = projectFlowBox(undefined);
    expect(flow.pageWidthPx).toBeCloseTo(11906 / 15, 5);
    expect(flow.contentWidthPx).toBeCloseTo((11906 - 3600) / 15, 5);
  });
});
