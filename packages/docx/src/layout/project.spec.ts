import type { LayoutBlock } from "@docen/layout";
import type {
  DocumentOptions,
  HorizontalPositionOptions,
  ParagraphStyleOptions,
  SectionChild,
  SectionPropertiesOptions,
  StylesOptions,
  VerticalPositionOptions,
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

  it("cascades every run field from the style chain to bare runs", () => {
    const withColor: StylesOptions = {
      ...styles,
      paragraphStyles: [
        ...styles.paragraphStyles!,
        {
          id: "Toc3",
          basedOn: "Normal",
          run: { bold: true, color: "C00000" },
        } as NonNullable<typeof styles.paragraphStyles>[number],
      ],
    };
    const { blocks } = projectDocumentOptions({
      styles: withColor,
      sections: [
        {
          children: [
            { paragraph: { style: "Toc3", children: ["entry"] } },
            { paragraph: { style: "Toc3", children: [{ text: "off", bold: false }] } },
          ],
        },
      ],
    });
    const bare = blocks[0];
    expect(bare?.kind).toBe("paragraph");
    if (bare?.kind === "paragraph") {
      const text = bare.inline.find((i) => i.kind === "text");
      // A run with no rPr of its own inherits the style's bold AND color.
      expect(text && text.kind === "text" ? text.style.bold : undefined).toBe(true);
      expect(text && text.kind === "text" ? text.style.color : undefined).toBe("C00000");
    }
    const off = blocks[1];
    expect(off?.kind).toBe("paragraph");
    if (off?.kind === "paragraph") {
      const text = off.inline.find((i) => i.kind === "text");
      // A direct w:b val=0 on the run still beats the style chain.
      expect(text && text.kind === "text" ? text.style.bold : undefined).toBe(false);
    }
  });

  it("keeps the chain font when a run carries an empty font shell", () => {
    // A round-tripped run can carry an rFonts object with no usable slot;
    // that shell must not shadow the style chain's face (it used to resolve
    // to empty slots and the painter fell to its fallback font).
    const { blocks } = projectDocumentOptions(
      doc([
        {
          paragraph: {
            children: [{ text: "x", font: {} as unknown as never }],
          },
        },
      ]),
    );
    const para = blocks[0];
    expect(para?.kind).toBe("paragraph");
    if (para?.kind !== "paragraph") return;
    const text = para.inline.find((i) => i.kind === "text");
    expect(text && text.kind === "text" ? text.style.family : undefined).toEqual({
      latin: "Times",
      eastAsia: "SimSun",
    });
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
    // "auto" color carries no decoration; underline type none is an explicit
    // off (false) that also beats any inherited style underline.
    expect(none?.color).toBeUndefined();
    expect(none?.underline).toBe(false);
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
  it("splits a paragraph at run-level page breaks into pageBreak blocks", () => {
    const { blocks } = projectDocumentOptions(
      doc([
        {
          paragraph: {
            children: ["before", { pageBreak: true }, "after", { pageBreak: true }],
          },
        },
      ]),
    );
    expect(blocks.map((b) => b.kind)).toEqual(["paragraph", "pageBreak", "paragraph", "pageBreak"]);
    // Each non-empty leg keeps its text. No empty tail paragraph: the
    // paragraph mark rides on the break's own line in Word ("———break———¶"),
    // so a trailing break must not open a blank row on the next page.
    const texts = blocks.map((b) =>
      b.kind === "paragraph"
        ? b.inline.map((i) => (i.kind === "text" ? i.text : "?")).join("")
        : "-",
    );
    expect(texts).toEqual(["before", "-", "after", "-"]);
  });

  it("drops the empty legs a break-only paragraph would produce", () => {
    const { blocks } = projectDocumentOptions(
      doc([{ paragraph: { children: [{ pageBreak: true }] } }]),
    );
    // Word renders the break row (break mark + paragraph mark together) on
    // the page the break closes — the next page starts clean, no blank row.
    expect(blocks.map((b) => b.kind)).toEqual(["pageBreak"]);
  });

  it("keeps a pageBreakBefore paragraph as one block", () => {
    const { blocks } = projectDocumentOptions(
      doc([{ paragraph: { children: ["x"], pageBreakBefore: true } }]),
    );
    expect(blocks).toHaveLength(1);
    if (blocks[0]?.kind !== "paragraph") throw new Error("expected paragraph");
    expect(blocks[0].pageBreakBefore).toBe(true);
  });

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

  it("projects rendered TOC entries as real paragraphs", () => {
    const { blocks } = projectDocumentOptions(
      doc([
        {
          toc: {
            entries: [
              {
                paragraph: {
                  style: "10",
                  tabStops: [{ position: 8504, type: "right", leader: "dot" }],
                  children: [
                    { complexField: { instruction: " HYPERLINK \\l _Toc1 ", result: "一、总则1" } },
                  ],
                },
              },
              { paragraph: { children: ["二、范围"] } },
            ],
          },
        },
      ]),
    );
    expect(blocks).toHaveLength(2);
    expect(blocks[0]).toMatchObject({ kind: "paragraph" });
    // The HYPERLINK field projects its cached result as static text.
    expect(JSON.stringify(blocks[0])).toContain("一、总则1");
    expect(blocks[1]).toMatchObject({ kind: "paragraph" });
    expect(JSON.stringify(blocks[1])).toContain("二、范围");
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

describe("projectDocumentOptions drawings", () => {
  it("projects a wpg group into an anchored drawing with resolved child space", () => {
    const px = (emu: number): number => emu / 9525;
    const { blocks } = projectDocumentOptions(
      doc([
        {
          paragraph: {
            children: [
              {
                wpgGroup: {
                  children: [
                    {
                      type: "png",
                      fileName: "image1.png",
                      data: new Uint8Array([105, 105]), // "ii" → base64 "aWk="
                      transformation: {
                        pixels: { x: 50, y: 10 },
                        emus: { x: 500000, y: 100000 },
                        offset: { pixels: { x: 0, y: 0 }, emus: { x: 200000, y: 60000 } },
                        flip: { horizontal: true },
                      },
                    },
                    {
                      type: "wps",
                      transformation: {
                        pixels: { x: 10, y: 2 },
                        emus: { x: 100000, y: 20000 },
                        offset: { pixels: { x: 0, y: 0 }, emus: { x: 300000, y: 0 } },
                      },
                      data: {
                        presetGeometry: { preset: "rect" },
                        fill: { type: "solid", color: "FF0000" },
                        outline: { width: 12700, color: { value: "0000FF" } },
                        children: [],
                      },
                    },
                    {
                      type: "wps",
                      transformation: {
                        pixels: { x: 10, y: 2 },
                        emus: { x: 100000, y: 20000 },
                        offset: { pixels: { x: 0, y: 0 }, emus: { x: 300000, y: 0 } },
                      },
                      // Custom geometry carries no canvas path — no member.
                      data: {
                        customGeometry: { pathList: [] },
                        fill: { type: "none" },
                        children: [],
                      },
                    },
                    {
                      type: "wps",
                      transformation: {
                        pixels: { x: 10, y: 2 },
                        emus: { x: 100000, y: 20000 },
                        offset: { pixels: { x: 0, y: 0 }, emus: { x: 300000, y: 0 } },
                      },
                      // Published 0.12.3 parse bug: nested shape data stringified.
                      data: "[object Object]" as unknown as { children: never[] },
                    },
                  ],
                  transformation: { width: 1905000, height: 1905000 },
                  childOffset: { x: 100000, y: 50000 },
                  childExtent: { cx: 1000000, cy: 200000 },
                  floating: {
                    horizontalPosition: { relative: "column", offset: 47625 },
                    verticalPosition: { relative: "paragraph", offset: 95250 },
                  },
                },
              },
            ],
          },
        },
      ]),
    );
    const para = blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected paragraph");
    const drawing = para.drawings?.[0];
    if (!drawing) throw new Error("expected a drawing");
    // Anchor: column/paragraph relatives carry the EMU offsets as px.
    expect(drawing.anchor.horizontal).toEqual({ relative: "column", offsetPx: px(47625) });
    expect(drawing.anchor.vertical).toEqual({ relative: "paragraph", offsetPx: px(95250) });
    expect(drawing.width).toBeCloseTo(px(1905000), 5);
    // Child space resolved: sx = 1905000/1000000, sy = 1905000/200000.
    const [pic, shape] = drawing.members;
    if (pic?.kind !== "picture") throw new Error("expected picture member");
    expect(pic.x).toBeCloseTo(px((200000 - 100000) * 1.905), 5);
    expect(pic.y).toBeCloseTo(px((60000 - 50000) * 9.525), 5);
    expect(pic.width).toBeCloseTo(px(500000 * 1.905), 5);
    expect(pic.src).toBe("data:image/png;base64,aWk=");
    expect(pic.flipH).toBe(true);
    if (shape?.kind !== "shape") throw new Error("expected shape member");
    expect(shape.preset).toBe("rect");
    expect(shape.fill).toBe("FF0000");
    expect(shape.line).toEqual({ px: px(12700), color: "0000FF" });
    // customGeometry and stringified data members are skipped, not corrupted.
    expect(drawing.members).toHaveLength(2);
  });

  it("mirrors members of a flipH nested group within that group's box", () => {
    const { blocks } = projectDocumentOptions(
      doc([
        {
          paragraph: {
            children: [
              {
                wpgGroup: {
                  children: [
                    {
                      type: "wpg",
                      transformation: {
                        pixels: { x: 0, y: 0 },
                        emus: { x: 500000, y: 100000 },
                        offset: {
                          pixels: { x: 0, y: 0 },
                          emus: { x: 200000, y: 0 },
                        },
                        flip: { horizontal: true },
                      },
                      childOffset: { x: 0, y: 0 },
                      childExtent: { cx: 500000, cy: 100000 },
                      children: [
                        {
                          type: "wps",
                          transformation: {
                            pixels: { x: 0, y: 0 },
                            emus: { x: 250000, y: 100000 },
                            offset: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
                          },
                          data: {
                            presetGeometry: { preset: "rect" },
                            fill: { type: "none" },
                            children: [],
                          },
                        },
                      ],
                    },
                  ],
                  transformation: { width: 1905000, height: 190500 },
                  childOffset: { x: 100000, y: 0 },
                  childExtent: { cx: 1000000, cy: 100000 },
                },
              },
            ],
          },
        },
      ]),
    );
    const para = blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected paragraph");
    const member = para.drawings?.[0]?.members[0];
    if (member?.kind !== "shape") throw new Error("expected shape member");
    // Nested box: x = (200000-100000)×1.905 EMU = 20px, width = 100px. The
    // child's unmirrored [20,70) span reflects to [70,120) inside that box.
    expect(member.x).toBeCloseTo(70, 5);
    expect(member.width).toBeCloseTo(50, 5);
  });

  it("projects a wps text-box child's paragraphs through the style cascade", () => {
    const { blocks } = projectDocumentOptions(
      doc([
        {
          paragraph: {
            children: [
              {
                wpgGroup: {
                  children: [
                    {
                      type: "wps",
                      transformation: {
                        pixels: { x: 100, y: 10 },
                        emus: { x: 952500, y: 95250 },
                        offset: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
                      },
                      data: {
                        presetGeometry: { preset: "rect" },
                        fill: { type: "none" },
                        bodyProperties: { lIns: 0, tIns: 0, rIns: 0, bIns: 0, anchor: "center" },
                        children: [{ alignment: "right", children: [{ text: "XX项目" }] }],
                      },
                    },
                  ],
                  transformation: { width: 952500, height: 190500 },
                },
              },
            ],
          },
        },
      ]),
    );
    const para = blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected paragraph");
    const box = para.drawings?.[0]?.members[0];
    if (box?.kind !== "textBox") throw new Error("expected textBox member");
    expect(box.anchor).toBe("center");
    expect(box.insets).toEqual({ left: 0, top: 0, right: 0, bottom: 0 });
    // No chOff/chExt: children live in the group's own units (1:1).
    expect(box.width).toBeCloseTo(952500 / 9525, 5);
    const inner = box.blocks[0];
    if (inner?.kind !== "paragraph") throw new Error("expected projected paragraph");
    expect(inner.align).toBe("right");
    expect(inner.inline[0]).toMatchObject({ kind: "text", text: "XX项目" });
  });

  it("maps every relativeFrom axis and the align/percent position forms", () => {
    const anchorOf = (h: HorizontalPositionOptions, v: VerticalPositionOptions) => {
      const { blocks } = projectDocumentOptions(
        doc([
          {
            paragraph: {
              children: [
                {
                  wpgGroup: {
                    children: [],
                    transformation: { width: 9525, height: 9525 },
                    floating: { horizontalPosition: h, verticalPosition: v },
                  },
                },
              ],
            },
          },
        ]),
      );
      const para = blocks[0];
      if (para?.kind !== "paragraph") throw new Error("expected paragraph");
      return para.drawings?.[0]?.anchor;
    };
    // margin collapses onto the column axis; align rides along verbatim.
    expect(
      anchorOf({ relative: "margin", align: "center" }, { relative: "margin", align: "bottom" }),
    ).toEqual({
      horizontal: { relative: "column", align: "center" },
      vertical: { relative: "topMargin", align: "bottom" },
    });
    // page axis + percentOffset (thousandths of the reference extent).
    expect(
      anchorOf({ relative: "page", percentOffset: 50000 }, { relative: "page", offset: 9144 }),
    ).toEqual({
      horizontal: { relative: "page", percent: 50 },
      vertical: { relative: "page", offsetPx: 9144 / 9525 },
    });
    // Edge relatives + line→paragraph collapse.
    expect(anchorOf({ relative: "outsideMargin" }, { relative: "line" })).toEqual({
      horizontal: { relative: "rightMargin", offsetPx: 0 },
      vertical: { relative: "paragraph", offsetPx: 0 },
    });
  });

  it("flattens nested wpg groups through the composed child-space mapping", () => {
    const { blocks } = projectDocumentOptions(
      doc([
        {
          paragraph: {
            children: [
              {
                wpgGroup: {
                  children: [
                    {
                      type: "wpg",
                      transformation: {
                        pixels: { x: 50, y: 10 },
                        emus: { x: 476250, y: 95250 },
                        offset: { pixels: { x: 0, y: 0 }, emus: { x: 190500, y: 0 } },
                      },
                      childOffset: { x: 100000, y: 0 },
                      childExtent: { cx: 238125, cy: 95250 },
                      children: [
                        {
                          type: "wps",
                          transformation: {
                            pixels: { x: 50, y: 10 },
                            emus: { x: 238125, y: 95250 },
                            offset: { pixels: { x: 0, y: 0 }, emus: { x: 100000, y: 0 } },
                          },
                          data: {
                            presetGeometry: { preset: "line" },
                            outline: {
                              width: 9525,
                              type: "solidFill",
                              color: { value: "C00000" },
                              dash: "sysDash",
                            },
                            children: [],
                          },
                        },
                      ],
                    },
                  ],
                  transformation: { width: 952500, height: 95250 },
                  childOffset: { x: 0, y: 0 },
                  childExtent: { cx: 952500, cy: 95250 },
                },
              },
            ],
          },
        },
      ]),
    );
    const para = blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected paragraph");
    const [member] = para.drawings?.[0]?.members ?? [];
    // The nested line lands at its composed position: outer 20px + inner 0.
    if (member?.kind !== "path") throw new Error("expected path member");
    expect(member.x).toBeCloseTo(20, 5);
    expect(member.y).toBe(0);
    expect(member.width).toBeCloseTo(50, 5);
    expect(member.height).toBeCloseTo(10, 5);
    expect(member.d).toBe("M 0 0 L 50 10");
    expect(member.line).toMatchObject({ px: 1, color: "C00000", dash: "sysDash" });
  });

  it("projects custom geometry into a scaled SVG path member", () => {
    const { blocks } = projectDocumentOptions(
      doc([
        {
          paragraph: {
            children: [
              {
                wpgGroup: {
                  children: [
                    {
                      type: "wps",
                      transformation: {
                        pixels: { x: 100, y: 5 },
                        emus: { x: 952500, y: 47625 },
                        offset: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
                      },
                      data: {
                        customGeometry: {
                          pathList: [
                            {
                              w: 100,
                              h: 50,
                              commands: [
                                { command: "moveTo", point: { x: "0", y: "25" } },
                                {
                                  command: "cubicBezTo",
                                  points: [
                                    { x: "10", y: "0" },
                                    { x: "20", y: "50" },
                                    { x: "30", y: "25" },
                                  ],
                                },
                                { command: "lineTo", point: { x: "100", y: "25" } },
                                { command: "close" },
                              ],
                            },
                          ],
                        },
                        fill: { type: "none" },
                        outline: {
                          width: 15875,
                          type: "solidFill",
                          color: { value: "7D181C" },
                          cap: "round",
                          join: "round",
                        },
                        children: [],
                      },
                    },
                  ],
                  transformation: { width: 952500, height: 47625 },
                },
              },
            ],
          },
        },
      ]),
    );
    const para = blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected paragraph");
    const [member] = para.drawings?.[0]?.members ?? [];
    if (member?.kind !== "path") throw new Error("expected path member");
    // Path space 100×50 → box 100×5 px: y coordinates scale ×0.1.
    expect(member.d).toBe("M 0 2.5 C 10 0 20 5 30 2.5 L 100 2.5 Z");
    expect(member.line).toEqual({ px: 15875 / 9525, color: "7D181C", cap: "round", join: "round" });
  });

  it("resolves object-shaped colors, srcRect crops, and broken svg gradients", () => {
    const svg =
      '<svg xmlns="http://www.w3.org/2000/svg"><defs><linearGradient id="wps{x}@#db8c90@#7d181c">' +
      '<stop offset="0" stop-color="#db8c90"/><stop offset="1" stop-color="#7d181c"/>' +
      "</linearGradient></defs>" +
      '<path fill="url(#wps{x}@#db8c90@#7d181c)" d="M0 0"/></svg>';
    const { blocks } = projectDocumentOptions(
      doc([
        {
          paragraph: {
            children: [
              {
                wpgGroup: {
                  children: [
                    {
                      type: "svg",
                      fileName: "band.svg",
                      data: new TextEncoder().encode(svg),
                      fallback: {
                        type: "png",
                        fileName: "band.png",
                        data: new Uint8Array([105]),
                        transformation: {
                          pixels: { x: 10, y: 1 },
                          emus: { x: 95250, y: 9525 },
                          offset: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
                        },
                      },
                      transformation: {
                        pixels: { x: 10, y: 1 },
                        emus: { x: 95250, y: 9525 },
                        offset: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
                      },
                    },
                    {
                      type: "png",
                      fileName: "stripes.png",
                      data: new Uint8Array([105]),
                      transformation: {
                        pixels: { x: 10, y: 1 },
                        emus: { x: 95250, y: 9525 },
                        offset: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
                      },
                      // Raw ST_Percentage ints, as office-open's wpg picture
                      // parse emits them (12632 = 12.632%).
                      sourceRectangle: { left: 12632, top: 34824, bottom: 36285 },
                    },
                    {
                      type: "wps",
                      transformation: {
                        pixels: { x: 10, y: 1 },
                        emus: { x: 95250, y: 9525 },
                        offset: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
                      },
                      data: {
                        presetGeometry: { preset: "rect" },
                        fill: { type: "solid", color: { value: "7D181C" } },
                        children: [
                          {
                            children: [
                              {
                                text: "XX项目",
                                color: { val: "FFFFFF", themeColor: "background1" },
                              },
                            ],
                          },
                        ],
                      },
                    },
                  ],
                  transformation: { width: 952500, height: 95250 },
                },
              },
            ],
          },
        },
      ]),
    );
    const para = blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected paragraph");
    const [band, stripes, box] = para.drawings?.[0]?.members ?? [];
    if (band?.kind !== "picture" || typeof band.src !== "string") {
      throw new Error("expected picture member");
    }
    // The WPS gradient (stop average #ac5256) flattens; the def drops.
    const decoded = atob(band.src.split(",")[1] ?? "");
    expect(decoded).toContain("#ac5256");
    expect(decoded).not.toContain("linearGradient");
    if (stripes?.kind !== "picture") throw new Error("expected stripes member");
    const crop = stripes.crop as Record<string, number>;
    expect(crop.left).toBeCloseTo(0.12632, 5);
    expect(crop.top).toBeCloseTo(0.34824, 5);
    expect(crop.right).toBe(0);
    expect(crop.bottom).toBeCloseTo(0.36285, 5);
    if (box?.kind !== "textBox") throw new Error("expected textBox member");
    expect(box.blocks[0]).toMatchObject({
      kind: "paragraph",
      inline: [{ kind: "text", style: { color: "FFFFFF" } }],
    });
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

describe("projectDocumentOptions inline containers", () => {
  const paraOf = (children: SectionChild[]): LayoutBlock[] =>
    projectDocumentOptions(doc(children)).blocks;

  it("projects hyperlink containers as their runs with the Hyperlink look", () => {
    const blocks = paraOf([
      {
        paragraph: {
          children: [
            {
              hyperlink: {
                url: "https://example.com",
                children: [{ text: "linked", bold: true }, " and "],
              },
            },
            { text: "plain" },
          ],
        },
      },
    ]);
    const para = blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected paragraph");
    const [linked, , plain] = para.inline;
    // The run's own bold wins; color/underline come from Word's Hyperlink style.
    expect(linked).toMatchObject({
      kind: "text",
      text: "linked",
      style: { bold: true, underline: true, color: "0563C1" },
    });
    expect(plain).toMatchObject({ kind: "text", text: "plain", style: { underline: undefined } });
  });

  it("projects tracked insertions/deletions with Word's revision display", () => {
    const blocks = paraOf([
      {
        paragraph: {
          children: [
            {
              insertion: { id: 1, author: "A", date: "2026-01-01T00:00:00Z", children: ["added"] },
            },
            { text: " kept " },
            { deletion: { id: 2, author: "A", date: "2026-01-01T00:00:00Z", children: ["gone"] } },
          ],
        },
      },
    ]);
    const para = blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected paragraph");
    const [added, kept, gone] = para.inline;
    // First-author red: insertions underline, deletions strike (Word defaults).
    expect(added).toMatchObject({
      kind: "text",
      text: "added",
      style: { underline: true, strikethrough: undefined, color: "FF0000" },
    });
    expect(kept).toMatchObject({ kind: "text", text: " kept ", style: { color: undefined } });
    expect(gone).toMatchObject({
      kind: "text",
      text: "gone",
      style: { underline: undefined, strikethrough: true, color: "FF0000" },
    });
  });

  it("lets a run's explicit props beat the container preset", () => {
    const blocks = paraOf([
      {
        paragraph: {
          children: [
            {
              deletion: {
                id: 3,
                author: "A",
                date: "2026-01-01T00:00:00Z",
                children: [{ text: "red-strike", color: "008000" }],
              },
            },
          ],
        },
      },
    ]);
    const para = blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected paragraph");
    expect(para.inline[0]).toMatchObject({
      kind: "text",
      style: { strikethrough: true, color: "008000" },
    });
  });
});
