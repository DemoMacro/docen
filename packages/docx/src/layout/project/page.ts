// Section-level projection: the body-child dispatch (placeholders for
// unprojectable shapes), the page-break paragraph split, the flow geometry,
// and the page furniture/background/borders a section paints with.

import {
  twipToPx,
  type LayoutBlock,
  type ProjectedColumns,
  type ProjectedFlowBox,
  type ProjectedLineNumbers,
  type ProjectedPageBackground,
  type ProjectedPageBorder,
  type ProjectedPageBorders,
  type ProjectedPageFurniture,
} from "@docen/layout";
import type { DocumentOptions, SectionChild, SectionOptions } from "@office-open/docx";

import { resolvePageSize } from "../../extensions/utils";
import { indexCharacterStyles } from "../../style-cascade";
import type { ProjectContext } from "./context";
import {
  PLACEHOLDER_PX,
  isRecord,
  measureTwip,
  num,
  str,
  type BodyParagraph,
  type Rec,
} from "./guards";
import { indexNumberings } from "./numbering";
import { projectParagraph } from "./paragraph";
import { projectTable } from "./table";

// ── section-child dispatch ──

/** One body child → layout block. Unprojectable shapes become placeholder
 *  boxes; zero-height markers (bookmarks) vanish — both OOXML-faithful. */
export function projectChild(
  child: SectionChild,
  ctx: ProjectContext,
): LayoutBlock | LayoutBlock[] | null {
  if ("paragraph" in child) return projectParagraphBlocks(child.paragraph, ctx);
  if ("table" in child) return projectTable(child.table, ctx);
  if ("toc" in child) return projectToc(child.toc, ctx);
  if ("bookmarkStart" in child || "bookmarkEnd" in child) return null;
  // sdt, textbox, altChunk, customXml, rawXml → a labeled box.
  const label = Object.keys(child)[0];
  return { kind: "placeholder", heightPx: PLACEHOLDER_PX, label };
}

/** A rendered TOC is plain paragraphs (TOC1-9 styles, tab + page number) —
 *  Word lays entries out exactly so. Each entry's paragraph carries its own
 *  style/tab stops, and the HYPERLINK fields project as their cached result
 *  text, so the entries flow as real blocks with real line heights. An
 *  unexpanded field (no entries) stays a placeholder — Word shows a field
 *  result there, not blank space. */
function projectToc(toc: unknown, ctx: ProjectContext): LayoutBlock | LayoutBlock[] | null {
  if (!isRecord(toc) || !Array.isArray(toc.entries) || toc.entries.length === 0) {
    return { kind: "placeholder", heightPx: PLACEHOLDER_PX, label: "toc" };
  }
  const blocks: LayoutBlock[] = [];
  for (const entry of toc.entries) {
    if (!isRecord(entry) || !isRecord(entry.paragraph)) continue;
    blocks.push(projectParagraph(entry.paragraph as BodyParagraph, ctx));
  }
  return blocks.length > 0
    ? blocks
    : { kind: "placeholder", heightPx: PLACEHOLDER_PX, label: "toc" };
}

/** A run-level page break (w:br type=page) or column break (w:br
 *  type=column) splits its paragraph: the flow engine consumes pageBreak /
 *  columnBreak blocks, so the paragraph is re-emitted around each break with
 *  its properties intact (Word keeps the paragraph running onto the next
 *  page/column). An empty chunk flushes to nothing: the paragraph mark rides
 *  on the break's own line (Word shows "———page break———¶" as one row), so a
 *  trailing break must not leave an empty paragraph behind on the next page. */
export function projectParagraphBlocks(
  p: BodyParagraph,
  ctx: ProjectContext,
): LayoutBlock | LayoutBlock[] {
  const runs: readonly unknown[] =
    typeof p === "string" ? [p] : (p?.children ?? (p?.text != null ? [p.text] : []));
  if (!runs.some((run) => isRecord(run) && (run.pageBreak === true || run.columnBreak === true))) {
    return projectParagraph(p, ctx);
  }
  const out: LayoutBlock[] = [];
  let chunk: unknown[] = [];
  const flush = (): void => {
    if (chunk.length === 0) return;
    out.push(
      projectParagraph({ ...(isRecord(p) ? p : {}), children: chunk } as BodyParagraph, ctx),
    );
    chunk = [];
  };
  for (const run of runs) {
    if (isRecord(run) && run.pageBreak === true) {
      flush();
      out.push({ kind: "pageBreak" });
    } else if (isRecord(run) && run.columnBreak === true) {
      flush();
      out.push({ kind: "columnBreak" });
    } else {
      chunk.push(run);
    }
  }
  flush();
  return out;
}

// ── section flow geometry ──

export function projectFlowBox(properties: unknown): ProjectedFlowBox {
  const sp: Rec = isRecord(properties) ? properties : {};
  const { width, height } = resolvePageSize(sp.pageSize);
  const m: Rec = isRecord(sp.pageMargin) ? sp.pageMargin : {};
  const side = (v: unknown, d: number): number => twipToPx(measureTwip(v) ?? d);
  const top = side(m.top, 1440);
  const bottom = side(m.bottom, 1440);
  const left = side(m.left, 1800);
  const right = side(m.right, 1800);
  const grid: Rec = isRecord(sp.grid) ? sp.grid : {};
  const pitchTw = measureTwip(grid.linePitch);
  const linePitchPx =
    grid.type && grid.type !== "default" && pitchTw && pitchTw > 0 ? twipToPx(pitchTw) : undefined;
  return {
    pageWidthPx: twipToPx(width),
    pageHeightPx: twipToPx(height),
    contentWidthPx: twipToPx(width) - left - right,
    contentHeightPx: twipToPx(height) - top - bottom,
    contentLeftPx: left,
    contentTopPx: top,
    linePitchPx,
  };
}

/** Project a section's headers/footers. An absent slot stays undefined (the
 *  painter falls back per OOXML: page 1 without titlePage and even pages
 *  without evenAndOddHeaders both use `default`). */
export function projectPageFurniture(
  section: SectionOptions | undefined,
  doc: DocumentOptions,
  prevFurniture?: ProjectedPageFurniture,
): ProjectedPageFurniture {
  const ctx: ProjectContext = {
    styles: doc.styles,
    characterStyles: indexCharacterStyles(doc.styles),
    numberings: indexNumberings(doc.numbering),
    listCounters: new Map(),
    openComments: new Set(),
    // Word forbids footnote references in headers/footers — a fresh counter
    // keeps the furniture walk independent even if malformed input carries one.
    footnoteOrdinals: new Map(),
    endnoteOrdinals: new Map(),
  };
  const projectSlots = (side: unknown): LayoutBlock[] | undefined => {
    if (!Array.isArray(side)) return undefined;
    const blocks: LayoutBlock[] = [];
    for (const child of side) {
      const block = projectChild(child, ctx);
      if (Array.isArray(block)) blocks.push(...block);
      else if (block) blocks.push(block);
    }
    return blocks.length > 0 ? blocks : undefined;
  };
  const props: Rec = isRecord(section?.properties) ? section.properties : {};
  const margin: Rec = isRecord(props.pageMargin) ? props.pageMargin : {};
  const explicitHeader = projectSlots(section?.headers?.default);
  const explicitFirstHeader = projectSlots(section?.headers?.first);
  const explicitEvenHeader = projectSlots(section?.headers?.even);
  const explicitFooter = projectSlots(section?.footers?.default);
  const explicitFirstFooter = projectSlots(section?.footers?.first);
  const explicitEvenFooter = projectSlots(section?.footers?.even);

  return {
    header: explicitHeader ?? prevFurniture?.header,
    firstHeader: explicitFirstHeader ?? prevFurniture?.firstHeader,
    evenHeader: explicitEvenHeader ?? prevFurniture?.evenHeader,
    footer: explicitFooter ?? prevFurniture?.footer,
    firstFooter: explicitFirstFooter ?? prevFurniture?.firstFooter,
    evenFooter: explicitEvenFooter ?? prevFurniture?.evenFooter,
    titlePage: props.titlePage === true,
    evenAndOddHeaders: doc.settings?.evenAndOddHeaders === true || props.evenAndOddHeaders === true,
    headerDistancePx: twipToPx(measureTwip(margin.header) ?? 720),
    footerDistancePx: twipToPx(measureTwip(margin.footer) ?? 720),
  };
}

/** Project w:background. A v:fill pattern's 1bpp hatch tile paints threads
 *  (set bits) in the fill's color2 and gaps (clear bits) in w:color — but
 *  Word's own rasterization smooths the weave into a near-flat tint (the PDF
 *  reference shows the same bit coverage with far weaker periodicity), and a
 *  1:1 dot screen on the canvas reads as harsh yellow speckle against it.
 *  Average the two colors by bit coverage into one flat color instead. */
export function projectPageBackground(doc: DocumentOptions): ProjectedPageBackground | undefined {
  const bg = doc.background as
    | {
        color?: string;
        rawXml?: string;
        rawMedia?: Array<{ fileName?: string; type?: string; data?: Uint8Array }>;
      }
    | undefined;
  if (!bg) return undefined;
  const raw = bg.rawXml ?? "";
  const hexOf = (m: RegExpMatchArray | null): string | undefined =>
    m ? m[1].toUpperCase() : undefined;
  // The structured color is the primary source (a plain w:background @w:color
  // parses there and never round-trips a rawXml); the verbatim XML arm is the
  // pattern-fill fallback.
  const structured =
    typeof bg.color === "string" && bg.color !== "auto"
      ? bg.color.replace("#", "").toUpperCase()
      : undefined;
  const color = structured ?? hexOf(raw.match(/<w:background[^>]*\sw:color="([0-9A-Fa-f]{6})"/));
  const out: ProjectedPageBackground = color ? { color } : {};
  const fill = raw.match(/<v:fill[^>]*type="pattern"[^>]*>/);
  const rid = fill?.[0].match(/\sr:id="\{?([^"}]+)\}?"/)?.[1];
  const media =
    (rid ? bg.rawMedia?.find((m) => m.fileName === rid) : undefined) ??
    bg.rawMedia?.find((m) => m.type === "bmp");
  const data = media?.data;
  if (!fill || !data) return Object.keys(out).length > 0 ? out : undefined;
  if (
    data.length < 62 ||
    data[0] !== 0x42 ||
    data[1] !== 0x4d // "BM" — a complete BMP file
  ) {
    return out;
  }
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const headerSize = view.getUint32(14, true);
  const width = view.getInt32(18, true);
  const height = Math.abs(view.getInt32(22, true));
  const bpp = view.getUint16(28, true);
  const clrUsed = view.getUint32(46, true) || (bpp <= 8 ? 1 << bpp : 0);
  const dataAt = view.getUint32(10, true);
  if (
    bpp !== 1 ||
    clrUsed !== 2 ||
    width <= 0 ||
    height <= 0 ||
    width > 64 ||
    height > 64 ||
    dataAt < 14 + headerSize + 8 ||
    dataAt + Math.ceil(width / 8) * height > data.length
  ) {
    return out;
  }
  // Count the set bits (threads): rows are padded to 4-byte boundaries.
  let ones = 0;
  const rowBytes = Math.ceil(width / 8);
  const rowStride = Math.ceil(rowBytes / 4) * 4;
  for (let y = 0; y < height; y++) {
    for (let b = 0; b < rowBytes; b++) {
      const byte = view.getUint8(dataAt + y * rowStride + b);
      for (let bit = 0; bit < 8 && b * 8 + bit < width; bit++) {
        ones += (byte >> (7 - bit)) & 1;
      }
    }
  }
  const coverage = ones / (width * height);
  const thread = hexOf(fill[0].match(/\scolor2="#?([0-9A-Fa-f]{6})"/)) ?? "000000";
  const gap = color ?? "FFFFFF";
  const chan = (t: string, g: string, i: number): number =>
    Math.round(
      coverage * parseInt(thread.slice(i * 2, i * 2 + 2), 16) +
        (1 - coverage) * parseInt(gap.slice(i * 2, i * 2 + 2), 16),
    );
  return {
    color: [0, 1, 2]
      .map((i) => chan(thread, gap, i).toString(16).padStart(2, "0"))
      .join("")
      .toUpperCase(),
  };
}

/** One side of the projected w:pgBorders (see {@link ProjectedPageBorders}). */
function projectPageBorderSide(v: unknown): ProjectedPageBorder | undefined {
  if (!isRecord(v)) return undefined;
  const style = str(v.style);
  // nil/none explicitly paint nothing; an absent side is simply not rendered.
  if (!style || style === "nil" || style === "none") return undefined;
  const size = num(v.size);
  return {
    style,
    // w:sz counts 1/8 pt; Word's default 4 = 0.5 pt ≈ 0.67 px.
    widthPx: size != null ? (size / 8) * (96 / 72) : (4 / 8) * (96 / 72),
    color: str(v.color),
    spacePt: num(v.space),
  };
}

/** Project a section's w:pgBorders for painting: per-side strokes, the pages
 *  that paint them, and the offset reference (page edge vs text margin). */
export function projectPageBorders(properties: unknown): ProjectedPageBorders | undefined {
  const raw =
    isRecord(properties) && isRecord(properties.pageBorders)
      ? (properties.pageBorders as Rec)
      : null;
  if (!raw) return undefined;
  const borders: ProjectedPageBorders = {
    display: str(raw.display) as ProjectedPageBorders["display"],
    offsetFrom: str(raw.offsetFrom) as ProjectedPageBorders["offsetFrom"],
    behind: raw.zOrder === "back" || undefined,
    top: projectPageBorderSide(raw.top),
    right: projectPageBorderSide(raw.right),
    bottom: projectPageBorderSide(raw.bottom),
    left: projectPageBorderSide(raw.left),
  };
  const hasSide = borders.top || borders.right || borders.bottom || borders.left;
  return hasSide ? borders : undefined;
}

/** Project a section's w:lnNumType for painting: the count stride, the
 *  restart number, what resets the counter, and the margin gap
 *  (w:distance twips → px). */
export function projectLineNumbers(properties: unknown): ProjectedLineNumbers | undefined {
  const raw =
    isRecord(properties) && isRecord(properties.lineNumberType)
      ? (properties.lineNumberType as Rec)
      : null;
  if (!raw) return undefined;
  return {
    countBy: num(raw.countBy) ?? 1,
    start: num(raw.start) ?? 1,
    restart: raw.restart === "continuous" || raw.restart === "newSection" ? raw.restart : "newPage",
    distancePx: twipToPx(measureTwip(raw.distance) ?? 0),
  };
}

/** Project a section's w:cols for the flow: the column count, the gap
 *  (w:space twips → px, Word's default 720), the separator line flag, and
 *  explicit per-column widths (w:col children, twips → px) when the widths
 *  are unequal. Absent/1-column sections project to undefined. */
export function projectColumns(properties: unknown): ProjectedColumns | undefined {
  const raw =
    isRecord(properties) && isRecord(properties.columns) ? (properties.columns as Rec) : null;
  if (!raw) return undefined;
  const count = num(raw.count) ?? 1;
  if (count <= 1) return undefined;
  const children = Array.isArray(raw.children)
    ? raw.children.filter(isRecord).map((col) => twipToPx(measureTwip(col.width) ?? 0))
    : [];
  return {
    count,
    spacePx: twipToPx(measureTwip(raw.space) ?? 720),
    separate: raw.separate === true,
    equalWidth: raw.equalWidth !== false,
    columnsPx: raw.equalWidth === false && children.length > 0 ? children : undefined,
  };
}
