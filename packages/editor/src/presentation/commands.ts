// Ribbon command edits over the source deck JSON: slide/object insertion and
// text formatting. Pure functions — the element owns undo/redo bookkeeping
// and re-projection; these only mutate the model. Text bodies normalize their
// input sugar in place (body.text, string paragraphs, string runs) so the
// commands always walk real run objects, the same expansion stringify does.

import { EMU_PER_PX } from "@docen/layout";
import type { PresentationOptions, SlideChild, SlideOptions, TableCellOptions } from "@docen/pptx";
import type { ColorSchemeOptions, FontSchemeOptions } from "@office-open/core";
// Text/paragraph types come from @office-open/core/drawing — @docen/pptx's
// re-export surface doesn't carry them yet.
import type {
  BulletOptions,
  CustomGeometryOptions,
  NonVisualDrawingPropertiesOptions,
  ParagraphDescriptorOptions,
  ShapeType,
  StrikeStyle,
  TextAlignment,
  TextBodyOptions,
  TextBreakOptions,
  TextRunOptions,
  UnderlineStyle,
} from "@office-open/core/drawing";
import type { ChartOptions, ShapeOptions } from "@office-open/pptx";

/** Normalize the paragraphs in place: string items and paragraph-level text
 *  sugar both become real children. */
function paragraphsOf(paragraphs: ParagraphDescriptorOptions[]): ParagraphDescriptorOptions[] {
  for (let i = 0; i < paragraphs.length; i++) {
    let p = paragraphs[i]!;
    if (typeof p === "string") p = paragraphs[i] = { children: [{ text: p }] };
    // Paragraph-level text sugar (input-only) expands to a one-run paragraph.
    if (p.children === undefined && typeof p.text === "string") {
      p.children = [{ text: p.text }];
      delete p.text;
    }
  }
  return paragraphs;
}

/** The body's paragraphs as real objects, consuming the body's text sugar. */
export function bodyParagraphsOf(body: TextBodyOptions): ParagraphDescriptorOptions[] {
  const paragraphs = (body.paragraphs ??=
    body.text !== undefined ? [{ children: [{ text: body.text }] }] : []);
  delete body.text;
  return paragraphsOf(paragraphs as ParagraphDescriptorOptions[]);
}

/** The paragraph's text runs as real objects — string children expand;
 *  non-run children (breaks, fields) stay out of formatting's way. */
function runsOf(paragraph: ParagraphDescriptorOptions): TextRunOptions[] {
  const children = (paragraph.children ??= []);
  for (let i = 0; i < children.length; i++) {
    const child = children[i]!;
    if (typeof child === "string") children[i] = { text: child };
  }
  return children.filter(
    (child): child is TextRunOptions =>
      typeof child === "object" && "text" in child && typeof child.text === "string",
  );
}

/** Every real text run of the paragraphs, in order (the find engine walks
 *  these; the formatting commands format through the same walk). */
export function runsIn(paragraphs: ParagraphDescriptorOptions[]): TextRunOptions[] {
  const runs: TextRunOptions[] = [];
  for (const paragraph of paragraphsOf(paragraphs)) runs.push(...runsOf(paragraph));
  return runs;
}

/** Every real text run of a slide child — a shape's text body, or the cells
 *  of its table. Other child kinds carry no text. */
export function childRunsOf(child: SlideChild): TextRunOptions[] {
  if ("shape" in child) {
    return child.shape.textBody ? runsIn(bodyParagraphsOf(child.shape.textBody)) : [];
  }
  if ("table" in child) {
    const runs: TextRunOptions[] = [];
    for (const row of child.table.rows ?? []) {
      for (const cell of row.cells ?? []) runs.push(...runsIn(cellParagraphsOf(cell)));
    }
    return runs;
  }
  return [];
}

/** The child's cNvPr surface (name/hidden) — a table carries none. */
export function nonVisualOf(child: SlideChild): NonVisualDrawingPropertiesOptions | null {
  if ("shape" in child) return child.shape;
  if ("picture" in child) return child.picture;
  if ("table" in child) return child.table;
  if ("chart" in child)
    return {
      name: child.chart.name,
      description: child.chart.description,
      hidden: child.chart.hidden,
    };
  if ("line" in child) return child.line;
  if ("connector" in child) return child.connector;
  if ("group" in child) return child.group;
  if ("smartart" in child) return child.smartart;
  if ("video" in child) return child.video;
  if ("audio" in child) return child.audio;
  if ("ole" in child) return child.ole;
  return null;
}

/** The cell's paragraphs as real objects, consuming the cell's text sugar. */
export function cellParagraphsOf(cell: TableCellOptions): ParagraphDescriptorOptions[] {
  const paragraphs = (cell.children ??=
    cell.text !== undefined ? [{ children: [{ text: cell.text }] }] : []);
  delete cell.text;
  return paragraphsOf(paragraphs as ParagraphDescriptorOptions[]);
}

/** Word's toggle: a flag every run already carries is cleared, otherwise it
 *  is applied. No runs — no change. */
export function toggleRunFlag(
  paragraphs: ParagraphDescriptorOptions[],
  flag: "bold" | "italic",
): void {
  const runs = runsIn(paragraphs);
  if (runs.length === 0) return;
  const on = runs.every((run) => run[flag] === true);
  for (const run of runs) {
    if (on) delete run[flag];
    else run[flag] = true;
  }
}

/** Toggle an underline/strike style: every run carrying a live value clears,
 *  otherwise the "on" style lands. */
export function toggleRunStyle(
  paragraphs: ParagraphDescriptorOptions[],
  key: "underline",
  on: UnderlineStyle,
): void;
export function toggleRunStyle(
  paragraphs: ParagraphDescriptorOptions[],
  key: "strike",
  on: StrikeStyle,
): void;
export function toggleRunStyle(
  paragraphs: ParagraphDescriptorOptions[],
  key: "underline" | "strike",
  on: UnderlineStyle | StrikeStyle,
): void {
  const runs = runsIn(paragraphs);
  if (runs.length === 0) return;
  const off = key === "underline" ? "none" : "noStrike";
  const applied = runs.every((run) => {
    const value = run[key];
    return value !== undefined && value !== off;
  });
  if (key === "underline") {
    for (const run of runs) {
      if (applied) delete run.underline;
      else run.underline = on as UnderlineStyle;
    }
  } else {
    for (const run of runs) {
      if (applied) delete run.strike;
      else run.strike = on as StrikeStyle;
    }
  }
}

export function setRunFont(paragraphs: ParagraphDescriptorOptions[], font: string): void {
  for (const run of runsIn(paragraphs)) run.font = font;
}

/** The paragraph's live bullet (a real glyph or numbering — buNone and the
 *  absent default are both "off"). */
const liveBulletOf = (paragraph: ParagraphDescriptorOptions): boolean => {
  const bullet = paragraph.properties?.bullet;
  return bullet != null && bullet.type !== "none";
};

/** PowerPoint's list toggles: every paragraph carrying a live bullet clears,
 *  otherwise the default lands — the run-flag toggle at paragraph grain. */
function toggleBulletKind(paragraphs: ParagraphDescriptorOptions[], on: BulletOptions): void {
  const targets = paragraphsOf(paragraphs);
  if (targets.length === 0) return;
  const applied = targets.every(liveBulletOf);
  for (const paragraph of targets) {
    if (applied) delete paragraph.properties?.bullet;
    else paragraph.properties = { ...paragraph.properties, bullet: { ...on } };
  }
}

/** The Bullets button: the default round bullet (a:buChar, glyph rendered by
 *  the projection's default). */
export function toggleBullet(paragraphs: ParagraphDescriptorOptions[]): void {
  toggleBulletKind(paragraphs, { type: "char" });
}

/** The Numbering button: the default arabic-with-period autonumber scheme. */
export function toggleNumbering(paragraphs: ParagraphDescriptorOptions[]): void {
  toggleBulletKind(paragraphs, { type: "autoNum" });
}

/** Line spacing as a percentage over every paragraph (100 = single). */
export function setLineSpacingPercent(
  paragraphs: ParagraphDescriptorOptions[],
  percent: number,
): void {
  for (const paragraph of paragraphsOf(paragraphs)) {
    paragraph.properties = { ...paragraph.properties, lineSpacingPercent: percent };
  }
}

/** Run font size in points (the JSON's own unit). */
export function setRunSize(paragraphs: ParagraphDescriptorOptions[], size: number): void {
  for (const run of runsIn(paragraphs)) run.size = size;
}

/** Set every paragraph's horizontal alignment (PowerPoint's align buttons
 *  assign, they don't toggle). */
export function setParagraphAlignment(
  paragraphs: ParagraphDescriptorOptions[],
  alignment: TextAlignment,
): void {
  for (const paragraph of paragraphsOf(paragraphs)) {
    paragraph.properties = { ...paragraph.properties, alignment };
  }
}

/** Insert a blank slide after `at` (the PowerPoint new-slide position). */
export function insertSlideAt(deck: PresentationOptions, at: number): void {
  (deck.slides ??= []).splice(at + 1, 0, { children: [] });
}

/** Remove the slide at `at`; returns the removed slide, or null when `at`
 *  falls outside the deck. */
export function deleteSlideAt(deck: PresentationOptions, at: number): SlideOptions | null {
  const [removed] = (deck.slides ??= []).splice(at, 1);
  return removed ?? null;
}

/** Deep-clone the slide at `at` and insert the copy right after it; returns
 *  the copy's index, or -1 when `at` falls outside the deck. */
export function duplicateSlideAt(deck: PresentationOptions, at: number): number {
  const source = deck.slides?.[at];
  if (!source) return -1;
  const copy = structuredClone(source);
  deck.slides!.splice(at + 1, 0, copy);
  return at + 1;
}

/** Move the child at `from` to `to` within the slide's children (z order —
 *  later children paint on top). Both indices clamp into range. */
export function reorderChild(children: SlideChild[], from: number, to: number): void {
  const clamped = Math.max(0, Math.min(to, children.length - 1));
  const [moved] = children.splice(from, 1);
  if (moved) children.splice(clamped, 0, moved);
}

/** The paragraph's text: runs concatenated, a soft break (a:br) reading as a
 *  \n — the plain-lines view can't tell a break from a paragraph split, so a
 *  written-back break downgrades to one (same visible lines). */
function paragraphTextOf(paragraph: ParagraphDescriptorOptions): string {
  let text = "";
  for (const child of paragraph.children ?? []) {
    if (typeof child === "string") text += child;
    else if ("text" in child && typeof child.text === "string") text += child.text;
    else if ((child as TextBreakOptions).break === true) text += "\n";
  }
  return text;
}

/** The paragraphs' text as plain lines (joined by \n). */
function linesOf(paragraphs: ParagraphDescriptorOptions[]): string {
  return paragraphs.map((paragraph) => paragraphTextOf(paragraph)).join("\n");
}

/** One paragraph per line: each keeps the original paragraph's properties
 *  (alignment et al) and the style attributes of its first run. */
function paragraphsFromLines(
  previous: ParagraphDescriptorOptions[],
  lines: string[],
): ParagraphDescriptorOptions[] {
  return lines.map((line, i) => {
    const old = previous[i];
    const style = runsOf(previous[i] ?? {})[0];
    const run: TextRunOptions = { text: line };
    if (style) {
      const { text: _dropped, ...attrs } = style;
      Object.assign(run, attrs);
    }
    return { properties: old?.properties ? { ...old.properties } : undefined, children: [run] };
  });
}

/** The shape's text body as plain lines, or null when it carries no body. */
export function shapeTextOf(shape: ShapeOptions): string | null {
  if (!shape.textBody) return null;
  return linesOf(bodyParagraphsOf(shape.textBody));
}

/** Write plain lines back into the shape's text body. Returns false when the
 *  shape has no body (the caller's edit records nothing). */
export function writeShapeText(shape: ShapeOptions, text: string): boolean {
  if (!shape.textBody) return false;
  const body = shape.textBody;
  body.paragraphs = paragraphsFromLines(bodyParagraphsOf(body), text.split("\n"));
  delete body.text;
  return true;
}

/** The cell's text as plain lines. */
export function cellTextOf(cell: TableCellOptions): string {
  return linesOf(cellParagraphsOf(cell));
}

/** Write plain lines back into the cell (the shape side of the same
 *  contract — see writeShapeText). */
export function writeCellText(cell: TableCellOptions, text: string): void {
  cell.children = paragraphsFromLines(cellParagraphsOf(cell), text.split("\n"));
  delete cell.text;
}

/** The body's first run font size in points (PowerPoint's 18pt default) —
 *  the in-place text editor's only typographic nod. */
export function firstRunSizeOf(child: SlideChild): number {
  if (!("shape" in child) || !child.shape.textBody) return 18;
  for (const run of runsIn(bodyParagraphsOf(child.shape.textBody))) {
    if (typeof run.size === "number" && run.size > 0) return run.size;
  }
  return 18;
}

/** The cell's first run font size in points (the same 18pt fallback). */
export function firstCellRunSizeOf(cell: TableCellOptions): number {
  for (const paragraph of cellParagraphsOf(cell)) {
    for (const run of runsOf(paragraph)) {
      if (typeof run.size === "number" && run.size > 0) return run.size;
    }
  }
  return 18;
}

/** The slide's speaker notes as plain text (the string shorthand, or the
 *  structured notes slide's text sugar). */
export function slideNotesOf(slide: SlideOptions): string {
  return typeof slide.notes === "string" ? slide.notes : (slide.notes?.text ?? "");
}

/** Write the speaker notes back: the string shorthand for the plain case,
 *  the structured object's text sugar when the slide already carries one
 *  (stringify prefers that object's structured path, so the edit is
 *  best-effort there). */
export function writeSlideNotes(slide: SlideOptions, text: string): void {
  if (slide.notes && typeof slide.notes === "object") {
    if (text) slide.notes.text = text;
    else delete slide.notes.text;
  } else if (text) slide.notes = text;
  else delete slide.notes;
}

const emu = (px: number): number => Math.round(px * EMU_PER_PX);

/** A centered text-box shape for a slide of the given size — white fill and
 *  a hairline border keep an empty box visible on the canvas. `text` seeds
 *  the body (the paste path). */
export function makeTextBox(
  slideWidthPx: number,
  slideHeightPx: number,
  text = "",
  placement?: ShapePlacement,
): SlideChild {
  const width = placement?.w ?? slideWidthPx * 0.4;
  const height = placement?.h ?? slideHeightPx * 0.15;
  return {
    shape: {
      x: emu(placement?.x ?? (slideWidthPx - width) / 2),
      y: emu(placement?.y ?? (slideHeightPx - height) / 2),
      width: emu(width),
      height: emu(height),
      properties: { geometry: "rect", fill: "FFFFFF", outline: { width: "1pt", color: "808080" } },
      textBody: { text },
    },
  };
}

/** A centered WordArt box: Impact-style display text with a transparent
 *  frame so only the styled run is visible, like PowerPoint's fresh WordArt. */
export function makeWordArt(slideWidthPx: number, slideHeightPx: number, text: string): SlideChild {
  const width = slideWidthPx * 0.6;
  const height = slideHeightPx * 0.18;
  return {
    shape: {
      x: emu((slideWidthPx - width) / 2),
      y: emu((slideHeightPx - height) / 2),
      width: emu(width),
      height: emu(height),
      properties: { geometry: "rect", fill: { type: "none" } },
      textBody: {
        paragraphs: [
          {
            properties: { alignment: "center" },
            children: [{ text, bold: true, size: 54, font: "Impact", fill: "1F6CBD" }],
          },
        ],
      },
    },
  };
}

/** A centered, frameless symbol box. A Unicode font keeps the insertion
 *  honest across the symbol grid's arrows, math, currency and dingbats. */
export function makeSymbol(
  slideWidthPx: number,
  slideHeightPx: number,
  symbol: string,
): SlideChild {
  const size = Math.min(160, slideWidthPx * 0.18);
  return {
    shape: {
      x: emu((slideWidthPx - size) / 2),
      y: emu((slideHeightPx - size) / 2),
      width: emu(size),
      height: emu(size),
      properties: { geometry: "rect", fill: { type: "none" } },
      textBody: {
        paragraphs: [
          {
            properties: { alignment: "center" },
            children: [{ text: symbol, size: 72, font: "Arial Unicode MS" }],
          },
        ],
      },
    },
  };
}

/** PowerPoint's fresh 2"-square preset shape, centered, in the same white +
 *  hairline dressing as the text box (an empty text body keeps double-click
 *  text editing available). */
export function makeShape(
  slideWidthPx: number,
  slideHeightPx: number,
  geometry: ShapeType,
  placement?: ShapePlacement,
): SlideChild {
  const width = placement?.w ?? Math.min(192, slideWidthPx * 0.25);
  const height = placement?.h ?? Math.min(192, slideHeightPx * 0.25);
  return {
    shape: {
      x: emu(placement?.x ?? (slideWidthPx - width) / 2),
      y: emu(placement?.y ?? (slideHeightPx - height) / 2),
      width: emu(width),
      height: emu(height),
      properties: { geometry, fill: "FFFFFF", outline: { width: "1pt", color: "808080" } },
      textBody: { text: "" },
    },
  };
}

/** A drag-to-draw landing rectangle in slide px. */
export interface ShapePlacement {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** A straight line/connector from a drawn segment; direction is carried by
 *  the endpoint pair, matching the OOXML endpoint model. */
export function makeLine(preset: "line" | "straightConnector1", rect: ShapePlacement): SlideChild {
  return {
    line: {
      x1: emu(rect.x),
      y1: emu(rect.y),
      x2: emu(rect.x + rect.w),
      y2: emu(rect.y + rect.h),
      properties: { outline: { width: "1pt", color: "808080" } },
    },
  };
}

/** PowerPoint's pen stroke: one literal custGeom path, so the ink stays a
 *  vector shape and re-opens as a freeform instead of a rasterized picture. */
export function makePenStroke(points: readonly { x: number; y: number }[]): SlideChild | null {
  if (points.length < 2) return null;
  const x = Math.min(...points.map((point) => point.x));
  const y = Math.min(...points.map((point) => point.y));
  const width = Math.max(...points.map((point) => point.x)) - x;
  const height = Math.max(...points.map((point) => point.y)) - y;
  const boxWidth = emu(Math.max(width, 1));
  const boxHeight = emu(Math.max(height, 1));
  const local = points.map((point) => ({
    x: emu(point.x - x),
    y: emu(point.y - y),
  }));
  const customGeometry: CustomGeometryOptions = {
    pathList: [
      {
        w: boxWidth,
        h: boxHeight,
        fill: "none",
        stroke: true,
        extrusionOk: false,
        commands: [
          { command: "moveTo", point: { x: String(local[0]!.x), y: String(local[0]!.y) } },
          ...local.slice(1).map((point) => ({
            command: "lineTo" as const,
            point: { x: String(point.x), y: String(point.y) },
          })),
        ],
      },
    ],
  };
  return {
    shape: {
      x: emu(x),
      y: emu(y),
      width: boxWidth,
      height: boxHeight,
      properties: {
        customGeometry,
        fill: { type: "none" },
        outline: { width: 12700, color: "262626" },
      },
    },
  };
}

/** A footer field box (the a:fld token plus its cached display text — the
 *  file's field re-evaluates per slide in a real renderer). Anchored bottom
 *  corner: `align: "right"` hugs the right edge, anything else the left. */
export function makeFieldBox(
  slideWidthPx: number,
  slideHeightPx: number,
  type: string,
  text: string,
  align: "left" | "right",
): SlideChild {
  const width = slideWidthPx * 0.2;
  const height = Math.round(0.4 * 96);
  const margin = Math.round(0.25 * 96);
  const x = align === "right" ? slideWidthPx - width - margin : margin;
  return {
    shape: {
      x: emu(x),
      y: emu(slideHeightPx - height - margin),
      width: emu(width),
      height: emu(height),
      properties: { geometry: "rect", fill: { type: "none" } },
      textBody: {
        paragraphs: [
          {
            properties: { alignment: align },
            children: [{ type, text }],
          },
        ],
      },
    },
  };
}

/** A picture child centered on a slide of the given size, scaled down to fit
 *  within half the slide (never upscaled). */
export function makePicture(
  slideWidthPx: number,
  slideHeightPx: number,
  naturalWidth: number,
  naturalHeight: number,
  data: string,
  type: "png" | "jpg" | "gif" | "bmp",
): SlideChild {
  const scale = Math.min(slideWidthPx / 2 / naturalWidth, slideHeightPx / 2 / naturalHeight, 1);
  const width = naturalWidth * scale;
  const height = naturalHeight * scale;
  return {
    picture: {
      x: emu((slideWidthPx - width) / 2),
      y: emu((slideHeightPx - height) / 2),
      width: emu(width),
      height: emu(height),
      data,
      type,
    },
  };
}

/** PowerPoint's embedded-object icon: source bytes stay an OLE embed, while
 *  the generated PNG is only the frame's `p:pic` preview. */
const OBJECT_PROG_IDS: ReadonlyMap<string, string> = new Map([
  ["doc", "Word.Document.8"],
  ["docx", "Word.Document.12"],
  ["xls", "Excel.Sheet.8"],
  ["xlsx", "Excel.Sheet.12"],
  ["ppt", "PowerPoint.Show.8"],
  ["pptx", "PowerPoint.Show.12"],
  ["pdf", "Acrobat.Document.DC"],
  ["txt", "txtfile"],
  ["zip", "CompressedFolder"],
]);

/** The registered OLE server most closely associated with a source file. */
export function objectProgIdOf(sourceName: string): string {
  const extension = sourceName.split(".").pop()?.toLowerCase() ?? "";
  return OBJECT_PROG_IDS.get(extension) ?? "Package";
}

export function makeObject(
  slideWidthPx: number,
  slideHeightPx: number,
  data: Uint8Array,
  iconData: string,
  sourceName: string,
  progId = "Package",
): SlideChild {
  const size = 96;
  return {
    ole: {
      x: emu((slideWidthPx - size) / 2),
      y: emu((slideHeightPx - size) / 2),
      width: emu(size),
      height: emu(size),
      name: sourceName,
      progId,
      showAsIcon: true,
      imageWidth: emu(size),
      imageHeight: emu(size),
      embed: { data },
      iconImage: { data: iconData, type: "png" },
    },
  };
}

/** PowerPoint's fresh 3×3 table: centered, 60% of the slide wide, equal
 *  columns, each row PowerPoint's fresh 0.35" height. The GUID is the
 *  built-in Medium Style 2 — Accent 1 that PowerPoint applies by default. */
export function makeTable(slideWidthPx: number, slideHeightPx: number): SlideChild {
  const rows = 3;
  const cols = 3;
  const width = slideWidthPx * 0.6;
  const rowHeight = Math.round(0.35 * 96);
  return {
    table: {
      x: emu((slideWidthPx - width) / 2),
      y: emu((slideHeightPx - rowHeight * rows) / 2),
      width: emu(width),
      height: emu(rowHeight * rows),
      firstRow: true,
      columnWidths: Array.from({ length: cols }, () => emu(width / cols)),
      tableStyleId: "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
      rows: Array.from({ length: rows }, () => ({
        height: emu(rowHeight),
        cells: Array.from({ length: cols }, () => ({ text: "" })),
      })),
    },
  };
}

/** A fresh four-step process SmartArt, centered at PowerPoint's common 60% ×
 *  48% frame. The nodes use the file's real data model; rendering starts with
 *  the stable painter fallback and can grow into full DGM layout later. */
export function makeSmartArt(slideWidthPx: number, slideHeightPx: number): SlideChild {
  const width = slideWidthPx * 0.6;
  const height = slideHeightPx * 0.48;
  return {
    smartart: {
      x: emu((slideWidthPx - width) / 2),
      y: emu((slideHeightPx - height) / 2),
      width: emu(width),
      height: emu(height),
      layout: "process1",
      nodes: [{ text: "Discover" }, { text: "Design" }, { text: "Build" }, { text: "Launch" }],
    },
  };
}

/** PowerPoint's fresh chart: a clustered column with the sample quarterly
 *  series, centered at the same 60% × 48% frame the other inserts share.
 *  The payload is the real ChartOptions model — parse and export both read
 *  it verbatim, so the chart round-trips like any hand-authored one. */
export function makeChart(slideWidthPx: number, slideHeightPx: number): SlideChild {
  const width = slideWidthPx * 0.6;
  const height = slideHeightPx * 0.48;
  const chart: ChartOptions = {
    x: emu((slideWidthPx - width) / 2),
    y: emu((slideHeightPx - height) / 2),
    width: emu(width),
    height: emu(height),
    type: "column",
    title: "Chart Title",
    categories: ["Category 1", "Category 2", "Category 3", "Category 4"],
    series: [
      { name: "Series 1", values: [32, 28, 41, 36] },
      { name: "Series 2", values: [18, 34, 25, 29] },
    ],
    showLegend: true,
  };
  return { chart };
}

/** A built-in theme preset: the real Office palette (the two shipped
 *  defaults every Office install carries) plus its font pair — the model a
 *  theme application writes into the master. */
export interface ThemePreset {
  id: string;
  colorScheme: ColorSchemeOptions;
  fontScheme?: FontSchemeOptions;
}

const officeFontPair = (major: string, minor: string): FontSchemeOptions => ({
  majorFont: { latin: { typeface: major } },
  minorFont: { latin: { typeface: minor } },
});

/** The Office palettes, verbatim from the theme parts Office ships. */
export const THEME_PRESETS: readonly ThemePreset[] = [
  {
    id: "office",
    colorScheme: {
      name: "Office",
      dark1: "000000",
      light1: "FFFFFF",
      dark2: "44546A",
      light2: "E7E6E6",
      accent1: "4472C4",
      accent2: "ED7D31",
      accent3: "A5A5A5",
      accent4: "FFC000",
      accent5: "5B9BD5",
      accent6: "70AD47",
      hyperlink: "0563C1",
      followedHyperlink: "954F72",
    },
    fontScheme: officeFontPair("Calibri Light", "Calibri"),
  },
  {
    id: "office-classic",
    colorScheme: {
      name: "Office 2007-2010",
      dark1: "000000",
      light1: "FFFFFF",
      dark2: "1F497D",
      light2: "EEECE1",
      accent1: "4F81BD",
      accent2: "C0504D",
      accent3: "9BBB59",
      accent4: "8064A2",
      accent5: "4BACC6",
      accent6: "F79646",
      hyperlink: "096B9E",
      followedHyperlink: "4F81BD",
    },
    fontScheme: officeFontPair("Cambria", "Calibri"),
  },
];

const ACCENT_SLOTS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const;

/** The theme variants for one scheme: the accents rotated so each accent
 *  leads in turn — the same recoloring role PowerPoint's Variants gallery
 *  plays for the current theme, derived mechanically so every deck gets the
 *  four variants its own colors define. */
export function variantSchemesOf(scheme: ColorSchemeOptions): ColorSchemeOptions[] {
  const accents = ACCENT_SLOTS.map((slot) => scheme[slot]);
  return ACCENT_SLOTS.slice(0, 4).map((_, lead) => {
    const rotated = accents.map((_, index) => accents[(index + lead) % accents.length]);
    return {
      ...scheme,
      name: `${scheme.name ?? "Custom"} ${lead + 1}`,
      ...Object.fromEntries(ACCENT_SLOTS.map((slot, index) => [slot, rotated[index]])),
    } as ColorSchemeOptions;
  });
}

/** One placeholder child Reset moved, with the undo pair's payloads: the
 *  snapshot Reset returns to, and the inherited geometry it applies. */
export interface PlaceholderReset {
  child: SlideChild & { shape: ShapeOptions };
  before: Partial<Pick<ShapeOptions, "x" | "y" | "width" | "height">>;
  after: Partial<Pick<ShapeOptions, "x" | "y" | "width" | "height">>;
}

/** The layout's placeholder shapes by their type token — Reset's first
 *  source. Placeholders without their own xfrm inherit the master's, so a
 *  miss here falls through to {@link masterPlaceholdersOf}. */
export function layoutPlaceholdersOf(layout: SlideOptions): Map<string, ShapeOptions> {
  return placeholdersOf(layout.children ?? []);
}

/** The master's placeholder shapes — the inheritance chain's resolved
 *  geometry, Reset's fallback when the layout placeholder has no xfrm. */
export function masterPlaceholdersOf(
  master: NonNullable<PresentationOptions["masters"]>[number],
): Map<string, ShapeOptions> {
  return placeholdersOf(master.children ?? []);
}

function placeholdersOf(children: SlideChild[]): Map<string, ShapeOptions> {
  const map = new Map<string, ShapeOptions>();
  for (const child of children) {
    if ("shape" in child && typeof child.shape?.placeholder === "string")
      map.set(child.shape.placeholder, child.shape);
  }
  return map;
}

/** Word's Reset over one slide: every placeholder child (walking groups)
 *  takes the inherited position/size from the layout's placeholder, or the
 *  master's when the layout has none — text and all other children stay
 *  untouched. Returns the moves; an empty result means nothing inherited. */
export function resetSlidePlaceholders(
  slide: SlideOptions,
  layout: SlideOptions | undefined,
  master: NonNullable<PresentationOptions["masters"]>[number] | undefined,
): PlaceholderReset[] {
  const layoutMap = layout ? layoutPlaceholdersOf(layout) : new Map<string, ShapeOptions>();
  const masterMap = master ? masterPlaceholdersOf(master) : new Map<string, ShapeOptions>();
  const sourceOf = (type: string): ShapeOptions | undefined => {
    const fromLayout = layoutMap.get(type);
    if (fromLayout?.x !== undefined && fromLayout?.y !== undefined) return fromLayout;
    const fromMaster = masterMap.get(type);
    if (fromMaster?.x !== undefined && fromMaster?.y !== undefined) return fromMaster;
    return undefined;
  };
  const moved: PlaceholderReset[] = [];
  const walk = (children: SlideChild[]): void => {
    for (const child of children) {
      if ("group" in child) {
        walk(child.group.children ?? []);
        continue;
      }
      if (!("shape" in child) || typeof child.shape.placeholder !== "string") continue;
      const source = sourceOf(child.shape.placeholder);
      if (!source) continue;
      const before = {
        x: child.shape.x,
        y: child.shape.y,
        width: child.shape.width,
        height: child.shape.height,
      };
      child.shape.x = source.x;
      child.shape.y = source.y;
      if (source.width !== undefined) child.shape.width = source.width;
      if (source.height !== undefined) child.shape.height = source.height;
      moved.push({
        child,
        before,
        after: { x: source.x, y: source.y, width: source.width, height: source.height },
      });
    }
  };
  walk(slide.children ?? []);
  return moved;
}

/** A centered media frame from a browser file read. PowerPoint keeps native
 *  bytes in the package; this inserts the same source model so rendering and
 *  export share one payload. */
export function makeMediaFrame(
  slideWidthPx: number,
  slideHeightPx: number,
  media: "video" | "audio",
  data: Uint8Array,
  type: "mp4" | "mov" | "wmv" | "avi" | "mp3" | "wav" | "wma" | "aac",
  fileName?: string,
): SlideChild {
  const width = slideWidthPx * (media === "video" ? 0.5 : 0.3);
  const height = media === "video" ? Math.round((width * 9) / 16) : Math.round(80);
  const frame = {
    x: emu((slideWidthPx - width) / 2),
    y: emu((slideHeightPx - height) / 2),
    width: emu(width),
    height: emu(height),
    data,
    type,
    ...(fileName ? { fileName } : {}),
  };
  return (media === "video" ? { video: frame } : { audio: frame }) as SlideChild;
}
