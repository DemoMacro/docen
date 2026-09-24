// Ribbon command edits over the source deck JSON: slide/object insertion and
// text formatting. Pure functions — the element owns undo/redo bookkeeping
// and re-projection; these only mutate the model. Text bodies normalize their
// input sugar in place (body.text, string paragraphs, string runs) so the
// commands always walk real run objects, the same expansion stringify does.

import { EMU_PER_PX } from "@docen/layout";
import type { PresentationOptions, SlideChild, SlideOptions, TableCellOptions } from "@docen/pptx";
// Text/paragraph types come from @office-open/core/drawing — @docen/pptx's
// re-export surface doesn't carry them yet.
import type {
  BulletOptions,
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
  if ("line" in child) return child.line;
  if ("connector" in child) return child.connector;
  if ("group" in child) return child.group;
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
  for (const run of runs) {
    if (applied) delete run[key];
    else run[key] = on as never;
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

/** The shape's text body as plain lines, or null when the child carries no
 *  text body. */
export function shapeTextOf(child: SlideChild): string | null {
  if (!("shape" in child) || !child.shape.textBody) return null;
  return linesOf(bodyParagraphsOf(child.shape.textBody));
}

/** Write plain lines back into the shape's text body. Returns false when the
 *  child has no text body (the caller's edit records nothing). */
export function writeShapeText(child: SlideChild, text: string): boolean {
  if (!("shape" in child) || !child.shape.textBody) return false;
  const body = child.shape.textBody;
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
export function makeTextBox(slideWidthPx: number, slideHeightPx: number, text = ""): SlideChild {
  const width = slideWidthPx * 0.4;
  const height = slideHeightPx * 0.15;
  return {
    shape: {
      x: emu((slideWidthPx - width) / 2),
      y: emu((slideHeightPx - height) / 2),
      width: emu(width),
      height: emu(height),
      properties: { geometry: "rect", fill: "FFFFFF", outline: { width: "1pt", color: "808080" } },
      textBody: { text },
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
): SlideChild {
  const width = Math.min(192, slideWidthPx * 0.25);
  const height = Math.min(192, slideHeightPx * 0.25);
  return {
    shape: {
      x: emu((slideWidthPx - width) / 2),
      y: emu((slideHeightPx - height) / 2),
      width: emu(width),
      height: emu(height),
      properties: { geometry, fill: "FFFFFF", outline: { width: "1pt", color: "808080" } },
      textBody: { text: "" },
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

/** PowerPoint's fresh 3×3 table: centered, 60% of the slide wide, equal
 *  columns, each row PowerPoint's fresh 0.35" height. Every edge carries a
 *  hairline border — an unstyled grid would paint invisible (the deck's
 *  table style lives in the theme part the editor doesn't resolve). */
export function makeTable(slideWidthPx: number, slideHeightPx: number): SlideChild {
  const rows = 3;
  const cols = 3;
  const width = slideWidthPx * 0.6;
  const rowHeight = Math.round(0.35 * 96);
  const edge = { width: "1pt", color: "808080" } as const;
  return {
    table: {
      x: emu((slideWidthPx - width) / 2),
      y: emu((slideHeightPx - rowHeight * rows) / 2),
      width: emu(width),
      height: emu(rowHeight * rows),
      columnWidths: Array.from({ length: cols }, () => emu(width / cols)),
      rows: Array.from({ length: rows }, () => ({
        height: emu(rowHeight),
        cells: Array.from({ length: cols }, () => ({
          text: "",
          borders: { top: edge, bottom: edge, left: edge, right: edge },
        })),
      })),
    },
  };
}
