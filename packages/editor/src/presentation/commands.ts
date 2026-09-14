// Ribbon command edits over the source deck JSON: slide/object insertion and
// text formatting. Pure functions — the element owns undo/redo bookkeeping
// and re-projection; these only mutate the model. Text bodies normalize their
// input sugar in place (body.text, string paragraphs, string runs) so the
// commands always walk real run objects, the same expansion stringify does.

import { EMU_PER_PX } from "@docen/layout";
import type { PresentationOptions, SlideChild } from "@docen/pptx";
// Text/paragraph types come from @office-open/core/drawing — @docen/pptx's
// re-export surface doesn't carry them yet.
import type {
  ParagraphDescriptorOptions,
  StrikeStyle,
  TextAlignment,
  TextBodyOptions,
  TextRunOptions,
  UnderlineStyle,
} from "@office-open/core/drawing";

/** The body's paragraphs as real objects, consuming sugar as needed. */
function paragraphsOf(body: TextBodyOptions): ParagraphDescriptorOptions[] {
  const paragraphs = (body.paragraphs ??=
    body.text !== undefined ? [{ children: [{ text: body.text }] }] : []);
  delete body.text;
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i]!;
    if (typeof p === "string") {
      paragraphs[i] = { children: [{ text: p }] };
      continue;
    }
    // Paragraph-level text sugar (input-only) expands to a one-run paragraph.
    if (p.children === undefined && typeof p.text === "string") {
      p.children = [{ text: p.text }];
      delete p.text;
    }
  }
  return paragraphs as ParagraphDescriptorOptions[];
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

function runsIn(body: TextBodyOptions): TextRunOptions[] {
  const runs: TextRunOptions[] = [];
  for (const paragraph of paragraphsOf(body)) runs.push(...runsOf(paragraph));
  return runs;
}

/** Word's toggle: a flag every run already carries is cleared, otherwise it
 *  is applied. No runs — no change. */
export function toggleRunFlag(body: TextBodyOptions, flag: "bold" | "italic"): void {
  const runs = runsIn(body);
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
  body: TextBodyOptions,
  key: "underline" | "strike",
  on: UnderlineStyle | StrikeStyle,
): void {
  const runs = runsIn(body);
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

export function setRunFont(body: TextBodyOptions, font: string): void {
  for (const run of runsIn(body)) run.font = font;
}

/** Run font size in points (the JSON's own unit). */
export function setRunSize(body: TextBodyOptions, size: number): void {
  for (const run of runsIn(body)) run.size = size;
}

/** Set every paragraph's horizontal alignment (PowerPoint's align buttons
 *  assign, they don't toggle). */
export function setParagraphAlignment(body: TextBodyOptions, alignment: TextAlignment): void {
  for (const paragraph of paragraphsOf(body)) {
    paragraph.properties = { ...paragraph.properties, alignment };
  }
}

/** Insert a blank slide after `at` (the PowerPoint new-slide position). */
export function insertSlideAt(deck: PresentationOptions, at: number): void {
  (deck.slides ??= []).splice(at + 1, 0, { children: [] });
}

/** The shape's text body as plain lines (paragraphs joined by \n), or null
 *  when the child carries no text body. */
export function shapeTextOf(child: SlideChild): string | null {
  if (!("shape" in child) || !child.shape.textBody) return null;
  return paragraphsOf(child.shape.textBody)
    .map((paragraph) =>
      runsOf(paragraph)
        .map((run) => run.text)
        .join(""),
    )
    .join("\n");
}

/** Write plain lines back into the shape's text body — one paragraph per
 *  line, each keeping the original paragraph's alignment and the style
 *  attributes of its first run. Returns false when the child has no text
 *  body (the caller's edit records nothing). */
export function writeShapeText(child: SlideChild, text: string): boolean {
  if (!("shape" in child) || !child.shape.textBody) return false;
  const body = child.shape.textBody;
  const previous = paragraphsOf(body);
  const lines = text.split("\n");
  const next = lines.map((line, i) => {
    const old = previous[i];
    const style = runsOf(previous[i] ?? {})[0];
    const run: TextRunOptions = { text: line };
    if (style) {
      const { text: _dropped, ...attrs } = style;
      Object.assign(run, attrs);
    }
    return { properties: old?.properties ? { ...old.properties } : undefined, children: [run] };
  });
  body.paragraphs = next;
  delete body.text;
  return true;
}

/** The body's first run font size in points (PowerPoint's 18pt default) —
 *  the in-place text editor's only typographic nod. */
export function firstRunSizeOf(child: SlideChild): number {
  if (!("shape" in child) || !child.shape.textBody) return 18;
  for (const run of runsIn(child.shape.textBody)) {
    if (typeof run.size === "number" && run.size > 0) return run.size;
  }
  return 18;
}

const emu = (px: number): number => Math.round(px * EMU_PER_PX);

/** A centered text-box shape for a slide of the given size — white fill and
 *  a hairline border keep an empty box visible on the canvas. */
export function makeTextBox(slideWidthPx: number, slideHeightPx: number): SlideChild {
  const width = slideWidthPx * 0.4;
  const height = slideHeightPx * 0.15;
  return {
    shape: {
      x: emu((slideWidthPx - width) / 2),
      y: emu((slideHeightPx - height) / 2),
      width: emu(width),
      height: emu(height),
      properties: { geometry: "rect", fill: "FFFFFF", outline: { width: "1pt", color: "808080" } },
      textBody: { text: "" },
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
