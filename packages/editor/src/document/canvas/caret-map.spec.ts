// @vitest-environment node
import { Document, Paragraph } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import type { FlowPage } from "@docen/layout";
import { describe, expect, it, vi } from "vitest";

// CaretMap grabs a 2d canvas at module load for per-grapheme measurements;
// node has neither — stub the document BEFORE the dynamic import with a
// deterministic 10px-per-grapheme font so boundary lattices are exact.
const fakeCtx = {
  set font(_v: string) {},
  measureText: (text: string) => ({ width: [...text].length * 10 }),
} as unknown as CanvasRenderingContext2D;
vi.stubGlobal("document", {
  createElement: (tag: string) => (tag === "canvas" ? { getContext: () => fakeCtx } : {}),
} as unknown as Document);

const { CaretMap } = await import("./caret-map");

// Tiptap's schema needs the plain text node (same trick as the TOC spec).
const Text = TextNode.create({ name: "text", group: "inline" });

const buildDoc = (texts: string[]): { editor: EditorType; doc: EditorType["state"]["doc"] } => {
  const editor = new Editor({
    element: null,
    extensions: [Document, Paragraph, Text],
    content: {
      type: "doc",
      content: texts.map((text) => ({
        type: "paragraph",
        content: text ? [{ type: "text", text }] : undefined,
      })),
    },
  });
  return { editor, doc: editor.state.doc };
};

interface FakeLine {
  text: string;
  xPx: number;
  yPx: number;
  maxWidthPx?: number;
  justifyGapPx?: number;
}

/** A laid paragraph whose items split the text across the given lines
 *  (10px/grapheme widths, matching the measurement stub). */
const fakePara = (lines: FakeLine[]): Record<string, unknown> => {
  const text = lines.map((l) => l.text).join("");
  return {
    kind: "paragraph",
    heightPx: lines.length * 20,
    beforePx: 0,
    afterPx: 0,
    inline: [{ kind: "text", text, style: { sizePx: 16, family: "Test" } }],
    lines: lines.map((l) => ({
      yPx: l.yPx,
      heightPx: 20,
      naturalPx: 16,
      items: l.text
        ? [{ kind: "text", text: l.text, xPx: l.xPx, widthPx: l.text.length * 10, inlineIndex: 0 }]
        : [],
      maxWidthPx: l.maxWidthPx,
      justifyGapPx: l.justifyGapPx,
      hangPx: undefined,
    })),
  };
};

const pageOf = (blocks: Record<string, unknown>[]): FlowPage[] =>
  [{ items: blocks.map((block, i) => ({ yPx: i * 50, block })) }] as unknown as FlowPage[];

describe("CaretMap click boundaries", () => {
  it("pairs each boundary x with its own doc position (click a full character off)", () => {
    // "abcd" at 10px/grapheme: boundaries at x=0/10/20/30/40. A click at x=15
    // is equidistant from boundaries 1 and 2 and must resolve to 1 — the old
    // pairing handed boundary 1's x to position 2, landing every click one
    // character right.
    const { editor, doc } = buildDoc(["abcd"]);
    const map = new CaretMap(
      pageOf([fakePara([{ text: "abcd", xPx: 0, yPx: 0, maxWidthPx: 100 }])]) as never,
      doc,
      () => ({ contentLeftPx: 0, contentTopPx: 0 }),
    );
    expect(map.valid).toBe(true);
    // Paragraph innerPos = 1 (doc > paragraph > text).
    expect(map.posAtPoint(0, 15, 5)).toBe(2);
    expect(map.posAtPoint(0, 5, 5)).toBe(1);
    expect(map.posAtPoint(0, 25, 5)).toBe(3);
    // Past the last glyph → the line-end boundary (Word's line-end click).
    expect(map.posAtPoint(0, 35, 5)).toBe(4);
    expect(map.posAtPoint(0, 38, 5)).toBe(5);
  });

  it("renders the caret at the clicked boundary, not one past it", () => {
    const { editor, doc } = buildDoc(["abcd"]);
    const map = new CaretMap(
      pageOf([fakePara([{ text: "abcd", xPx: 0, yPx: 0, maxWidthPx: 100 }])]) as never,
      doc,
      () => ({ contentLeftPx: 0, contentTopPx: 0 }),
    );
    // Click at x=15 → position 2 → the caret draws at boundary 1's x (10),
    // within half a glyph of the click.
    expect(map.posAtPoint(0, 15, 5)).toBe(2);
    expect(map.caretRect(2)?.xPx).toBe(10);
    // The line-end position (5 = innerPos + 4 glyphs) draws at the line's
    // right edge (the advance sum), not the last glyph's left edge.
    expect(map.caretRect(5)?.xPx).toBe(40);
  });
});

describe("CaretMap selection rectangles", () => {
  it("stretches fully crossed lines to the wrap edge; the last line stops at text", () => {
    const { editor, doc } = buildDoc(["ab", "cde"]);
    const map = new CaretMap(
      pageOf([
        fakePara([{ text: "ab", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
        fakePara([{ text: "cde", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
      ]) as never,
      doc,
      () => ({ contentLeftPx: 0, contentTopPx: 0 }),
    );
    const rects = map.selectionRects(0, doc.content.size);
    // Paragraph 1's line: selected past its end (the selection continues into
    // paragraph 2) → the highlight runs to the wrap edge (100), not "ab" (20).
    expect(rects[0]).toMatchObject({ xPx: 0, widthPx: 100 });
    // The document's last line stops at the last glyph (30) — there is no
    // content after it to stretch to.
    expect(rects[1]).toMatchObject({ xPx: 0, widthPx: 30 });
  });

  it("keeps multi-line highlights contiguous with line-box heights", () => {
    const { editor, doc } = buildDoc(["abc"]);
    // One paragraph wrapped into two lines ("ab" / "c", 20px tall each).
    const map = new CaretMap(
      pageOf([
        fakePara([
          { text: "ab", xPx: 0, yPx: 0, maxWidthPx: 100 },
          { text: "c", xPx: 0, yPx: 20, maxWidthPx: 100 },
        ]),
      ]) as never,
      doc,
      () => ({ contentLeftPx: 0, contentTopPx: 0 }),
    );
    const rects = map.selectionRects(0, doc.content.size);
    expect(rects).toHaveLength(2);
    expect(rects[0]).toMatchObject({ yPx: 0, heightPx: 20 });
    expect(rects[1]).toMatchObject({ yPx: 20, heightPx: 20 });
  });

  it("highlights the paragraph gap once the next paragraph is selected too", () => {
    const { editor, doc } = buildDoc(["ab", "cde"]);
    // Paragraph 2's block sits 50px down (fakePara page items step 50): the
    // gap between the paragraphs belongs to a whole-document selection.
    const map = new CaretMap(
      pageOf([
        fakePara([{ text: "ab", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
        fakePara([{ text: "cde", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
      ]) as never,
      doc,
      () => ({ contentLeftPx: 0, contentTopPx: 0 }),
    );
    const rects = map.selectionRects(0, doc.content.size);
    // Paragraph 1's rect bottom reaches paragraph 2's line top (page-local 50).
    expect(rects[0]?.heightPx).toBe(50);
    // Selecting paragraph 1 alone stops at its own line box.
    const solo = map.selectionRects(0, 3);
    expect(solo[0]?.heightPx).toBe(20);
  });

  it("shows an empty paragraph as a caret-width block", () => {
    const { editor, doc } = buildDoc(["ab", ""]);
    const map = new CaretMap(
      pageOf([
        fakePara([{ text: "ab", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
        fakePara([{ text: "", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
      ]) as never,
      doc,
      () => ({ contentLeftPx: 0, contentTopPx: 0 }),
    );
    const rects = map.selectionRects(0, doc.content.size);
    expect(rects).toHaveLength(2);
    expect(rects[1]).toMatchObject({ xPx: 0, widthPx: 8, heightPx: 20 });
  });
});

describe("CaretMap tolerant zip", () => {
  it("stays valid and keeps later positions correct across unlaid PM textblocks", () => {
    // A floating table's cells paint in the scene without flow items: the PM
    // doc carries their paragraphs but the layout never lays them. The zip
    // must skip them unmapped instead of invalidating the whole map (every
    // click dead) — paragraphs after the gap still pair with their own text.
    const { editor, doc } = buildDoc(["aaa", "bbb", "ccc", "ddd"]);
    const map = new CaretMap(
      pageOf([
        fakePara([{ text: "aaa", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
        fakePara([{ text: "ddd", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
      ]) as never,
      doc,
      () => ({ contentLeftPx: 0, contentTopPx: 0 }),
    );
    expect(map.valid).toBe(true);
    // "aaa" pairs with the first paragraph, "ddd" with the last — not
    // shifted onto the unlaid cells (inner positions 1 and 16).
    expect(map.posAtPoint(0, 5, 5)).toBe(1); // aaa's start
    expect(map.posAtPoint(0, 5, 55)).toBe(16); // ddd's start
    expect(map.caretRect(1)?.yPx).toBe(0);
    expect(map.caretRect(16)?.yPx).toBe(50);
  });

  it("stays valid when the layout emits render-only paragraphs the PM lacks", () => {
    // A repeated table header row renders on every continuation page but is
    // one PM row: the duplicated header block pairs with nothing. The zip
    // skips it once the NEXT laid block's text matches the current textblock.
    const { editor, doc } = buildDoc(["aaa", "bbb"]);
    const map = new CaretMap(
      pageOf([
        fakePara([{ text: "aaa", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
        fakePara([{ text: "bbb", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
        // Render-only repeat of the header on the next page.
        fakePara([{ text: "aaa", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
      ]) as never,
      doc,
      () => ({ contentLeftPx: 0, contentTopPx: 0 }),
    );
    expect(map.valid).toBe(true);
    expect(map.caretRect(1)?.yPx).toBe(0);
    expect(map.caretRect(6)?.yPx).toBe(50);
  });

  it("pairs same-position text drift positionally (a TOC laid from cached text)", () => {
    // A TOC's PM field-content paragraphs hold no text; the laid entries render
    // their cached option text at the same positions. Text differs at every
    // pair, no resync fires, and the positional pairing keeps the map valid.
    const { editor, doc } = buildDoc(["title", "", ""]);
    const map = new CaretMap(
      pageOf([
        fakePara([{ text: "title", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
        fakePara([{ text: "cached entry one", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
        fakePara([{ text: "cached entry two", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
      ]) as never,
      doc,
      () => ({ contentLeftPx: 0, contentTopPx: 0 }),
    );
    expect(map.valid).toBe(true);
    expect(map.caretRect(1)?.yPx).toBe(0);
    expect(map.caretRect(8)?.yPx).toBe(50);
    expect(map.caretRect(10)?.yPx).toBe(100);
  });
});
