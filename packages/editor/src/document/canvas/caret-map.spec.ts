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
  measureText: (text: string) => ({ width: text.length * 10 }),
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
    const { doc } = buildDoc(["abcd"]);
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
    const { doc } = buildDoc(["abcd"]);
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

describe("CaretMap trimmed-space boundaries", () => {
  // "ab cd": pretext trims the inter-word space into the gap between the
  // items (ab@0..20, cd@30..50) — the character itself stays in the PM text,
  // so the collapsed-char space must count it back in.
  const gappedPara = (text: string, cdX: number): Record<string, unknown> => ({
    kind: "paragraph",
    heightPx: 20,
    beforePx: 0,
    afterPx: 0,
    inline: [{ kind: "text", text, style: { sizePx: 16, family: "Test" } }],
    lines: [
      {
        yPx: 0,
        heightPx: 20,
        naturalPx: 16,
        items: [
          { kind: "text", text: "ab", xPx: 0, widthPx: 20, inlineIndex: 0 },
          { kind: "text", text: "cd", xPx: cdX, widthPx: 20, inlineIndex: 0 },
        ],
        maxWidthPx: 100,
      },
    ],
  });

  it("counts trimmed gap characters so positions match the PM text", () => {
    // Before: the collapsed space held 4 chars, so the line-end click mapped
    // to innerPos+4 — one char short of the paragraph's real end (5 chars).
    const { doc } = buildDoc(["ab cd"]);
    const map = new CaretMap(pageOf([gappedPara("ab cd", 30)]) as never, doc, () => ({
      contentLeftPx: 0,
      contentTopPx: 0,
    }));
    // Past the last glyph → the paragraph's true end (innerPos 1 + 5 chars).
    expect(map.posAtPoint(0, 48, 5)).toBe(6);
    // Inside the gap → the space character's own boundaries.
    expect(map.posAtPoint(0, 22, 5)).toBe(3);
    expect(map.caretRect(3)?.xPx).toBe(20); // the space's left edge (ab's end)
    expect(map.caretRect(4)?.xPx).toBe(30); // its right edge = cd's start
    expect(map.caretRect(6)?.xPx).toBe(50); // the line's advance sum
  });

  it("splits a multi-space gap's boundaries evenly across the gap", () => {
    // "ab  cd": two trimmed spaces share the 20px gap — the caret boundaries
    // sit at 20 / 30 inside it, the same split the space dots center in.
    const { doc } = buildDoc(["ab  cd"]);
    const map = new CaretMap(pageOf([gappedPara("ab  cd", 40)]) as never, doc, () => ({
      contentLeftPx: 0,
      contentTopPx: 0,
    }));
    expect(map.caretRect(3)?.xPx).toBe(20); // first space's left edge
    expect(map.caretRect(4)?.xPx).toBe(30); // between the two spaces
    expect(map.caretRect(5)?.xPx).toBe(40); // cd's start
    // The paragraph's real end: innerPos 1 + 6 chars.
    expect(map.posAtPoint(0, 58, 5)).toBe(7);
  });
});

describe("CaretMap selection rectangles", () => {
  it("stretches fully crossed lines to the wrap edge; the last line stops at text", () => {
    const { doc } = buildDoc(["ab", "cde"]);
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
    const { doc } = buildDoc(["abc"]);
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
    const { doc } = buildDoc(["ab", "cde"]);
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
    const { doc } = buildDoc(["ab", ""]);
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

  it("highlights a render-only line across its full width (TOC entry)", () => {
    // A TOC entry paints from its cached options while the PM paragraph stays
    // empty: the zip pairs them as-is, so no PM position maps into the line.
    // The selection crossing the paragraph must still highlight the line the
    // reader sees — not drop it (the old char-intersection test did).
    const { doc } = buildDoc(["ab", ""]);
    const map = new CaretMap(
      pageOf([
        fakePara([{ text: "ab", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
        fakePara([{ text: "一、示例条目", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
      ]) as never,
      doc,
      () => ({ contentLeftPx: 0, contentTopPx: 0 }),
    );
    const rects = map.selectionRects(0, doc.content.size);
    expect(rects).toHaveLength(2);
    expect(rects[1]).toMatchObject({ xPx: 0, yPx: 50, widthPx: 100, heightPx: 20 });
  });
});

describe("CaretMap tolerant zip", () => {
  it("stays valid and keeps later positions correct across unlaid PM textblocks", () => {
    // A floating table's cells paint in the scene without flow items: the PM
    // doc carries their paragraphs but the layout never lays them. The zip
    // must skip them unmapped instead of invalidating the whole map (every
    // click dead) — paragraphs after the gap still pair with their own text.
    const { doc } = buildDoc(["aaa", "bbb", "ccc", "ddd"]);
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
    const { doc } = buildDoc(["aaa", "bbb"]);
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
    const { doc } = buildDoc(["title", "", ""]);
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

  it("resyncs after a render-only run longer than the blank-paragraph supply", () => {
    // A multi-entry TOC lays N paragraphs over one empty field paragraph: the
    // first entry pair-as-is's onto the blank, the remaining entries meet the
    // REAL body heading as `there` and must skip ahead (laid-side anchor) to
    // the heading's laid block — not pair onto the heading and shift every
    // later paragraph by one.
    const { editor: _editor, doc } = buildDoc(["", "heading", "body"]);
    const map = new CaretMap(
      pageOf([
        fakePara([{ text: "heading1", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
        fakePara([{ text: "entry2", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
        fakePara([{ text: "heading", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
        fakePara([{ text: "body", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
      ]) as never,
      doc,
      () => ({ contentLeftPx: 0, contentTopPx: 0 }),
    );
    expect(map.valid).toBe(true);
    // The heading's laid line (y=100) maps the heading's own PM position
    // (doc: blank 0-2, heading 2-11 → its inner start is 3).
    expect(map.posAtPoint(0, 5, 105)).toBe(3);
    expect(map.caretRect(3)?.yPx).toBe(100);
    // The paragraph after it stays aligned too (inner start 12).
    expect(map.posAtPoint(0, 5, 155)).toBe(12);
    expect(map.caretRect(12)?.yPx).toBe(150);
  });

  it("navigates posVertical across paragraph boundaries", () => {
    // Two paragraphs: "Hello" (pos 1..6) and "World" (pos 8..13).
    const { doc } = buildDoc(["Hello", "World"]);
    const map = new CaretMap(
      pageOf([
        fakePara([{ text: "Hello", xPx: 0, yPx: 0, maxWidthPx: 100 }]),
        fakePara([{ text: "World", xPx: 0, yPx: 50, maxWidthPx: 100 }]),
      ]) as never,
      doc,
      () => ({ contentLeftPx: 0, contentTopPx: 0 }),
    );
    expect(map.valid).toBe(true);
    // From inside first paragraph, stepping down lands in the second paragraph at the same column.
    // 'H' is at pos 1, stepping down lands at 'W' at pos 8.
    expect(map.posVertical(1, 1)).toBe(8);
    // 'l' (second l) is at pos 4, stepping down lands at 'l' at pos 11.
    expect(map.posVertical(4, 1)).toBe(11);
    // From second paragraph, stepping up lands in the first paragraph.
    expect(map.posVertical(8, -1)).toBe(1);
    expect(map.posVertical(11, -1)).toBe(4);
    // Top-edge and bottom-edge return null.
    expect(map.posVertical(1, -1)).toBeNull();
    expect(map.posVertical(8, 1)).toBeNull();
  });

  it("resolves clicks in empty space far below text to the nearest line", () => {
    const { doc } = buildDoc(["Single line at top"]);
    const map = new CaretMap(
      pageOf([
        fakePara([{ text: "Single line at top", xPx: 0, yPx: 0, maxWidthPx: 200 }]),
      ]) as never,
      doc,
      () => ({ contentLeftPx: 0, contentTopPx: 0 }),
    );
    expect(map.valid).toBe(true);
    // Click at y=500 (far below the line at y=0, dist = 480 > 40).
    // Should resolve to the line rather than returning null.
    const pos = map.posAtPoint(0, 50, 500);
    expect(pos).not.toBeNull();
    expect(pos).toBeGreaterThanOrEqual(1);
  });
});
