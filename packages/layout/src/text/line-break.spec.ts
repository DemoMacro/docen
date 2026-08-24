import { describe, expect, it } from "vitest";

import { fakeFontMetrics, installFakeCanvas } from "../../test/fake-canvas";
import { type LayoutInline, type LayoutTextStyle } from "../layout-doc";
import type { LaidOutLineItem } from "../layout-result";
import { packLines } from "./line-break";
import { TextMeasurer } from "./measure";

installFakeCanvas();
const measurer = new TextMeasurer(fakeFontMetrics);

// Synthetic world: latin grapheme = em/2, CJK = em, space = em/4 (16px em →
// 8 / 16 / 4). All expected widths below are hand-derived from that.
const latin: LayoutTextStyle = { family: "serif", sizePx: 16 };
const cjk: LayoutTextStyle = { family: { latin: "serif", eastAsia: "SimSun" }, sizePx: 16 };

const text = (t: string, style: LayoutTextStyle = latin): LayoutInline => ({
  kind: "text",
  text: t,
  style,
});

function pack(
  inline: LayoutInline[],
  width: number,
  opts: Partial<Parameters<typeof packLines>[1]> = {},
) {
  return packLines(inline, {
    measurer,
    width,
    lineHeight: ({ naturalPx }) => naturalPx,
    ...opts,
  });
}

const textsOf = (line: { items: LaidOutLineItem[] }): string => {
  const out: string[] = [];
  for (const item of line.items) if (item.kind === "text") out.push(item.text);
  return out.join("");
};

describe("packLines", () => {
  it("wraps Latin text at word boundaries", () => {
    // "Hello world" = 10 letters + 1 space = 84px; one more letter overflows.
    // pretext keeps the line-final space (it hangs invisibly past the right
    // edge when painted) — assert through trimEnd for the break position.
    const lines = pack([text("Hello world foo")], 84.5);
    expect(lines).toHaveLength(2);
    expect(textsOf(lines[0]).trimEnd()).toBe("Hello world");
    expect(textsOf(lines[1])).toBe("foo");
  });

  it("breaks CJK between characters", () => {
    // 5 CJK chars = 80px; the 6th overflows.
    const lines = pack([text("中文测试中文后续", cjk)], 80.5);
    expect(lines.map(textsOf)).toEqual(["中文测试中", "文后续"]);
  });

  it("shrinks only the first line by the first-line indent", () => {
    // "abcd ab" = 32+4+16 = 52px; line 0 at 72−28.8 = 43.2px fits "abcd" only.
    const lines = pack([text("abcd ab")], 72, { firstLineIndentPx: 28.8 });
    expect(lines).toHaveLength(2);
    expect(lines[0].maxWidthPx).toBeCloseTo(43.2, 5);
    expect(lines[1].maxWidthPx).toBeCloseTo(72, 5);
    expect(textsOf(lines[0]).trimEnd()).toBe("abcd");
    expect(textsOf(lines[1])).toBe("ab");
  });

  it("splits an oversized word character-by-character", () => {
    // A word that cannot fit at all: pretext breaks it mid-word rather than
    // overflowing — Word overflows instead (registered divergence).
    const lines = pack([text("supercalifragilistic")], 40);
    expect(lines.map(textsOf).join("")).toBe("supercalifragilistic");
    expect(lines.length).toBeGreaterThan(1);
  });

  it("ends a line at a hard break", () => {
    const lines = pack([text("ab"), { kind: "break" }, text("cd")], 1000);
    expect(lines).toHaveLength(2);
    expect(textsOf(lines[0])).toBe("ab");
    expect(textsOf(lines[1])).toBe("cd");
  });

  it("collapses the boundary space on a soft wrap across inlines", () => {
    // "aaaa" = 32px; the trailing inline's leading space collapses at wrap.
    const lines = pack([text("aaaa"), text(" bbbb")], 32.5);
    expect(lines).toHaveLength(2);
    expect(textsOf(lines[1])).toBe("bbbb");
  });

  it("packs pictures as atoms and wraps them like characters", () => {
    const pic = (size: number): LayoutInline => ({
      kind: "picture",
      widthPx: size,
      heightPx: size,
    });
    const both = pack([pic(60), pic(60)], 130);
    expect(both).toHaveLength(1);
    expect(both[0].heightPx).toBe(60);

    const wrapped = pack([pic(60), pic(60)], 100);
    expect(wrapped).toHaveLength(2);
    expect(wrapped[0].endInlineIndex).toBe(0);
    expect(wrapped[1].endInlineIndex).toBe(1);
  });

  it("floors a picture-only line at the strut", () => {
    const lines = pack([{ kind: "picture", widthPx: 10, heightPx: 10 }], 100, { strutPx: 30 });
    expect(lines[0].heightPx).toBe(30);
  });

  it("reduces line width through an active float zone", () => {
    // Full "aaaa bbbb" = 68px; a 32px zone at the first line's top leaves 36.
    const lines = pack([text("aaaa bbbb cccc dddd")], 68.5, {
      startY: 0,
      floatZones: [{ widthPx: 32, topPx: -10, bottomPx: 5 }],
    });
    expect(lines.length).toBeGreaterThan(1);
    expect(lines[0].maxWidthPx).toBeLessThan(68.5);
  });

  it("mixes text and pictures in one flow", () => {
    const lines = pack([text("ab"), { kind: "picture", widthPx: 16, heightPx: 40 }], 40);
    expect(lines).toHaveLength(1);
    expect(lines[0].items).toHaveLength(2);
    expect(lines[0].heightPx).toBe(40); // picture sizes the line
  });

  it("advances a bare tab to the default 48px grid", () => {
    // "ab" = 16px; the tab jumps to the next 48px slot → text at x=48.
    const lines = pack([text("ab"), { kind: "tab" }, text("cd")], 200);
    expect(lines).toHaveLength(1);
    const items = lines[0].items;
    expect(items[items.length - 1].xPx).toBe(48);
  });

  it("right-aligns the run after a right tab stop", () => {
    // "ab" = 16px; a right stop at 100 with 16px following → tab lands at 84.
    const lines = pack([text("ab"), { kind: "tab" }, text("cd")], 200, {
      tabStops: [{ positionPx: 100, type: "right" }],
    });
    expect(lines).toHaveLength(1);
    const items = lines[0].items;
    expect(items[items.length - 1].xPx).toBe(84);
    expect(items[items.length - 1].xPx + items[items.length - 1].widthPx).toBe(100);
  });

  it("continues the same line across a tab group boundary", () => {
    const lines = pack([text("ab"), { kind: "tab" }, text("cd")], 200);
    expect(lines).toHaveLength(1);
    expect(lines[0].items.map((i) => (i.kind === "text" ? i.text : ""))).toEqual(["ab", "cd"]);
  });

  it("returns no lines for empty inline content", () => {
    expect(pack([], 100)).toHaveLength(0);
  });
});
