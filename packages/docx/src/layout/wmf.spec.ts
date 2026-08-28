import type { LayoutDrawingMember } from "@docen/layout";
import { describe, expect, it } from "vitest";

import { wmfMembers } from "./wmf";
import { dib, decodedBmp, gbk, logfont, words, wmfWithRecords } from "./wmf-test-util";

// ── record builders (fn + raw params), in the order a GDI emitter pushes ──

const brushRec = (colorRef: number, style = 0): { fn: number; params: Uint8Array } => {
  const p = new Uint8Array(10);
  const v = new DataView(p.buffer);
  v.setUint16(0, style, true);
  v.setUint32(2, colorRef, true);
  return { fn: 0x02fc, params: p };
};

const penRec = (colorRef: number, width = 0, style = 1): { fn: number; params: Uint8Array } => {
  const p = new Uint8Array(10);
  const v = new DataView(p.buffer);
  v.setUint16(0, style, true);
  v.setInt32(2, width, true);
  v.setUint32(6, colorRef, true);
  return { fn: 0x02fa, params: p };
};

const select = (i: number): { fn: number; params: Uint8Array } => ({
  fn: 0x012d,
  params: words(i),
});

const del = (i: number): { fn: number; params: Uint8Array } => ({
  fn: 0x01f0,
  params: words(i),
});

const poly = (pts: [number, number][], fn = 0x0324): { fn: number; params: Uint8Array } => ({
  fn,
  params: words(pts.length, ...pts.flat()),
});

/** ExtTextOut record: Y, X, byte count, options, [rect], raw string bytes,
 *  [per-byte dx words]. */
const extTextOut = (
  y: number,
  x: number,
  text: Uint8Array,
  o: { opts?: number; rect?: number[]; dx?: number[] } = {},
): { fn: number; params: Uint8Array } => {
  const before = words(y, x, text.length, o.opts ?? 0, ...(o.rect ?? []));
  const after = o.dx ? words(...o.dx) : new Uint8Array(0);
  const params = new Uint8Array(before.length + text.length + after.length);
  params.set(before, 0);
  params.set(text, before.length);
  params.set(after, before.length + text.length);
  return { fn: 0x0a32, params };
};

const RED = 0x0000ff; // COLORREF 0x00BBGGRR
const GREEN = 0x00ff00;

/** One little-endian u32 param. */
const u32 = (n: number): Uint8Array => {
  const b = new Uint8Array(4);
  new DataView(b.buffer).setUint32(0, n, true);
  return b;
};

function firstPath(members: LayoutDrawingMember[]) {
  const m = members.find((mem) => mem.kind === "path");
  expect(m).toBeDefined();
  return m as Extract<LayoutDrawingMember, { kind: "path" }>;
}

/** The first text run of a replayed text box. */
function firstRun(m: Extract<LayoutDrawingMember, { kind: "textBox" }>) {
  const block = m.blocks[0];
  if (!block || block.kind !== "paragraph") throw new Error("expected a paragraph block");
  const run = block.inline[0];
  if (!run || run.kind !== "text") throw new Error("expected a text run");
  return run;
}

describe("wmfMembers", () => {
  it("replays a polygon as a path with brush fill and pen stroke, box-scaled", () => {
    const members = wmfMembers(
      wmfWithRecords([
        brushRec(RED),
        penRec(GREEN),
        select(0),
        select(1),
        poly([
          [10, 10],
          [50, 10],
          [30, 60],
        ]),
      ]),
      100,
      50,
    )!;
    const m = firstPath(members);
    // window 200×100 → box 100×50 halves every coordinate.
    expect(m.x).toBe(5);
    expect(m.y).toBe(5);
    expect(m.width).toBe(20);
    expect(m.height).toBe(25);
    expect(m.d).toBe("M0,0L20,0L10,25Z");
    expect(m.fill).toBe("ff0000");
    expect(m.line?.color).toBe("00ff00");
    expect(m.line?.px).toBe(1); // width-0 pen → hairline
  });

  it("carries SetPolyFillMode into the path fillRule", () => {
    const stream = (mode?: number) =>
      wmfWithRecords([
        brushRec(RED),
        select(0),
        ...(mode != null ? [{ fn: 0x0106, params: words(mode) }] : []),
        poly([
          [10, 10],
          [60, 10],
          [35, 60],
        ]),
      ]);
    expect(firstPath(wmfMembers(stream(), 200, 100)!).fillRule).toBe("evenodd");
    expect(firstPath(wmfMembers(stream(2), 200, 100)!).fillRule).toBe("nonzero");
  });

  it("leaves polylines unfilled and merges a PolyPolygon into one path", () => {
    const line = wmfMembers(
      wmfWithRecords([
        penRec(GREEN),
        select(0),
        poly(
          [
            [10, 10],
            [40, 20],
            [40, 50],
          ],
          0x0325,
        ),
      ]),
      200,
      100,
    )!;
    const m = firstPath(line);
    expect(m.fill).toBeUndefined();
    expect(m.d.startsWith("M0,0L30,10L30,40")).toBe(true);
    expect(m.d.endsWith("Z")).toBe(false);

    const tri = (x: number): number[] => [x, 10, x + 20, 10, x + 10, 30];
    const pp = wmfMembers(
      wmfWithRecords([
        brushRec(RED),
        select(0),
        { fn: 0x0538, params: words(2, 3, 3, ...tri(10), ...tri(60)) },
      ]),
      200,
      100,
    )!;
    expect(pp.filter((mem) => mem.kind === "path")).toHaveLength(1);
    expect(firstPath(pp).d).toBe("M0,0L20,0L10,20ZM50,0L70,0L60,20Z");
  });

  it("maps Rectangle/RoundRect/Ellipse to shape presets and normalizes swapped pairs", () => {
    const members = wmfMembers(
      wmfWithRecords([
        brushRec(RED),
        select(0),
        // params are Bottom, Right, Top, Left; RoundRect leads Height/Width
        { fn: 0x041b, params: words(50, 80, 10, 20) },
        { fn: 0x061c, params: words(6, 6, 50, 80, 10, 20) },
        { fn: 0x0418, params: words(10, 80, 50, 20) }, // top/bottom swapped
      ]),
      200,
      100,
    )!;
    expect(members).toHaveLength(3);
    const [rect, round, ellipse] = members as Extract<LayoutDrawingMember, { kind: "shape" }>[];
    expect(rect).toMatchObject({
      preset: "rect",
      x: 20,
      y: 10,
      width: 60,
      height: 40,
      fill: "ff0000",
    });
    expect(round).toMatchObject({ preset: "roundRect", x: 20, y: 10, width: 60, height: 40 });
    expect(ellipse).toMatchObject({ preset: "ellipse", x: 20, y: 10, width: 60, height: 40 });
  });

  it("replays ExtTextOut as a text box: GBK font family, dyed color, dx advance", () => {
    const gbkText = gbk("示例文本");
    // lead byte advances 16, trail byte 0 — one dx word per byte
    const dx = Array.from(gbkText, (_, i) => (i % 2 === 0 ? 16 : 0));
    const members = wmfMembers(
      wmfWithRecords([
        { fn: 0x02fb, params: logfont({ height: -20, weight: 700, face: "微软雅黑" }) },
        select(0),
        { fn: 0x0209, params: u32(RED) },
        { fn: 0x012e, params: words(24) }, // TA_BASELINE — the reference y is the baseline
        extTextOut(50, 100, gbkText, { dx }),
      ]),
      200,
      100,
    )!;
    const m = members[0] as Extract<LayoutDrawingMember, { kind: "textBox" }>;
    const { style } = firstRun(m);
    expect(firstRun(m).text).toBe("示例文本");
    expect(m.x).toBe(100);
    expect(m.y).toBeCloseTo(50 - 0.8 * 20); // TA_BASELINE ascent hoist
    expect(m.width).toBe(4 * 16 + 2); // dx sum + wrap guard
    expect(style.family).toBe("微软雅黑");
    expect(style.sizePx).toBe(20);
    expect(style.color).toBe("ff0000");
    expect(style.bold).toBe(true);
  });

  it("treats the device-default TA_TOP reference y as the cell top (no hoist)", () => {
    const members = wmfMembers(
      wmfWithRecords([
        { fn: 0x02fb, params: logfont({ height: -20, face: "微软雅黑" }) },
        select(0),
        // No SetTextAlign — GDI's default TA_TOP: the y names the box top.
        extTextOut(50, 100, gbk("示例文本"), {}),
      ]),
      200,
      100,
    )!;
    const m = members[0] as Extract<LayoutDrawingMember, { kind: "textBox" }>;
    expect(m.y).toBe(50);
  });

  it("skips the ETO_OPAQUE rect without shifting the string", () => {
    const members = wmfMembers(
      wmfWithRecords([
        { fn: 0x02fb, params: logfont({ height: -12, face: "微软雅黑" }) },
        select(0),
        extTextOut(60, 10, gbk("示例文本"), { opts: 0x0002, rect: [0, 0, 90, 40] }),
      ]),
      200,
      100,
    )!;
    const m = members[0] as Extract<LayoutDrawingMember, { kind: "textBox" }>;
    expect(firstRun(m).text).toBe("示例文本");
  });

  it("frames SRCCOPY stretch-blits as picture members and skips mask layers", () => {
    const rop = (r: number): number[] => [r & 0xffff, r >>> 16];
    const dibBytes = dib(9, 9, 24);
    const blt = (r: number, dest: [number, number, number, number]) => {
      const head = words(...rop(r), 9, 9, 0, 0, dest[3], dest[2], dest[1], dest[0]);
      const params = new Uint8Array(head.length + dibBytes.length);
      params.set(head, 0);
      params.set(dibBytes, head.length);
      return { fn: 0x0b41, params };
    };
    const members = wmfMembers(
      wmfWithRecords([blt(0x00cc0020, [20, 40, 60, 30]), blt(0xee0086, [0, 0, 60, 30])]),
      100,
      50,
    )!;
    expect(members).toHaveLength(1);
    const m = members[0] as Extract<LayoutDrawingMember, { kind: "picture" }>;
    expect(m.x).toBe(10);
    expect(m.y).toBe(20);
    expect(m.width).toBe(30);
    expect(m.height).toBe(15);
    expect(m.src?.startsWith("data:image/bmp;base64,")).toBe(true);
    expect(decodedBmp(m.src).length).toBe(14 + dibBytes.length);
    // SRCPAINT mask layer alone yields nothing drawable.
    expect(wmfMembers(wmfWithRecords([blt(0xee0086, [0, 0, 60, 30])]), 100, 50)).toBeUndefined();
  });

  it("restores the DC state across SaveDC/RestoreDC", () => {
    const members = wmfMembers(
      wmfWithRecords([
        brushRec(RED),
        select(0),
        { fn: 0x001e, params: new Uint8Array(0) },
        brushRec(GREEN),
        select(1),
        { fn: 0x0127, params: words(-1) },
        poly([
          [10, 10],
          [60, 10],
          [35, 60],
        ]),
      ]),
      200,
      100,
    )!;
    expect(firstPath(members).fill).toBe("ff0000");
  });

  it("releases slots on DeleteObject so a recreated object is re-selected", () => {
    const members = wmfMembers(
      wmfWithRecords([
        brushRec(RED),
        select(0),
        del(0),
        brushRec(GREEN), // reuses freed slot 0
        select(0),
        poly([
          [10, 10],
          [60, 10],
          [35, 60],
        ]),
      ]),
      200,
      100,
    )!;
    expect(firstPath(members).fill).toBe("00ff00");
  });

  it("normalizes against an offset placeable bbox", () => {
    const tri: [number, number][] = [
      [150, 75],
      [200, 75],
      [175, 125],
    ];
    // A brushless, penless polygon draws nothing — no member at all.
    expect(wmfMembers(wmfWithRecords([poly(tri)], 200, 100, 100, 50), 200, 100)).toBeUndefined();
    const withBrush = wmfMembers(
      wmfWithRecords([brushRec(RED), select(0), poly(tri)], 200, 100, 100, 50),
      200,
      100,
    )!;
    expect(firstPath(withBrush).x).toBe(50);
    expect(firstPath(withBrush).y).toBe(25);
  });

  it("returns undefined for non-placeable bytes, empty output, and truncated streams", () => {
    const notPlaceable = wmfWithRecords([brushRec(RED), select(0)]);
    notPlaceable[0] = 0;
    expect(wmfMembers(notPlaceable, 100, 50)).toBeUndefined();
    // unknown fn + degenerate geometry: skipped, nothing drawable
    expect(
      wmfMembers(wmfWithRecords([{ fn: 0x0626, params: words(1, 2, 3) }]), 100, 50),
    ).toBeUndefined();
    // record header claims more bytes than the stream holds — walk stops safe
    const truncated = wmfWithRecords([brushRec(RED), select(0)]);
    const view = new DataView(truncated.buffer);
    view.setUint32(40 + 6, 9999, true);
    expect(() => wmfMembers(truncated.subarray(0, 60), 100, 50)).not.toThrow();
  });
});
