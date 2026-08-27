// Spec for the dual-mode metafile player (emf-plus.ts): WMFC chunk
// reassembly, GDI+ record replay, and member shaping. Fixtures are built at
// the byte level against layouts verified on real corpus files.

import { describe, expect, it } from "vitest";

import { embeddedEmfStream, emfPlusMembers } from "./emf-plus";
import { emfPlusWmf, epRecord } from "./wmf-test-util";

// EmfPlus record codes used by the fixtures.
const HEADER = 0x4001;
const OBJECT = 0x4008;
const FILL_PATH = 0x4014;
const DRAW_IMAGE_POINTS = 0x401b;
const SET_WORLD_TRANSFORM = 0x402a;
const EOF = 0x4002;
const VERSION = 0xdbc01002;

const epHeader = () => epRecord(HEADER, 1, new Uint8Array(24));

const epEof = () => epRecord(EOF, 0, new Uint8Array(0));

/** Path object: uncompressed float pairs, all type bytes = start/line. */
function pathObject(slot: number, pts: Array<[number, number]>): Uint8Array {
  const body = new Uint8Array(16 + pts.length * 9);
  const v = new DataView(body.buffer);
  v.setUint32(0, body.length, true); // TotalObjectSize
  v.setUint32(4, VERSION, true);
  v.setUint32(8, pts.length, true); // point count
  v.setUint32(12, 0, true); // format: uncompressed
  pts.forEach(([x, y], i) => {
    v.setFloat32(16 + i * 8, x, true);
    v.setFloat32(20 + i * 8, y, true);
    body[16 + pts.length * 8 + i] = i === 0 ? 0 : 1; // start, then lines
  });
  return epRecord(OBJECT, slot | (3 << 8), body); // type 3 = path
}

/** Solid brush object with the given opaque ARGB. */
function brushObject(slot: number, argb: number): Uint8Array {
  const body = new Uint8Array(16);
  const v = new DataView(body.buffer);
  v.setUint32(0, body.length, true); // TotalObjectSize
  v.setUint32(4, VERSION, true);
  v.setUint32(8, 0, true); // brushType solid
  v.setUint32(12, argb, true);
  return epRecord(OBJECT, slot | (1 << 8), body); // type 1 = brush
}

/** Image object carrying a PNG payload (spliced as a data URL on replay).
 *  The encoding sits right behind the shared header and an image-type word,
 *  matching real exporter output. */
function imageObject(slot: number, bytes: Uint8Array): Uint8Array {
  const body = new Uint8Array(12 + bytes.length);
  const v = new DataView(body.buffer);
  v.setUint32(0, body.length, true);
  v.setUint32(4, VERSION, true);
  v.setUint32(8, 1, true); // image type: bitmap
  body.set(bytes, 12);
  return epRecord(OBJECT, slot | (5 << 8), body); // type 5 = image
}

/** SetWorldTransform whose body carries the byte-length word the real
 *  exporter writes before the six-float matrix. */
function setWorld(scale: number, dx: number, dy: number): Uint8Array {
  const body = new Uint8Array(28);
  const v = new DataView(body.buffer);
  v.setUint32(0, 24, true);
  v.setFloat32(4, scale, true);
  v.setFloat32(8, 0, true);
  v.setFloat32(12, 0, true);
  v.setFloat32(16, scale, true);
  v.setFloat32(20, dx, true);
  v.setFloat32(24, dy, true);
  return epRecord(SET_WORLD_TRANSFORM, 0, body);
}

describe("embeddedEmfStream", () => {
  it("reassembles a multi-chunk payload", () => {
    const wmf = emfPlusWmf([epHeader(), epEof()], 64);
    const out = embeddedEmfStream(wmf);
    expect(out).toBeDefined();
    expect(out!.length).toBeGreaterThan(64); // multi-chunk path actually taken
    expect(out![0]).toBe(1); // EMR_HEADER type of the carrier's first record
  });

  it("rejects bytes without WMFC escape chunks", () => {
    const bare = new Uint8Array(80);
    bare[76] = 3;
    expect(embeddedEmfStream(bare)).toBeUndefined();
  });
});

describe("emfPlusMembers", () => {
  it("replays a filled path into a member scaled to the box", () => {
    const wmf = emfPlusWmf([
      epHeader(),
      setWorld(1, 0, 0),
      brushObject(5, 0xffe8989a),
      pathObject(6, [
        [0, 0],
        [100, 0],
        [0, 50],
      ]),
      epRecord(FILL_PATH, 6, new Uint8Array(4)),
      epEof(),
    ]);
    const members = emfPlusMembers(wmf, 200, 100);
    expect(members).toBeDefined();
    const p = members!.find((m) => m.kind === "path");
    expect(p?.kind).toBe("path");
    if (p?.kind !== "path") return;
    expect(p.fill).toBe("e8989a");
    expect(p.d.startsWith("M")).toBe(true);
    // A lone path is normalized to fill the whole display box.
    expect(p.width).toBe(200);
    expect(p.height).toBe(100);
  });

  it("draws image points as picture members shaped by the parallelogram", () => {
    const pngMagic = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    // Two draws of the same object: second offset by (+40,+10).
    const drawAt = (ox: number, oy: number): Uint8Array => {
      const body = new Uint8Array(44);
      const v = new DataView(body.buffer);
      v.setUint32(28, 3, true); // parallelogram point count
      [
        [88, 92],
        [216, 92],
        [88, 127],
      ].forEach(([x, y], i) => {
        v.setInt16(32 + i * 4, x + ox, true);
        v.setInt16(34 + i * 4, y + oy, true);
      });
      return epRecord(DRAW_IMAGE_POINTS, 0x4000, body);
    };
    const wmf = emfPlusWmf([
      epHeader(),
      setWorld(1, 0, 0),
      imageObject(0, pngMagic),
      drawAt(0, 0),
      drawAt(40, 10),
      epEof(),
    ]);
    const members = emfPlusMembers(wmf, 605, 240);
    expect(members).toBeDefined();
    const pics = members!.filter((m) => m.kind === "picture");
    expect(pics.length).toBe(2);
    const [a0, b0] = pics;
    // Union box is (88..256)×(92..137): the equal-size pictures must split it
    // proportionally — 128 of 168 wide, offset 40.
    if (a0?.kind !== "picture" || b0?.kind !== "picture") return;
    for (const p of [a0, b0]) expect(p.src?.startsWith("data:image/png;base64,")).toBe(true);
    expect(b0.x - a0.x).toBeCloseTo((a0.width * 40) / 128, 0);
    expect(a0.width).toBeCloseTo(b0.width, 0);
    expect(a0.height).toBeCloseTo(b0.height, 0);
  });

  it("returns undefined for bytes without an embedded stream", () => {
    const bare = new Uint8Array(80);
    bare[76] = 3;
    expect(emfPlusMembers(bare, 100, 100)).toBeUndefined();
  });

  it("returns undefined when no record draws anything", () => {
    const wmf = emfPlusWmf([epHeader(), epEof()]);
    expect(emfPlusMembers(wmf, 100, 100)).toBeUndefined();
  });

  it("ignores an image-point record whose declared count exceeds its bytes", () => {
    const drawBody = new Uint8Array(44);
    new DataView(drawBody.buffer).setUint32(28, 99_999, true);
    const wmf = emfPlusWmf([
      epHeader(),
      imageObject(0, new Uint8Array([1, 2, 3])),
      epRecord(DRAW_IMAGE_POINTS, 0x4000, drawBody),
      epEof(),
    ]);
    expect(emfPlusMembers(wmf, 605, 240)).toBeUndefined();
  });
});
