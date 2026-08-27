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

/** ImageAttributes definition: complete object (stamped) but undecodable. */
function attrsObject(slot: number): Uint8Array {
  const body = new Uint8Array(12);
  const v = new DataView(body.buffer);
  v.setUint32(0, body.length, true); // chunk size
  v.setUint32(4, VERSION, true);
  return epRecord(OBJECT, slot | (8 << 8), body); // type 8 = image attributes
}

/** DrawImagePoints for one destination rectangle (i16 parallelogram). */
function drawRect(slot: number, x: number, y: number, w: number, h: number): Uint8Array {
  const body = new Uint8Array(44);
  const v = new DataView(body.buffer);
  v.setUint32(28, 3, true); // parallelogram point count
  [
    [x, y],
    [x + w, y],
    [x, y + h],
  ].forEach(([px, py], i) => {
    v.setInt16(32 + i * 4, px, true);
    v.setInt16(34 + i * 4, py, true);
  });
  return epRecord(DRAW_IMAGE_POINTS, 0x4000 | slot, body);
}

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

  it("reassembles a giant image from continued same-slot records", () => {
    // Real exporters split one image across same-slot records (each chunk
    // opening with its own size word), interleave foreign-slot definitions,
    // and draw right after the last chunk.
    const pngHead = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x11]);
    const pngTail = new Uint8Array([0x22, 0x33, 0x44]);
    const cont = 0x8000; // chunks carry it even on the last one
    const first = new Uint8Array(12 + pngHead.length);
    const fv = new DataView(first.buffer);
    fv.setUint32(0, first.length, true); // chunk size incl. shared header
    fv.setUint32(4, VERSION, true);
    fv.setUint32(8, 1, true); // image type: bitmap
    first.set(pngHead, 12);
    const tail = new Uint8Array(8 + pngTail.length);
    const tv = new DataView(tail.buffer);
    tv.setUint32(0, tail.length, true); // own chunk size
    tv.setUint32(4, 93004, true); // repeated object total (ignored)
    tail.set(pngTail, 8);

    const wmf = emfPlusWmf([
      epHeader(),
      setWorld(1, 0, 0),
      epRecord(OBJECT, 1 | (5 << 8) | cont, first),
      attrsObject(2), // interleaved mid-run foreign definition
      epRecord(OBJECT, 1 | (5 << 8) | cont, tail),
      drawRect(1, 10, 20, 30, 40),
      epEof(),
    ]);
    const members = emfPlusMembers(wmf, 605, 240);
    expect(members).toBeDefined();
    const pic = members!.find((m) => m.kind === "picture");
    if (pic?.kind !== "picture" || !pic.src?.startsWith("data:image/png;base64,")) return;
    const joined = Buffer.from(pic.src.slice("data:image/png;base64,".length), "base64");
    expect(joined.equals(Buffer.concat([pngHead, pngTail]))).toBe(true);
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

  // Path objects whose point format flags come straight from corpus files
  // (word/media of a real dual-mode document).
  describe("path point formats", () => {
    /** Path object body with explicit format word and int16 point pairs. */
    function rawPath(
      pts: Array<[number, number]>,
      fmt: number,
      i16: boolean,
      types?: number[],
    ): Uint8Array {
      const step = i16 ? 4 : 8;
      const tb = types ?? pts.map((_, i) => (i === 0 ? 0x00 : 0x01));
      const body = new Uint8Array(16 + pts.length * step + tb.length);
      const v = new DataView(body.buffer);
      v.setUint32(0, body.length, true); // TotalObjectSize
      v.setUint32(4, VERSION, true);
      v.setUint32(8, pts.length, true); // PathPointCount
      v.setUint32(12, fmt, true); // PathPointFlags
      pts.forEach(([x, y], i) => {
        if (i16) {
          v.setInt16(16 + i * 4, x, true);
          v.setInt16(18 + i * 4, y, true);
        } else {
          v.setFloat32(16 + i * 8, x, true);
          v.setFloat32(20 + i * 8, y, true);
        }
      });
      body.set(tb, 16 + pts.length * step);
      return epRecord(OBJECT, 6 | (3 << 8), body);
    }

    it("reads int16 pairs under the 0x4000 compression bit", () => {
      const rect: Array<[number, number]> = [
        [10, 10],
        [110, 10],
        [110, 60],
        [10, 60],
      ];
      const wmf = emfPlusWmf([
        epHeader(),
        setWorld(1, 0, 0),
        brushObject(5, 0xff38761d),
        rawPath(rect, 0x4000, true, [0x00, 0x01, 0x01, 0x81]),
        epRecord(FILL_PATH, 6, new Uint8Array(4)),
        epEof(),
      ]);
      const members = emfPlusMembers(wmf, 200, 100);
      expect(members).toBeDefined();
      const p = members!.find((m) => m.kind === "path");
      expect(p).toBeDefined(); // compressed readback must produce geometry
      if (p?.kind !== "path") return;
      expect(p.fill).toBe("38761d");
      expect(p.d.startsWith("M")).toBe(true);
      expect(p.d).toContain("Z"); // closed figure made it through
    });

    it("decodes float pairs when only the 0x2000 flag bit is set", () => {
      // Corpus census: every real object carrying just this bit stores float
      // pairs — the bit is NOT the compression switch.
      const rect: Array<[number, number]> = [
        [10, 10],
        [110, 10],
        [110, 60],
        [10, 60],
      ];
      const wmf = emfPlusWmf([
        epHeader(),
        setWorld(1, 0, 0),
        brushObject(5, 0xff38761d),
        rawPath(rect, 0x2000, false),
        epRecord(FILL_PATH, 6, new Uint8Array(4)),
        epEof(),
      ]);
      const members = emfPlusMembers(wmf, 200, 100);
      expect(members).toBeDefined();
      const p = members!.find((m) => m.kind === "path");
      expect(p).toBeDefined(); // float pairs must survive the flag
      if (p?.kind !== "path") return;
      expect(p.d.startsWith("M")).toBe(true);
    });

    it("drops geometry that never reaches a start-type point", () => {
      const rect: Array<[number, number]> = [
        [10, 10],
        [110, 10],
        [110, 60],
        [10, 60],
      ];
      const wmf = emfPlusWmf([
        epHeader(),
        setWorld(1, 0, 0),
        brushObject(5, 0xff38761d),
        rawPath(rect, 0, false, [0x01, 0x01, 0x01, 0x01]),
        epRecord(FILL_PATH, 6, new Uint8Array(4)),
        epEof(),
      ]);
      expect(emfPlusMembers(wmf, 200, 100)).toBeUndefined();
    });
  });
});
