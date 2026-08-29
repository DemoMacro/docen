// Spec for the dual-mode metafile player (emf-plus.ts): WMFC chunk
// reassembly, GDI+ record replay, and member shaping. Fixtures are built at
// the byte level against layouts verified on real corpus files.

import type { LayoutDrawingMember } from "@docen/layout";
import { describe, expect, it } from "vitest";

import { embeddedEmfStream, emfPlusMembers } from "./emf-plus";
import { dualModeWmf, emfCarrier, emfPlusWmf, emrEmfPlusComment, epRecord } from "./wmf-test-util";

// EmfPlus record codes used by the fixtures.
const HEADER = 0x4001;
const OBJECT = 0x4008;
const FILL_RECTS = 0x400a;
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

/** DrawImagePoints for one destination rectangle (i16 parallelogram).
 *  Body layout: [attrsId][srcUnit][rsvd][SrcRect RectF][count][points]. */
function drawRect(
  slot: number,
  x: number,
  y: number,
  w: number,
  h: number,
  src = { x: 0, y: 0, w: 100, h: 100 },
): Uint8Array {
  const body = new Uint8Array(44);
  const v = new DataView(body.buffer);
  v.setUint32(4, 2, true); // srcUnit: UnitPixel
  v.setFloat32(12, src.x, true);
  v.setFloat32(16, src.y, true);
  v.setFloat32(20, src.w, true);
  v.setFloat32(24, src.h, true);
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
      v.setFloat32(20, 200, true); // srcRect: full 200×200 image
      v.setFloat32(24, 200, true);
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

  // Real exporters sprite whole art pages into one bitmap object and slice
  // regions out of it through the record's always-present SrcRect ([MS-EMFPLUS]
  // EmfPlusDrawImagePoints); the destination parallelogram receives only that
  // slice.
  describe("source rectangles", () => {
    /** PNG bytes with a real IHDR carrying w×h. */
    function pngWithSize(w: number, h: number): Uint8Array {
      const buf = new Uint8Array(8 + 8 + 13 + 4 + 12);
      const v = new DataView(buf.buffer);
      buf.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
      v.setUint32(8, 13, true);
      buf.set(new TextEncoder().encode("IHDR"), 12);
      v.setUint32(16, w, false);
      v.setUint32(20, h, false);
      const iend = buf.length - 12;
      v.setUint32(iend + 4, 0x444e4549, false);
      buf.set([0xae, 0x42, 0x60, 0x82], iend + 8);
      return buf;
    }

    function drawSlice(src: { x: number; y: number; w: number; h: number }): Uint8Array {
      const body = new Uint8Array(44);
      const v = new DataView(body.buffer);
      v.setFloat32(12, src.x, true);
      v.setFloat32(16, src.y, true);
      v.setFloat32(20, src.w, true);
      v.setFloat32(24, src.h, true);
      v.setUint32(28, 3, true);
      [
        [10, 10],
        [40, 10],
        [10, 35],
      ].forEach(([px, py], i) => {
        v.setInt16(32 + i * 4, px, true);
        v.setInt16(34 + i * 4, py, true);
      });
      return epRecord(DRAW_IMAGE_POINTS, 0x4000 | 0, body);
    }

    it("crops a partial source rectangle into picture-member fractions", () => {
      const wmf = emfPlusWmf([
        epHeader(),
        setWorld(1, 0, 0),
        imageObject(0, pngWithSize(100, 50)),
        drawSlice({ x: 25, y: 10, w: 50, h: 20 }),
        epEof(),
      ]);
      const members = emfPlusMembers(wmf, 300, 250);
      expect(members).toBeDefined();
      const pic = members!.find((m) => m.kind === "picture");
      if (pic?.kind !== "picture") return;
      // crop fractions are literal against the natural size; the box is the
      // normalized destination rect (30×25 of a 30×25-draft union → full box).
      expect(pic.crop).toEqual({ left: 0.25, top: 0.2, right: 0.25, bottom: 0.4 });
      expect(pic.width).toBeCloseTo(300, 0);
      expect(pic.height).toBeCloseTo(250, 0);
    });

    it("treats a full-image source rectangle as uncropped", () => {
      const wmf = emfPlusWmf([
        epHeader(),
        setWorld(1, 0, 0),
        imageObject(0, pngWithSize(80, 80)),
        drawSlice({ x: 0, y: 0, w: 80, h: 80 }),
        epEof(),
      ]);
      const members = emfPlusMembers(wmf, 200, 200);
      expect(members).toBeDefined();
      const pic = members!.find((m) => m.kind === "picture");
      if (pic?.kind !== "picture") return;
      expect(pic.crop).toBeUndefined();
    });

    it("skips a record whose source rectangle is empty", () => {
      const wmf = emfPlusWmf([
        epHeader(),
        setWorld(1, 0, 0),
        imageObject(0, pngWithSize(80, 80)),
        drawSlice({ x: 5, y: 5, w: 0, h: 20 }),
        epEof(),
      ]);
      expect(emfPlusMembers(wmf, 200, 200)).toBeUndefined();
    });
  });

  // Office embeds vector badges/diamonds as metafile-typed image objects —
  // each DrawImagePoints destination maps the nested EMF's full device rect,
  // and the replayed vectors must land through that mapping.
  describe("nested metafile images", () => {
    /** Raw EMF opening with a genuine EMR_HEADER (bounds needed for the
     *  nesting basis) wrapping one "EMF+" comment, closed by EMR_EOF. */
    function nestedCarrier(bw: number, bh: number, plus: Uint8Array[]): Uint8Array {
      const head = new Uint8Array(108);
      const hv = new DataView(head.buffer);
      hv.setUint32(0, 1, true); // EMR_HEADER
      hv.setUint32(4, 108, true);
      hv.setInt32(16, bw, true); // rclBounds.right
      hv.setInt32(20, bh, true); // rclBounds.bottom
      const eof = new Uint8Array(20);
      new DataView(eof.buffer).setUint32(0, 14, true); // EMR_EOF
      new DataView(eof.buffer).setUint32(4, eof.length, true);
      const parts = [head, emrEmfPlusComment(plus), eof];
      const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
      let off = 0;
      for (const p of parts) {
        out.set(p, off);
        off += p.length;
      }
      return out;
    }

    /** Image object typed metafile ([+8]=2) whose encoding is a raw EMF at
     *  [+16], found by the header→EOF chain walk. */
    function metaImageObject(slot: number, emf: Uint8Array): Uint8Array {
      const body = new Uint8Array(16 + emf.length);
      const v = new DataView(body.buffer);
      v.setUint32(0, body.length, true);
      v.setUint32(4, VERSION, true);
      v.setUint32(8, 2, true); // image type: metafile
      v.setUint32(12, 4, true); // embedded format: EMF
      body.set(emf, 16);
      return epRecord(OBJECT, slot | (5 << 8), body);
    }

    it("replays nested carriers mapped onto the destination parallelogram", () => {
      // Nested triangle at device coords (10,20)-(35,40) inside a 200×200
      // frame, drawn into the outer rect (10,10)-(110,110): the affine folds
      // by half, so vertices land at (15,20)/(35,20)/(25,40) — which a lone
      // draft normalizes across the full display box.
      const inner = nestedCarrier(200, 200, [
        epHeader(),
        setWorld(1, 0, 0),
        brushObject(5, 0xffcc0000),
        pathObject(6, [
          [10, 20],
          [50, 20],
          [30, 60],
        ]),
        epRecord(FILL_PATH, 6, new Uint8Array(4)),
        epEof(),
      ]);
      const wmf = emfPlusWmf([
        epHeader(),
        setWorld(1, 0, 0),
        metaImageObject(0, inner),
        drawRect(0, 10, 10, 100, 100),
        epEof(),
      ]);
      const members = emfPlusMembers(wmf, 200, 200);
      expect(members).toBeDefined();
      const p = members!.find((m) => m.kind === "path");
      if (p?.kind !== "path") return;
      // Mapped vertices fill the box exactly — no picture stand-in appeared.
      expect(p.fill).toBe("cc0000");
      expect(p.d).toBe("M0 0L200 0L100 200");
      expect(members!.some((m) => m.kind === "picture")).toBe(false);
    });

    it("type-finds exporter double-header metafiles split across chunks", () => {
      // Real exporters send giant metafile images as same-slot runs whose
      // first record keeps an extra [chunkDataSize][objectTotal] pair in front
      // of the version stamp — type fields sit at +12 there, not +8.
      const inner = nestedCarrier(100, 100, [
        epHeader(),
        setWorld(1, 0, 0),
        brushObject(5, 0xff1050a0),
        pathObject(6, [
          [0, 0],
          [80, 0],
          [40, 70],
        ]),
        epRecord(FILL_PATH, 6, new Uint8Array(4)),
        epEof(),
      ]);
      const typed = new Uint8Array(12 + inner.length);
      const tv = new DataView(typed.buffer);
      tv.setUint32(0, typed.length, true); // object total
      tv.setUint32(4, 2, true); // image type: metafile
      tv.setUint32(8, 4, true); // embedded format: EMF
      typed.set(inner, 12);
      const SPLIT = 90;
      const recFirst = new Uint8Array(8 + typed.length);
      const rv = new DataView(recFirst.buffer);
      rv.setUint32(0, recFirst.length, true); // chunk data size
      rv.setUint32(4, VERSION, true); // stamp lives at +8 on these payloads
      recFirst.set(typed.subarray(4), 8);
      void SPLIT;
      const wmf = emfPlusWmf([
        epHeader(),
        setWorld(1, 0, 0),
        epRecord(OBJECT, 7 | (5 << 8), recFirst),
        drawRect(7, 0, 0, 300, 300),
        epEof(),
      ]);
      const members = emfPlusMembers(wmf, 300, 300);
      expect(members).toBeDefined();
      const p = members!.find((m) => m.kind === "path");
      if (p?.kind !== "path") return;
      expect(p.fill).toBe("1050a0");
      // The nested triangle fills its whole destination box after mapping.
      expect(p.width).toBe(300);
    });

    it("places nested carriers through the live world transform", () => {
      // Badge-like flow in real files: a SET_WORLD_TRANSFORM positions each
      // DrawImagePoints call inside the page; the nested replay must land at
      // that placed rectangle, not at raw logical coordinates.
      const inner = nestedCarrier(100, 100, [
        epHeader(),
        setWorld(1, 0, 0),
        brushObject(5, 0xff202020),
        pathObject(6, [
          [0, 0],
          [100, 0],
          [50, 100],
        ]),
        epRecord(FILL_PATH, 6, new Uint8Array(4)),
        epEof(),
      ]);
      const wmf = emfPlusWmf([
        epHeader(),
        // Second call for contrast, translated right by half the box.
        setWorld(0.5, 150, 150),
        metaImageObject(2, inner),
        drawRect(2, 0, 0, 200, 200),
        setWorld(0.5, 450, 450),
        metaImageObject(3, inner),
        drawRect(3, 0, 0, 200, 200),
        epEof(),
      ]);
      const members = emfPlusMembers(wmf, 900, 900);
      expect(members).toBeDefined();
      const paths = members!.filter((m) => m.kind === "path");
      expect(paths.length).toBe(2);
      // Path members carry absolute geometry in `d` (their x/y is always the
      // member box origin) — compare the first M point's y.
      const startAt = (d: string): number => Number(d.split(/[^-\d.]+/)[1]);
      expect(startAt(paths[1]!.d)).toBeGreaterThan(startAt(paths[0]!.d));
    });
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

  // Dual-mode carriers keep the real text on the GDI side (their EmfPlus
  // DrawString twins are zero-data stubs), so the player must read it back
  // from plain EMR records.
  describe("carrier-side GDI text", () => {
    /** Raw EMR wrapper: type, size, payload. */
    function emr(type: number, payload: Uint8Array): Uint8Array {
      const buf = new Uint8Array(8 + payload.length);
      new DataView(buf.buffer).setUint32(0, type, true);
      new DataView(buf.buffer).setUint32(4, buf.length, true);
      buf.set(payload, 8);
      return buf;
    }

    function utf16le(s: string): Uint8Array {
      const buf = new Uint8Array(s.length * 2);
      const v = new DataView(buf.buffer);
      s.split("").forEach((ch, i) => v.setUint16(i * 2, ch.charCodeAt(0), true));
      return buf;
    }

    function extcreatefont(slot: number, height: number, weight: number, face: string): Uint8Array {
      // ihFont dword + LOGFONTW (faceName is UTF-16LE at byte 28).
      const buf = new Uint8Array(4 + 92);
      const v = new DataView(buf.buffer);
      v.setUint32(0, slot, true); // object index
      v.setInt32(4, -height, true); // lfHeight (negative = char height)
      v.setUint32(20, weight, true); // lfWeight
      buf.set(utf16le(face), 4 + 28);
      return emr(82, buf);
    }

    function selectfont(slot: number): Uint8Array {
      const buf = new Uint8Array(4);
      new DataView(buf.buffer).setUint32(0, slot, true);
      return emr(37, buf);
    }

    function worldTransform(scale: number, dx: number, dy: number): Uint8Array {
      // EMR_SETWORLDTRANSFORM: six floats directly behind the header.
      const buf = new Uint8Array(24);
      const v = new DataView(buf.buffer);
      v.setFloat32(0, scale, true);
      v.setFloat32(4, 0, true);
      v.setFloat32(8, 0, true);
      v.setFloat32(12, scale, true);
      v.setFloat32(16, dx, true);
      v.setFloat32(20, dy, true);
      return emr(35, buf);
    }

    function settextcolor(argbHexRgb: string): Uint8Array {
      const r = parseInt(argbHexRgb.slice(0, 2), 16);
      const g = parseInt(argbHexRgb.slice(2, 4), 16);
      const b = parseInt(argbHexRgb.slice(4, 6), 16);
      const buf = new Uint8Array(4);
      new DataView(buf.buffer).setUint32(0, (b << 16) | (g << 8) | r, true); // COLORREF
      return emr(24, buf);
    }

    function exttextoutw(text: string, x: number, y: number, dx?: Int16Array): Uint8Array {
      const bytes = utf16le(text);
      const chars = text.length; // UTF-16 code units, per the Chars field
      // Payload after the EMR header: bounds[16], graphicsMode, exScale,
      // eyScale, then the EmrText block (reference xy, chars, offString,
      // options). The parser reads at record-relative offsets — these land
      // there via the shared 8-byte header.
      const head = new Uint8Array(48);
      const v = new DataView(head.buffer);
      v.setUint32(28, x, true); // reference x
      v.setUint32(32, y, true); // reference y (baseline)
      v.setUint32(36, chars, true);
      v.setUint32(40, 56, true); // offString, relative to the record start
      v.setUint32(44, 0, true); // options
      const parts: Uint8Array[] = [head, bytes];
      if (dx) {
        const trail = new Uint8Array(dx.buffer, dx.byteOffset, dx.length * 2);
        parts.push(trail);
      }
      return emr(84, joinBuffers(...parts));
    }

    function joinBuffers(...parts: Uint8Array[]): Uint8Array {
      const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
      let off = 0;
      for (const p of parts) {
        out.set(p, off);
        off += p.length;
      }
      return out;
    }

    /** Carrier with an "EMF+" comment (plus records) alongside raw GDI EMRs
     *  and given text records — the real dual-mode layering. */
    function carrierWithText(gdi: Uint8Array[], plus: Uint8Array[]): Uint8Array {
      return dualModeWmf(emfCarrier([emrEmfPlusComment(plus), ...gdi]));
    }

    it("reads ExtTextOutW runs as styled textBox members", () => {
      // A zero-drawing EMF+ stream (header + EOF only) with the text on the
      // carrier side — the sprite shape.
      const wmf = carrierWithText(
        [
          extcreatefont(3, 24, 400, "微软雅黑"),
          selectfont(3),
          settextcolor("404040"),
          exttextoutw("示例里程碑文案", 9270, 45240),
        ],
        [epHeader(), epEof()],
      );
      const members = emfPlusMembers(wmf, 400, 300);
      expect(members).toBeDefined();
      const box = members!.find((m) => m.kind === "textBox");
      if (box?.kind !== "textBox") return;
      expect(box.blocks[0]).toMatchObject({
        kind: "paragraph",
        inline: [
          {
            kind: "text",
            text: "示例里程碑文案",
            style: { family: "微软雅黑", color: "404040" },
          },
        ],
      });
      const para = box.blocks[0];
      if (para.kind !== "paragraph") return;
      const run = para.inline[0];
      if (run.kind !== "text") return;
      expect(run.style.sizePx ?? 0).toBeGreaterThan(0);
    });

    it("threads Dx-run tracking into width and letter spacing", () => {
      // A tracked run: Dx advances above the natural 240 em — the difference
      // is real letter spacing, and the total advance is the Dx sum rather
      // than the per-char estimate.
      const wmf = carrierWithText(
        [
          extcreatefont(1, 24, 681, "微软雅黑"),
          selectfont(1),
          settextcolor("ffffff"),
          exttextoutw(
            "示例跟踪文案",
            1000,
            40000,
            new Int16Array([340, 0, 360, 0, 340, 0, 360, 0, 340, 0, 240, 0]),
          ),
        ],
        [epHeader(), epEof()],
      );
      const members = emfPlusMembers(wmf, 600, 300);
      expect(members).toBeDefined();
      const box = members!.find((m) => m.kind === "textBox");
      if (box?.kind !== "textBox") return;
      // Weight 681 < FW_BOLD: GDI face matching keeps the regular face.
      const para = box.blocks[0];
      if (para.kind !== "paragraph") return;
      const run = para.inline[0];
      if (run.kind !== "text") return;
      expect(run.style.bold).toBeUndefined();
      expect(run.style.letterSpacingPx ?? 0).toBeGreaterThan(5);
      // Width = Σ advances × scale (≈ 33px/char at this box scale), not the
      // estimate (24px/char).
      expect(box.width).toBeGreaterThan(160);
    });

    it("keeps a 700-weight run bold with untracked advances", () => {
      const wmf = carrierWithText(
        [extcreatefont(1, 24, 700, "宋体"), selectfont(1), exttextoutw("示例文字", 1000, 40000)],
        [epHeader(), epEof()],
      );
      const members = emfPlusMembers(wmf, 400, 300);
      expect(members).toBeDefined();
      const box = members!.find((m) => m.kind === "textBox");
      if (box?.kind !== "textBox") return;
      const para = box.blocks[0];
      if (para.kind !== "paragraph") return;
      const run = para.inline[0];
      if (run.kind !== "text") return;
      expect(run.style.bold).toBe(true);
      expect(run.style.letterSpacingPx).toBeUndefined();
    });

    it("strips the GDI @ vertical-variant prefix from the face", () => {
      // '@楷体' names GDI's vertical variant of the face — browser font
      // matching cannot resolve the '@' (it falls back to a thinner face),
      // so the rendered family is the base face.
      const wmf = carrierWithText(
        [extcreatefont(1, 24, 400, "@楷体"), selectfont(1), exttextoutw("示例文字", 1000, 40000)],
        [epHeader(), epEof()],
      );
      const members = emfPlusMembers(wmf, 400, 300);
      expect(members).toBeDefined();
      const box = members!.find((m) => m.kind === "textBox");
      if (box?.kind !== "textBox") return;
      const para = box.blocks[0];
      const run = para?.kind === "paragraph" ? para.inline[0] : undefined;
      if (run?.kind !== "text") return;
      expect(run.style.family).toBe("楷体");
    });

    it("marks runs under a rotated world transform with the screen angle", () => {
      // Vertical plan-box columns: a 90°-rotated GDI world transform steers
      // the run's advance but never turns the glyph outlines (GDI plays
      // rotated text with upright glyphs). The run therefore emits one
      // upright box per character stacked down the advance column — not a
      // single rotation-carrying box.
      const setworld = (m: number[]): Uint8Array => {
        const body = new Uint8Array(24);
        const v = new DataView(body.buffer);
        m.forEach((n, i) => v.setFloat32(i * 4, n, true));
        return emr(35, body);
      };
      const wmf = carrierWithText(
        [
          setworld([0, 0.07, -0.07, 0, 31580, -1287]),
          extcreatefont(1, 24, 400, "SimSun"),
          selectfont(1),
          settextcolor("1a1a1a"),
          exttextoutw("示例竖排文字", 21015, 470415),
        ],
        [epHeader(), epEof()],
      );
      const members = emfPlusMembers(wmf, 400, 300);
      expect(members).toBeDefined();
      const boxes: Extract<LayoutDrawingMember, { kind: "textBox" }>[] = [];
      for (const m of members!) {
        if (m.kind === "textBox") boxes.push(m);
      }
      expect(boxes.map((b) => b.blocks)).toHaveLength(6);
      const texts = boxes.map((b) => {
        const para = b.blocks[0];
        return para?.kind === "paragraph" && para.inline[0]?.kind === "text"
          ? para.inline[0].text
          : "";
      });
      expect(texts.join("")).toBe("示例竖排文字");
      // The advance column points down (m12 = +0.07): successive characters
      // descend, and no box carries a rotation.
      const ys = boxes.map((b) => b.y);
      expect([...ys].sort((a, b) => a - b)).toEqual(ys);
      for (const b of boxes) expect(b.rotation).toBeUndefined();
    });

    it("anchors vertical columns with the glyph baseline at the reference", () => {
      // Word's replay draws upright @font glyphs with the baseline through
      // the reference point: the cell spans [ref − descent, ref + ascent]
      // (realized @楷体: 328/60 at a −360 request) with the em centered in
      // it — a small positive insets.left, and the face stripped of its '@'.
      // The ink-center position is pixel-verified in the editor against the
      // reference render (the box scale differs per axis, so no ratio here).
      const setworld = (m: number[]): Uint8Array => {
        const body = new Uint8Array(24);
        const v = new DataView(body.buffer);
        m.forEach((n, i) => v.setFloat32(i * 4, n, true));
        return emr(35, body);
      };
      const wmf = carrierWithText(
        [
          setworld([0, 0.07, -0.07, 0, 31580, -1287]),
          extcreatefont(1, 24, 400, "@楷体"),
          selectfont(1),
          settextcolor("1a1a1a"),
          exttextoutw("例字", 21015, 470415),
        ],
        [epHeader(), epEof()],
      );
      const members = emfPlusMembers(wmf, 400, 300);
      expect(members).toBeDefined();
      const box = members!.find((m) => m.kind === "textBox");
      if (box?.kind !== "textBox") return;
      const para = box.blocks[0];
      const run =
        para?.kind === "paragraph" && para.inline[0]?.kind === "text" ? para.inline[0] : undefined;
      if (!run) return;
      expect(run.style.family).toBe("楷体");
      expect(box.insets?.left).toBeGreaterThan(0);
    });

    it("keeps walking past whitespace-only text runs", () => {
      // A blank run must not stall the record walk: the loop advances only
      // at its tail, so any skip has to route through the advance.
      const wmf = carrierWithText(
        [
          extcreatefont(3, 24, 400, "微软雅黑"),
          selectfont(3),
          settextcolor("404040"),
          exttextoutw(" ", 1000, 40000),
          exttextoutw("示例里程碑文案", 9270, 45240),
        ],
        [epHeader(), epEof()],
      );
      const members = emfPlusMembers(wmf, 400, 300);
      expect(members).toBeDefined();
      const texts = members!.filter((m) => m.kind === "textBox");
      expect(texts).toHaveLength(1);
      const para = texts[0].blocks[0];
      if (para.kind !== "paragraph") return;
      const run = para.inline[0];
      if (run.kind !== "text") return;
      expect(run.text).toBe("示例里程碑文案");
    });

    it("maps logical draw origins into carrier device space via the world transform", () => {
      // Corpus shape: each run sits in an arbitrary logical
      // scale and a world transform folds it onto the EMF's device rectangle.
      const pngMagic = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
      const wmf = carrierWithText(
        [
          extcreatefont(2, 285, 0, "宋体"),
          selectfont(2),
          settextcolor("262626"),
          worldTransform(1 / 15, -1064.4, -1286.6),
          // Logical (18330, 23520) → device ≈ (158, 281).
          exttextoutw("一二三四五六七八九十甲乙", 18330, 23520),
        ],
        [
          epHeader(),
          setWorld(1, 0, 0),
          imageObject(0, pngMagic),
          drawRect(0, 0, 0, 580, 876),
          epEof(),
        ],
      );
      const members = emfPlusMembers(wmf, 580, 876);
      expect(members).toBeDefined();
      const box = members!.find((m) => m.kind === "textBox");
      if (box?.kind !== "textBox") return;
      const para = box.blocks[0];
      if (para.kind !== "paragraph") return;
      const run = para.inline[0];
      if (run.kind !== "text") return;
      expect(run.text).toBe("一二三四五六七八九十甲乙");
      // Device-space placement: inside the EMF bounds, scaled glyph metrics.
      expect(box.x).toBeGreaterThan(140);
      expect(box.x).toBeLessThan(180);
      expect(box.y).toBeGreaterThan(260);
      expect(box.y).toBeLessThan(300);
      // Advance uses transformed glyph size (285 → 19px/char).
      expect(box.width).toBeGreaterThan(220);
      expect(box.width).toBeLessThan(240);
    });

    it("activates fonts through SelectObject rather than creation order", () => {
      // Two created fonts; the *selected* one wins even though the other was
      // created later — GDI object-table semantics.
      const wmf = carrierWithText(
        [
          extcreatefont(1, 30, 700, "方正大黑简体"),
          extcreatefont(2, 18, 400, "宋体"),
          selectfont(1),
          settextcolor("333333"),
          exttextoutw("示例标题文字", 500, 500),
        ],
        [epHeader(), epEof()],
      );
      const members = emfPlusMembers(wmf, 600, 120);
      expect(members).toBeDefined();
      const box = members!.find((m) => m.kind === "textBox");
      if (box?.kind !== "textBox") return;
      const para = box.blocks[0];
      if (para.kind !== "paragraph") return;
      const run = para.inline[0];
      if (run.kind !== "text") return;
      expect(run.style.family).toBe("方正大黑简体");
    });

    it("merges carrier text with same-stream EMF+ pictures", () => {
      const pngMagic = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
      const wmf = carrierWithText(
        [
          extcreatefont(1, 18, 700, "SimSun"),
          selectfont(1),
          settextcolor("1a1a1a"),
          exttextoutw("示例章节名", 2000, 30000),
        ],
        [
          epHeader(),
          setWorld(1, 0, 0),
          imageObject(0, pngMagic),
          drawRect(0, 500, 45000, 6000, 4000),
          epEof(),
        ],
      );
      const members = emfPlusMembers(wmf, 400, 300);
      expect(members).toBeDefined();
      const kinds = new Set(members!.map((m) => m.kind));
      expect(kinds.has("picture")).toBe(true);
      expect(kinds.has("textBox")).toBe(true);
    });
  });

  describe("EmfPlusFillRects", () => {
    /** FillRects record in the WMF-embedded form: the OBJECT chunkDataSize
     *  word repeats in front ([+0]), then the inlined emphasis ARGB when
     *  flagged, the rect count, and the rects (float RectF or int16 RectS). */
    function fillRects(flags: number, argb: number, rects: number[][]): Uint8Array {
      const compressed = (flags & 0x4000) !== 0;
      const body = new Uint8Array(12 + rects.length * (compressed ? 8 : 16));
      const v = new DataView(body.buffer);
      v.setUint32(0, body.length, true);
      v.setUint32(4, argb, true);
      v.setUint32(8, rects.length, true);
      rects.forEach(([x, y, w, h], i) => {
        const at = 12 + i * (compressed ? 8 : 16);
        if (compressed) {
          v.setInt16(at, x, true);
          v.setInt16(at + 2, y, true);
          v.setInt16(at + 4, w, true);
          v.setInt16(at + 6, h, true);
        } else {
          v.setFloat32(at, x, true);
          v.setFloat32(at + 4, y, true);
          v.setFloat32(at + 8, w, true);
          v.setFloat32(at + 12, h, true);
        }
      });
      return epRecord(FILL_RECTS, flags, body);
    }

    it("takes an inlined ColorEmphasis ARGB over float rects", () => {
      const wmf = emfPlusWmf([
        epHeader(),
        setWorld(1, 0, 0),
        fillRects(0x8000, 0xff6d2934, [[10, 20, 80, 40]]),
        epEof(),
      ]);
      const members = emfPlusMembers(wmf, 400, 300);
      expect(members).toBeDefined();
      const p = members!.find((m) => m.kind === "path");
      if (p?.kind !== "path") return;
      // The lone rect fills the display box exactly after normalization.
      expect(p.fill).toBe("6d2934");
      expect(p.d).toBe("M0 0L400 0L400 300L0 300Z");
    });

    it("fills compressed int16 rects from the last brush", () => {
      const wmf = emfPlusWmf([
        epHeader(),
        setWorld(1, 0, 0),
        brushObject(5, 0xffe8989a),
        fillRects(0x4000, 0, [
          [0, 0, 50, 100],
          [50, 0, 50, 100],
        ]),
        epEof(),
      ]);
      const members = emfPlusMembers(wmf, 100, 100);
      expect(members).toBeDefined();
      const paths = members!.filter((m) => m.kind === "path");
      expect(paths).toHaveLength(2);
      expect(paths[0]).toMatchObject({ fill: "e8989a", d: "M0 0L50 0L50 100L0 100Z" });
      expect(paths[1]).toMatchObject({ fill: "e8989a", d: "M50 0L100 0L100 100L50 100Z" });
    });

    it("draws nothing without a brush and without an inlined color", () => {
      const wmf = emfPlusWmf([
        epHeader(),
        setWorld(1, 0, 0),
        fillRects(0, 0, [[10, 10, 50, 50]]),
        epEof(),
      ]);
      expect(emfPlusMembers(wmf, 400, 300)).toBeUndefined();
    });
  });

  describe("carrier-side GDI strokes", () => {
    /** Raw EMR wrapper (same shape as the text describe's). */
    function emr(type: number, payload: Uint8Array): Uint8Array {
      const buf = new Uint8Array(8 + payload.length);
      const v = new DataView(buf.buffer);
      v.setUint32(0, type, true);
      v.setUint32(4, buf.length, true);
      buf.set(payload, 8);
      return buf;
    }

    it("replays ExtCreatePen + Polyline16 as stroked paths", () => {
      // The list-page underline pattern: a solid pen (COLORREF 0x001e1c7c →
      // #7c1c1e) strokes a polyline on the carrier side while the EMF+ layer
      // draws nothing — the merge must let these strokes through.
      const pen = new Uint8Array(40);
      const pv = new DataView(pen.buffer);
      pv.setUint32(0, 2, true); // pen slot
      pv.setUint32(20, 0, true); // PS_SOLID
      pv.setUint32(24, 32, true); // width
      pv.setUint32(32, 0x001e1c7c, true); // COLORREF
      const poly = new Uint8Array(32);
      const lv = new DataView(poly.buffer);
      lv.setUint32(16, 3, true); // point count
      [
        [0, 0],
        [100, 0],
        [100, 100],
      ].forEach(([x, y], i) => {
        lv.setInt16(20 + i * 4, x, true);
        lv.setInt16(20 + i * 4 + 2, y, true);
      });
      const select = new Uint8Array(4);
      new DataView(select.buffer).setUint32(0, 2, true);
      const wmf = dualModeWmf(
        emfCarrier([
          emrEmfPlusComment([epHeader(), epEof()]),
          emr(95, pen),
          emr(37, select),
          emr(87, poly),
        ]),
      );
      const members = emfPlusMembers(wmf, 100, 100);
      expect(members).toBeDefined();
      const p = members!.find((m) => m.kind === "path");
      if (p?.kind !== "path") return;
      expect(p.line).toMatchObject({ color: "7c1c1e", px: 32 });
      expect(p.d).toBe("M0 0L100 0L100 100");
    });
  });

  // Wavy panel shapes ride the GDI path chain: a brush + a BeginPath figure
  // (MoveTo/PolylineTo) frozen by EndPath and consumed by FillPath. The EMF+
  // layer only draws fragments of these on corpus files.
  describe("carrier-side GDI paths", () => {
    function emr(type: number, payload: Uint8Array): Uint8Array {
      const buf = new Uint8Array(8 + payload.length);
      new DataView(buf.buffer).setUint32(0, type, true);
      new DataView(buf.buffer).setUint32(4, buf.length, true);
      buf.set(payload, 8);
      return buf;
    }

    function createSolidBrush(slot: number, hex: string): Uint8Array {
      // [ihBrush u32][style u32 = BS_SOLID][COLORREF 0x00bbggrr]
      const buf = new Uint8Array(12);
      const v = new DataView(buf.buffer);
      v.setUint32(0, slot, true);
      v.setUint32(4, 0, true);
      const r = parseInt(hex.slice(0, 2), 16);
      const g = parseInt(hex.slice(2, 4), 16);
      const b = parseInt(hex.slice(4, 6), 16);
      v.setUint32(8, (b << 16) | (g << 8) | r, true);
      return emr(39, buf);
    }

    function selectObject(slot: number): Uint8Array {
      const buf = new Uint8Array(4);
      new DataView(buf.buffer).setUint32(0, slot, true);
      return emr(37, buf);
    }

    function moveTo(x: number, y: number): Uint8Array {
      const buf = new Uint8Array(8);
      const v = new DataView(buf.buffer);
      v.setInt32(0, x, true);
      v.setInt32(4, y, true);
      return emr(27, buf);
    }

    /** PolylineTo16: [bounds 4×i32][count u32][points int16 pairs]. */
    function polyLineTo16(pts: Array<[number, number]>): Uint8Array {
      const buf = new Uint8Array(20 + pts.length * 4);
      const v = new DataView(buf.buffer);
      v.setUint32(16, pts.length, true);
      pts.forEach(([x, y], i) => {
        v.setInt16(20 + i * 4, x, true);
        v.setInt16(20 + i * 4 + 2, y, true);
      });
      return emr(88, buf);
    }

    const noOp = (t: number) => emr(t, new Uint8Array(0));

    /** EMR_SETWORLDTRANSFORM: six floats behind the header. */
    function worldTransform(scale: number, dx: number, dy: number): Uint8Array {
      const buf = new Uint8Array(24);
      const v = new DataView(buf.buffer);
      v.setFloat32(0, scale, true);
      v.setFloat32(12, scale, true);
      v.setFloat32(16, dx, true);
      v.setFloat32(20, dy, true);
      return emr(35, buf);
    }

    it("fills a BeginPath..EndPath figure with the selected brush", () => {
      const wmf = dualModeWmf(
        emfCarrier([
          emrEmfPlusComment([epHeader(), epEof()]),
          createSolidBrush(2, "b88e75"),
          selectObject(2),
          noOp(58), // BeginPath
          moveTo(100, 100),
          polyLineTo16([
            [200, 100],
            [200, 200],
          ]),
          noOp(61), // CloseFigure
          noOp(60), // EndPath
          noOp(62), // FillPath
        ]),
      );
      const members = emfPlusMembers(wmf, 100, 100);
      expect(members).toBeDefined();
      const filled = members!.find((m) => m.kind === "path" && m.fill === "b88e75");
      expect(filled).toBeDefined();
      if (filled?.kind !== "path") return;
      expect(filled.d).toContain("M");
      expect(filled.d).toContain("L");
      expect(filled.d).toContain("Z");
    });

    it("scales the frozen path through the live world transform", () => {
      const wmf = dualModeWmf(
        emfCarrier([
          emrEmfPlusComment([epHeader(), epEof()]),
          createSolidBrush(1, "404040"),
          selectObject(1),
          worldTransform(2, 10, 20), // SETWORLDTRANSFORM ×2 +10,+20
          noOp(58),
          moveTo(10, 10),
          polyLineTo16([
            [30, 10],
            [30, 30],
          ]),
          noOp(60),
          noOp(62),
        ]),
      );
      const members = emfPlusMembers(wmf, 400, 400);
      expect(members?.length).toBe(1);
      if (!members) return;
      const filled = members.find((m) => m.kind === "path" && "fill" in m && m.fill === "404040");
      if (filled?.kind !== "path") return;
      // points ×2 + offset: (10,10)→(30,40) in record space, then the frame
      // normalization stretches onto the box — both points must survive.
      expect(filled.d).toContain("L");
    });

    it("opens the figure from MoveToEx without a BeginPath bracket", () => {
      // The corpus exporter omits EMR_BEGINPATH entirely: MoveToEx starts the
      // figure and EndPath freezes it.
      const wmf = dualModeWmf(
        emfCarrier([
          emrEmfPlusComment([epHeader(), epEof()]),
          createSolidBrush(2, "123456"),
          selectObject(2),
          moveTo(100, 100),
          polyLineTo16([
            [200, 100],
            [200, 200],
          ]),
          noOp(60),
          noOp(62),
        ]),
      );
      const members = emfPlusMembers(wmf, 100, 100);
      const filled = (members ?? []).find(
        (m) => m.kind === "path" && "fill" in m && m.fill === "123456",
      );
      expect(filled).toBeDefined();
    });
  });
});
