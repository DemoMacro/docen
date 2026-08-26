import { describe, expect, it } from "vitest";

import { wmfDibFallback } from "./wmf-dib";

/** Uncompressed top-down DIB (BITMAPINFOHEADER + zeroed pixels). */
function dib(w: number, h: number, bpp: number): Uint8Array {
  const row = Math.ceil((w * bpp) / 8 / 4) * 4; // rows are 4-byte aligned
  const buf = new Uint8Array(40 + row * h);
  const v = new DataView(buf.buffer);
  v.setUint32(0, 40, true);
  v.setInt32(4, w, true);
  v.setInt32(8, h, true);
  v.setUint16(12, 1, true);
  v.setUint16(14, bpp, true);
  return buf;
}

/** Placeable WMF: aldus magic + standard header + the given records. */
function wmfWithRecords(records: { fn: number; params: Uint8Array }[]): Uint8Array {
  const placeable = new Uint8Array(22);
  new DataView(placeable.buffer).setUint32(0, 0x9ac6cdd7, true);
  const header = new Uint8Array(18);
  new DataView(header.buffer).setUint16(0, 1, true); // memory metafile
  new DataView(header.buffer).setUint16(2, 9, true); // header size in words
  const body = records.map((rec) => {
    const buf = new Uint8Array(6 + rec.params.length);
    new DataView(buf.buffer).setUint32(0, buf.length / 2, true); // size in words
    new DataView(buf.buffer).setUint16(4, rec.fn, true);
    buf.set(rec.params, 6);
    return buf;
  });
  const total = [placeable, header, ...body].reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const part of [placeable, header, ...body]) {
    out.set(part, off);
    off += part.length;
  }
  return out;
}

function decodedBmp(src: string | undefined): Uint8Array {
  expect(src).toBeDefined();
  const bin = atob(src!.slice("data:image/bmp;base64,".length));
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

describe("wmfDibFallback", () => {
  it("reframes a stretch-DIB blit record as a BMP data URL", () => {
    const src = wmfDibFallback(wmfWithRecords([{ fn: 0x0f43, params: dib(9, 9, 24) }]));
    const bmp = decodedBmp(src);
    expect(String.fromCharCode(bmp[0], bmp[1])).toBe("BM");
    expect(new DataView(bmp.buffer).getUint32(2, true)).toBe(bmp.length);
    expect(new DataView(bmp.buffer).getUint32(10, true)).toBe(14 + 40);
    // the DIB passes through raw right after the file header
    expect(new DataView(bmp.buffer).getUint32(14, true)).toBe(40);
    expect(new DataView(bmp.buffer).getInt32(18, true)).toBe(9);
  });

  it("finds the signature behind leading record params", () => {
    const lead = new Uint8Array(8); // 4 words of blit params before the DIB
    const params = new Uint8Array(lead.length + dib(12, 10, 24).length);
    params.set(dib(12, 10, 24), lead.length);
    expect(wmfDibFallback(wmfWithRecords([{ fn: 0x0b41, params }]))).toBeDefined();
  });

  it("keeps the largest DIB across records", () => {
    const src = wmfDibFallback(
      wmfWithRecords([
        { fn: 0x0f43, params: dib(9, 9, 24) },
        { fn: 0x0f43, params: dib(20, 10, 24) },
      ]),
    );
    const bmp = decodedBmp(src);
    expect(bmp.length).toBe(14 + dib(20, 10, 24).length);
  });

  it("rejects mask-form bpp, non-placeable bytes, and DIB-less streams", () => {
    // 1bpp is the blt AND-mask form, not a viewable picture
    expect(wmfDibFallback(wmfWithRecords([{ fn: 0x0b41, params: dib(9, 9, 1) }]))).toBeUndefined();
    const notPlaceable = wmfWithRecords([{ fn: 0x0f43, params: dib(9, 9, 24) }]);
    notPlaceable[0] = 0;
    expect(wmfDibFallback(notPlaceable)).toBeUndefined();
    expect(
      wmfDibFallback(wmfWithRecords([{ fn: 0x0213, params: new Uint8Array(30) }])),
    ).toBeUndefined();
  });

  it("offsets pixels past the 8bpp palette", () => {
    const body = dib(16, 16, 8);
    const dibBytes = new Uint8Array(body.length + 256 * 4); // header + palette
    dibBytes.set(body, 0);
    const src = wmfDibFallback(wmfWithRecords([{ fn: 0x0f43, params: dibBytes }]));
    const bmp = decodedBmp(src);
    expect(new DataView(bmp.buffer).getUint32(10, true)).toBe(14 + 40 + 1024);
  });
});
