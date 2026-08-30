import { describe, expect, it } from "vitest";

import { wmfDibFallback } from "./dib";
import { decodedBmp, dib, wmfWithRecords } from "./test-util";

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
