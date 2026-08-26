// Shared WMF byte-level builders for the player (wmf.spec.ts) and DIB
// fallback (wmf-dib.spec.ts) specs.

import { expect } from "vitest";

/** Uncompressed top-down DIB (BITMAPINFOHEADER + zeroed pixels). */
export function dib(w: number, h: number, bpp: number): Uint8Array {
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

/** Little-endian int16 params for one record. */
export function words(...ws: number[]): Uint8Array {
  const buf = new Uint8Array(ws.length * 2);
  const v = new DataView(buf.buffer);
  ws.forEach((w, i) => v.setInt16(i * 2, w, true));
  return buf;
}

/** Placeable WMF: aldus magic, a real bbox, standard header, then the given
 *  records. The bbox defaults to (0,0)-(w,h); an offset origin exercises
 *  normalization against org. */
export function wmfWithRecords(
  records: { fn: number; params: Uint8Array }[],
  w = 200,
  h = 100,
  left = 0,
  top = 0,
): Uint8Array {
  const placeable = new Uint8Array(22);
  const pv = new DataView(placeable.buffer);
  pv.setUint32(0, 0x9ac6cdd7, true);
  pv.setInt16(6, left, true);
  pv.setInt16(8, top, true);
  pv.setInt16(10, left + w, true);
  pv.setInt16(12, top + h, true);
  pv.setUint16(14, 1440, true); // inch
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

/** LOGFONT16 (50 bytes) for CreateFontIndirect: height/weight/italic +
 *  faceName GBK bytes at @18. */
export function logfont(
  o: { height?: number; weight?: number; italic?: boolean; face?: string } = {},
): Uint8Array {
  const buf = new Uint8Array(50);
  const v = new DataView(buf.buffer);
  v.setInt16(0, o.height ?? -16, true);
  v.setUint16(8, o.weight ?? 400, true);
  if (o.italic) buf[10] = 1;
  buf.set(gbk(o.face ?? ""), 18);
  return buf;
}

// Node ships no GBK encoder — the codes for exactly the chars the specs use
// (verified against Python's gbk codec).
const GBK_CODES: Record<string, [number, number]> = {
  示: [0xca, 0xbe],
  例: [0xc0, 0xfd],
  文: [0xce, 0xc4],
  本: [0xb1, 0xbe],
  微: [0xce, 0xa2],
  软: [0xc8, 0xed],
  雅: [0xd1, 0xc5],
  黑: [0xba, 0xda],
};

/** Mixed ASCII + CJK string → GBK bytes (ASCII passes through 1:1). */
export function gbk(s: string): Uint8Array {
  const out: number[] = [];
  for (const ch of s) {
    const code = GBK_CODES[ch];
    if (code) out.push(code[0], code[1]);
    else if (ch.charCodeAt(0) < 0x80) out.push(ch.charCodeAt(0));
    else throw new Error(`no GBK code for ${ch}`);
  }
  return new Uint8Array(out);
}

export function decodedBmp(src: string | undefined): Uint8Array {
  expect(src).toBeDefined();
  const bin = atob(src!.slice("data:image/bmp;base64,".length));
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}
