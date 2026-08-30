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

// ── dual-mode (EMF+-carrying) fixtures ──
// A placeable WMF whose META_ESCAPE/MFCOMMENT chunks embed a complete EMF of
// EMR_GDICOMMENT "EMF+" payloads; the player must reassemble both layers.

/** One EmfPlus record: type/flags header plus raw body. */
export function epRecord(type: number, flags: number, body: Uint8Array): Uint8Array {
  const buf = new Uint8Array(8 + body.length);
  const v = new DataView(buf.buffer);
  v.setUint16(0, type, true);
  v.setUint16(2, flags, true);
  v.setUint32(4, buf.length, true);
  buf.set(body, 8);
  return buf;
}

/** An EMR wrapping a GDI comment whose private data is an "EMF+" stream. */
export function emrEmfPlusComment(plusRecords: Uint8Array[]): Uint8Array {
  const plusLen = plusRecords.reduce((n, r) => n + r.length, 0);
  const buf = new Uint8Array(16 + plusLen);
  const v = new DataView(buf.buffer);
  v.setUint32(0, 70, true); // EMR_GDICOMMENT
  v.setUint32(4, buf.length, true);
  v.setUint32(8, plusLen + 4, true); // DataSize
  v.setUint32(12, 0x2b464d45, true); // "EMF+" signature
  let off = 16;
  for (const r of plusRecords) {
    buf.set(r, off);
    off += r.length;
  }
  return buf;
}

/** Minimal EMF carrier: header-sized dummy record then the given raw EMR(s). */
export function emfCarrier(comments: Uint8Array[]): Uint8Array {
  const head = new Uint8Array(100);
  new DataView(head.buffer).setUint32(0, 1, true); // EMR_HEADER
  new DataView(head.buffer).setUint32(4, head.length, true);
  const eof = new Uint8Array(20);
  new DataView(eof.buffer).setUint32(0, 14, true); // EMR_EOF
  new DataView(eof.buffer).setUint32(4, eof.length, true);
  const parts = [head, ...comments, eof];
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}

/** Wrap one EMF into WMFC escape chunks over a bare placeable WMF. Chunk
 *  headers carry the reassembly fields the real exporter writes; the player
 *  reads only the magic and the declared byte count. */
export function dualModeWmf(emf: Uint8Array, chunkSize = 8192): Uint8Array {
  const chunks: Uint8Array[] = [];
  for (let o = 0; o < emf.length || o === 0; o += chunkSize)
    chunks.push(emf.subarray(o, Math.min(o + chunkSize, emf.length)));
  // escape data: magic(4) + chunk header(34) + slice
  const records = chunks.map((c) => {
    // escape data: magic + 30-byte chunk header (34 total), then the slice
    const data = new Uint8Array(34 + c.length);
    new DataView(data.buffer).setUint32(0, 0x43464d57, true); // "WMFC"
    const v = new DataView(data.buffer);
    v.setUint16(18, chunks.length, true); // [+18] total chunks
    v.setUint16(22, c.length, true); // [+22] chunk payload length
    data.set(c, 34);
    // WMF records are word-sized; an odd tail is padded by the zero fill.
    const rec = new Uint8Array((10 + data.length + 1) & ~1);
    const rv = new DataView(rec.buffer);
    rv.setUint32(0, rec.length / 2, true); // sizeWords
    rv.setUint16(4, 0x0626, true); // META_ESCAPE
    rv.setUint16(6, 0x0626, true); // MFCOMMENT
    rv.setUint16(8, data.length, true); // byte count
    rec.set(data, 10);
    return rec;
  });
  const placeable = new Uint8Array(40); // placeable block + WMF header
  const pv = new DataView(placeable.buffer);
  pv.setUint32(0, 0x9ac6cdd7, true);
  pv.setInt16(6, 0, true);
  pv.setInt16(8, 0, true);
  pv.setInt16(10, 605, true);
  pv.setInt16(12, 240, true);
  pv.setUint16(14, 1440, true);
  const total = [placeable, ...records].reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of [placeable, ...records]) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}

/** Dual-mode WMF carrying the given EmfPlus records inside one comment. */
export function emfPlusWmf(plusRecords: Uint8Array[], chunkSize?: number): Uint8Array {
  return dualModeWmf(emfCarrier([emrEmfPlusComment(plusRecords)]), chunkSize);
}
