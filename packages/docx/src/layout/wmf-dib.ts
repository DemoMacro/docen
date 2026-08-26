// Placeable-WMF fallback: browsers cannot rasterize GDI metafiles, so a
// metafile picture renders as an empty frame. Most business-document WMFs
// (Office's ".emf" exports included — the magic here is the WMF placeable
// header) carry their main visual as one big DIB blit record; re-framing
// that DIB as a BMP data URL recovers the dominant imagery without a
// metafile player. Vector overlays (text, strokes) stay lost — a partial
// picture beats a gray box.

/** Extract the largest blit-record DIB as a BMP data URL, or undefined when
 *  the bytes are not a placeable WMF or hold no usable DIB. */
export function wmfDibFallback(bytes: Uint8Array): string | undefined {
  if (bytes.length < 22 + 18 + 8) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  // Placeable header magic 0x9AC6CDD7 (the 22-byte aldus header before the
  // standard WMF stream).
  if (view.getUint32(0, true) !== 0x9ac6cdd7) return undefined;
  let off = 22 + 18; // placeable + standard WMF header
  let best: { start: number; length: number } | undefined;
  while (off + 6 <= bytes.length) {
    const sizeWords = view.getUint32(off, true);
    const fn = view.getUint16(off + 4, true);
    if (sizeWords < 3) break;
    const recordEnd = off + sizeWords * 2;
    if (recordEnd > bytes.length) break;
    // The DIB-carrying blts: DIBBITBLT (fn 0x41), STRETCHDIB (0x40),
    // DIBSTRETCHBLT (0x43) — the full record word varies with the parameter
    // size, so match the low byte.
    const dib = bltDibAt(view, off, recordEnd, fn & 0xff);
    if (dib && (!best || dib.length > best.length)) best = dib;
    off = recordEnd;
  }
  if (!best) return undefined;
  return bmpDataUrl(bytes, best.start, best.length);
}

/** The DIB payload of one blt record: the BITMAPINFOHEADER signature (40)
 *  sits within the record's leading parameter words; the payload runs to the
 *  record end. Validated down to the forms a BMP consumer can take raw:
 *  uncompressed 8/24/32bpp with a real extent. DIBs past 12MB are skipped —
 *  a memory guard for degenerate metafiles, well above any real photo blit. */
function bltDibAt(
  view: DataView,
  recordStart: number,
  recordEnd: number,
  fnLow: number,
): { start: number; length: number } | undefined {
  if (fnLow !== 0x41 && fnLow !== 0x40 && fnLow !== 0x43) return undefined;
  if (recordEnd - recordStart < 26 + 40 + 8) return undefined;
  if (recordEnd - recordStart > 12 * 1024 * 1024) return undefined;
  for (let probe = recordStart + 6; probe < recordStart + 60; probe += 2) {
    if (view.getUint32(probe, true) !== 40) continue;
    const w = view.getInt32(probe + 4, true);
    const h = view.getInt32(probe + 8, true);
    const planes = view.getUint16(probe + 12, true);
    const bpp = view.getUint16(probe + 14, true);
    const compression = view.getUint32(probe + 16, true);
    if (w > 8 && Math.abs(h) > 8 && planes === 1 && compression === 0) {
      if (bpp === 8 || bpp === 24 || bpp === 32) {
        return { start: probe, length: recordEnd - probe };
      }
    }
    return undefined;
  }
  return undefined;
}

/** A DIB (BITMAPINFOHEADER onward) becomes a viewable BMP by prepending the
 *  14-byte BITMAPFILEHEADER; the pixel data and palette pass through raw. The
 *  pixel offset must skip the palette (biClrUsed entries, or 2^bpp, for
 *  palettized depths) — decoders honor it literally. */
function bmpDataUrl(bytes: Uint8Array, dibStart: number, dibLength: number): string {
  const dibView = new DataView(bytes.buffer, bytes.byteOffset + dibStart, dibLength);
  const bpp = dibView.getUint16(14, true);
  const clrUsed = dibView.getUint32(32, true);
  const paletteBytes = bpp <= 8 ? (clrUsed || 1 << bpp) * 4 : 0;
  const bmp = new Uint8Array(14 + dibLength);
  bmp[0] = 0x42; // "BM"
  bmp[1] = 0x4d;
  const view = new DataView(bmp.buffer);
  view.setUint32(2, bmp.length, true); // file size
  view.setUint32(10, 14 + 40 + paletteBytes, true); // pixel offset
  bmp.set(bytes.subarray(dibStart, dibStart + dibLength), 14);
  let bin = "";
  const CHUNK = 0x8000;
  for (let i = 0; i < bmp.length; i += CHUNK) {
    bin += String.fromCharCode(...bmp.subarray(i, i + CHUNK));
  }
  return `data:image/bmp;base64,${btoa(bin)}`;
}
