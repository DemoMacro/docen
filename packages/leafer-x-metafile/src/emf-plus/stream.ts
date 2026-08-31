const PLACEABLE_MAGIC = 0x9ac6cdd7;
const META_ESCAPE = 0x0626;
const WMFC_MAGIC = 0x43464d57; // "WMFC"
const WMFC_CHUNK_HEADER = 34;

/** Reassemble the nested EMF carried by the WMFC escape chunks, or undefined
 *  when the bytes are not such a WMF. */
export function embeddedEmfStream(bytes: Uint8Array): Uint8Array | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < 44 || view.getUint32(0, true) !== PLACEABLE_MAGIC) return undefined;
  const chunks: Uint8Array[] = [];
  let off = 40; // placeable + standard WMF header
  while (off + 6 <= bytes.length) {
    const sizeWords = view.getUint32(off, true);
    const fn = view.getUint16(off + 4, true);
    if (sizeWords < 3) break;
    const end = off + sizeWords * 2;
    if (end > bytes.length) break;
    if (fn === META_ESCAPE && view.getUint32(off + 10, true) === WMFC_MAGIC) {
      const cb = Math.min(view.getUint16(off + 8, true), end - off - 10);
      chunks.push(bytes.subarray(off + 10 + WMFC_CHUNK_HEADER, off + 10 + cb));
    }
    off = end;
  }
  if (!chunks.length) return undefined;
  const total = chunks.reduce((n, c) => n + c.length, 0);
  const emf = new Uint8Array(total);
  let o = 0;
  for (const c of chunks) {
    emf.set(c, o);
    o += c.length;
  }
  return total >= 108 ? emf : undefined;
}
