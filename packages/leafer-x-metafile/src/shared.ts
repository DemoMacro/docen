// Tiny GDI-era helpers the WMF player and the EMF+ replay share — record
// languages that encode the same device semantics the same way.

/** COLORREF (0x00BBGGRR) → hex RRGGBB. */
export function colorRefHex(colorref: number): string {
  const r = colorref & 0xff;
  const g = (colorref >> 8) & 0xff;
  const b = (colorref >> 16) & 0xff;
  return ((r << 16) | (g << 8) | b).toString(16).padStart(6, "0");
}

/** One character's advance as an em fraction — GDI's blanket CJK-vs-Latin
 *  rule (fullwidth = 1 em, everything else 0.55 em), the fallback when a
 *  text record ships no per-glyph Dx run. */
export function gdiAdvanceEm(ch: string): number {
  return ch.charCodeAt(0) > 0xff ? 1 : 0.55;
}
