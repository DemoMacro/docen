// Deterministic synthetic measurement for Node-side specs. Measurement runs
// through a canvas 2d context; in Node there is none, so specs install this
// fake OffscreenCanvas whose measureText derives widths from the font string
// alone: CJK graphemes 1em rounded UP to the next integer px (mirroring the
// canvas rounding the CJK correction exists for), ordinary graphemes 0.5em,
// spaces 0.25em.

/** The fake's advance for one grapheme at `em` px (canvas-style CJK rounding). */
export function fakeAdvanceOf(ch: string, em: number): number {
  if (ch === " " || ch === "\t") return em / 4;
  if (isCjkCodePoint(ch)) return Math.ceil(em);
  return em / 2;
}

function isCjkCodePoint(ch: string): boolean {
  const c = ch.codePointAt(0) ?? 0;
  return (
    (c >= 0x2e80 && c <= 0x9fff) || (c >= 0x3400 && c <= 0x4dbf) || (c >= 0xf900 && c <= 0xfaff)
  );
}

/** Install the fake OffscreenCanvas (idempotent). */
export function installFakeCanvas(): void {
  if (typeof OffscreenCanvas !== "undefined") return;
  const fontShorthandEm = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? Number(m[1]) : 16;
  };
  const ctx = () => ({
    _font: "16px serif",
    set font(v: string) {
      this._font = v;
    },
    get font(): string {
      return this._font;
    },
    measureText(s: string): { width: number } {
      const em = fontShorthandEm(this._font);
      let w = 0;
      for (const ch of s) w += fakeAdvanceOf(ch, em);
      return { width: w };
    },
  });
  // The minimal duck-typed surface the measurement module touches.
  (globalThis as unknown as Record<string, unknown>).OffscreenCanvas = class {
    getContext(): unknown {
      return ctx();
    }
  };
}

/** Total width the fake assigns a string at `em` px. */
export function fakeWidthOf(text: string, em: number): number {
  let w = 0;
  for (const ch of text) w += fakeAdvanceOf(ch, em);
  return w;
}
