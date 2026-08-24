// Deterministic synthetic measurement for Node-side specs. pretext measures
// through a canvas 2d context; in Node there is none, so specs install this
// fake OffscreenCanvas whose measureText derives widths from the font string
// alone: CJK graphemes 1em, ordinary graphemes 0.5em, spaces 0.25em. Font
// PRECISION tests belong in browser mode (real system fonts); these specs
// assert the packer's STRUCTURE (breaks, positions, heights), which only
// needs deterministic widths.

import { isCjkCodePoint } from "../src/font";

export const FAKE_NORMAL_RATIO = 1.2;

/** The fake's advance for one grapheme at `em` px. */
export function fakeAdvanceOf(ch: string, em: number): number {
  if (ch === " " || ch === "\t") return em / 4;
  if (isCjkCodePoint(ch)) return em;
  return em / 2;
}

/** Install the fake OffscreenCanvas (idempotent). Returns the px-per-em the
 *  fake reads out of a CSS font shorthand. */
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
  // eslint-disable-next-line @typescript-eslint/no-explicit-any -- the minimal duck-typed surface pretext touches
  (globalThis as any).OffscreenCanvas = class {
    getContext(): unknown {
      return ctx();
    }
  };
}

/** A FontMetrics stub matching the fake canvas's synthetic world. */
export const fakeFontMetrics = { normalRatio: (): number => FAKE_NORMAL_RATIO };

/** Total width the fake assigns a string at `em` px. */
export function fakeWidthOf(text: string, em: number): number {
  let w = 0;
  for (const ch of text) w += fakeAdvanceOf(ch, em);
  return w;
}
