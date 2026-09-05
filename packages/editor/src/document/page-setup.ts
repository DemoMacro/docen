// Page-setup presets and section-properties merging: MS Office paper sizes
// and margin presets, the CSS→twips margin parse, and the deep merge that
// patches a section without losing sides the patch omits.

import { convertMillimetersToTwip, type SectionPropertiesOptions } from "@docen/docx";

/** MS Office standard paper sizes (mm, portrait width × height). Page-setup
 *  presets resolve to raw mm here; <docen-document-area> takes only raw page-width /
 *  page-height, so presets stay in this document layer, not the UI component. */
export const PAPER_SIZES: Readonly<Record<string, readonly [number, number]>> = {
  letter: [215.9, 279.4],
  legal: [215.9, 355.6],
  statement: [139.7, 215.9],
  executive: [184.15, 266.7],
  tabloid: [279.4, 431.8],
  a3: [297, 420],
  a4: [210, 297],
  a5: [148, 210],
  a6: [105, 148],
  b5: [182, 257],
};

/** MS Office margin presets (mm). `normal` matches the engine default
 *  (@office-open/docx sectionMarginDefaults: top/bottom 25.4mm, left/right
 *  31.75mm = MS Office zh-CN "Normal"). */
export const MARGINS: Readonly<Record<string, string>> = {
  normal: "25.4mm 31.75mm",
  narrow: "12.7mm",
  moderate: "25.4mm 19.05mm",
  wide: "25.4mm 50.8mm",
};

/** Parse a CSS padding list (mm, 1–4 values) into OOXML page margins (twips),
 *  via the engine's convertMillimetersToTwip (mm → twips). */
export function marginTwipsFromCss(css: string): {
  top: number;
  right: number;
  bottom: number;
  left: number;
} {
  const mm = css.split(/\s+/).map((s) => parseFloat(s));
  const [t, r, b, l] =
    mm.length === 1
      ? [mm[0], mm[0], mm[0], mm[0]]
      : mm.length === 2
        ? [mm[0], mm[1], mm[0], mm[1]]
        : [mm[0], mm[1], mm[2] ?? mm[1], mm[3] ?? mm[1]];
  return {
    top: convertMillimetersToTwip(t),
    right: convertMillimetersToTwip(r),
    bottom: convertMillimetersToTwip(b),
    left: convertMillimetersToTwip(l),
  };
}

/** Deep-merge a sectionProperties patch (page.size / page.margin) into a base,
 *  preserving sides/dims the patch omits — so e.g. changing only the margins
 *  keeps the page size. `pageBorders` replaces whole (a present-but-undefined
 *  key clears it — Word's "none" preset; an absent key keeps the base).
 *  Reuses the engine's SectionPropertiesOptions type. */
export function mergeSectionProperties(
  base: SectionPropertiesOptions | null | undefined,
  patch: SectionPropertiesOptions,
): SectionPropertiesOptions {
  const mergeGroup = <T extends object>(
    b: T | false | undefined,
    p: T | false | undefined,
  ): T | false | undefined =>
    p === undefined ? b : p === false || b === undefined || b === false ? p : { ...b, ...p };
  return {
    ...base,
    ...("pageBorders" in patch ? { pageBorders: patch.pageBorders } : {}),
    // Present-but-undefined clears verticalAlign (the dialog's "top" — Word
    // drops w:vAlign); an absent key keeps the base.
    ...("verticalAlign" in patch ? { verticalAlign: patch.verticalAlign } : {}),
    pageSize: mergeGroup(base?.pageSize, patch.pageSize),
    pageMargin: mergeGroup(base?.pageMargin, patch.pageMargin),
  };
}
