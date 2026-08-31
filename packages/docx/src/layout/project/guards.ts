// Loose-shape guards shared by every projection domain: the Options unions
// are structurally loose at their edges (optional everything, per-side
// sub-objects), so fields are read through these narrow helpers instead of
// per-site casts — plus the universal-measure (number | UM string) parsing.

import type { LayoutTable } from "@docen/layout";
import type { ParagraphOptions } from "@office-open/docx";

// The paragraph leg of SectionChild is `string | ParagraphOptions` (shorthand
// or full options); null appears at runtime (empty paragraph legs from
// parse/compile), so the projection accepts it defensively.
export type BodyParagraph = string | ParagraphOptions | null;
export type LayoutCell = LayoutTable["rows"][number]["cells"][number];

// ── loose-shape guards ──

export type Rec = Record<string, unknown>;

/** Options unions are structurally loose at their edges (optional everything,
 *  per-side sub-objects); this guard narrows unknown/union picks to a record so
 *  the rest of the module reads fields without per-site casts. */
export function isRecord(v: unknown): v is Rec {
  return !!v && typeof v === "object";
}

export const num = (v: unknown): number | undefined => (typeof v === "number" ? v : undefined);

export const str = (v: unknown): string | undefined => (typeof v === "string" && v ? v : undefined);

/** The five predefined XML entities (numeric refs stay rare in field text). */
export function unescapeXml(v: string): string {
  return v
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

/** Estimated height of one placeholder box: three default body lines. */
export const PLACEHOLDER_PX = 3 * 16;

// ── universal-measure parsing (number = native unit, string = UM) ──

const UM_IN_TWIPS = { pt: 20, pc: 240, in: 1440, mm: 1440 / 25.4, cm: 1440 / 2.54, px: 15 };
const UM_RE = /^(-?[\d.]+)(pt|pc|in|mm|cm|px)$/;

/** A measure field to twips: number passes through (native), UM resolves. */
export function measureTwip(v: unknown): number | undefined {
  const n = num(v);
  if (n != null) return n;
  if (typeof v !== "string") return undefined;
  const m = UM_RE.exec(v);
  return m ? Number(m[1]) * UM_IN_TWIPS[m[2] as keyof typeof UM_IN_TWIPS] : undefined;
}

/** A measure field whose native unit is EMU (drawing extents): number passes,
 *  UM resolves (px at 96 dpi). */
export function measureEmu(v: unknown): number | undefined {
  const n = num(v);
  if (n != null) return n;
  if (typeof v !== "string") return undefined;
  const tw = measureTwip(v);
  return tw != null ? (tw / 1440) * 914400 : undefined;
}

/** A color field: bare hex string, or the round-trip object shape
 *  (`{value}` on outline/fill colors, `{val, themeColor}` on run colors — the
 *  parse emits both key spellings). Theme-only colors resolve later. */
export function colorOf(v: unknown): string | undefined {
  if (typeof v === "string") return v === "auto" ? undefined : v;
  if (isRecord(v)) return str(v.value) ?? str(v.val);
  return undefined;
}
