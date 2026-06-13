import type { JSONContent } from "@tiptap/core";

import { Heading as BaseHeading } from "./tiptap";
import {
  renderParagraphStyles,
  alignmentFromElement,
  indentFromElement,
  spacingFromElement,
  bordersFromElement,
  shadingFromElement,
} from "./utils";

/**
 * Heading extension with nested office-open attrs (mirrors Paragraph).
 *
 * Attrs mirror ParagraphPropertiesOptionsBase (alignment/indent/spacing/border/
 * shading/frame as nested objects + scalar OOXML properties) plus the inherited
 * Tiptap `level` (1-6, rendered: false). DOCX round-trip is near-identity:
 * renderDocx/parseDocx pass attrs through and only map `level` ↔ OOXML
 * `heading` (HeadingLevel literal). CSS conversion happens solely in
 * renderHTML via utils mappers.
 */

// HeadingLevel literals: "Heading1".."Heading6", "Title".
const HEADING_COMPILE_MAP: Record<number, string> = {
  1: "Heading1",
  2: "Heading2",
  3: "Heading3",
  4: "Heading4",
  5: "Heading5",
  6: "Heading6",
};

const HEADING_PARSE_MAP: Record<string, number> = {
  Heading1: 1,
  Heading2: 2,
  Heading3: 3,
  Heading4: 4,
  Heading5: 5,
  Heading6: 6,
  Title: 1,
};

// ── DOCX serialization (near-identity: attrs mirror ParagraphPropertiesOptionsBase) ──

export function renderDocx(node: JSONContent): Record<string, unknown> {
  const attrs = (node.attrs ?? {}) as Record<string, unknown>;
  const opts: Record<string, unknown> = {};

  // `level` is an attrs-only field; express it via OOXML `heading`.
  const level = attrs.level as number | undefined;
  if (level) opts.heading = HEADING_COMPILE_MAP[level] ?? "Heading1";

  // Pass remaining attrs through verbatim (skip nulls and the level field).
  for (const [key, value] of Object.entries(attrs)) {
    if (key === "level") continue;
    if (value !== null && value !== undefined) opts[key] = value;
  }
  return opts;
}

/** Structural/semantic keys expressed elsewhere (heading handled separately, list handling, run children). */
const SKIP_KEYS = new Set([
  "heading",
  "style",
  "bullet",
  "numbering",
  "run",
  "children",
  "text",
  "thematicBreak",
]);

export function parseDocx(opts: Record<string, unknown>): Record<string, unknown> {
  const resolved = typeof opts === "string" ? { text: opts } : opts;
  const attrs: Record<string, unknown> = {};

  // Reverse-map OOXML `heading` → Tiptap `level`.
  if (resolved.heading) {
    const level = HEADING_PARSE_MAP[resolved.heading as string];
    if (level) attrs.level = level;
  }

  // Pass remaining opts through (skip structural/semantic keys).
  for (const [key, value] of Object.entries(resolved)) {
    if (SKIP_KEYS.has(key)) continue;
    attrs[key] = value ?? null;
  }
  return attrs;
}

// ── Attr that stores an office-open native value (not parsed from HTML) ──

const attrNative = () => ({ default: null, parseHTML: () => null, rendered: false });

// ── Extension ──

export const Heading = BaseHeading.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      // Nested office-open objects (parsed from HTML where CSS exists)
      alignment: {
        default: null,
        rendered: false,
        parseHTML: (el: HTMLElement) => alignmentFromElement(el),
      },
      indent: {
        default: null,
        rendered: false,
        parseHTML: (el: HTMLElement) => indentFromElement(el),
      },
      spacing: {
        default: null,
        rendered: false,
        parseHTML: (el: HTMLElement) => spacingFromElement(el),
      },
      shading: {
        default: null,
        rendered: false,
        parseHTML: (el: HTMLElement) => shadingFromElement(el),
      },
      border: {
        default: null,
        rendered: false,
        parseHTML: (el: HTMLElement) => bordersFromElement(el),
      },
      frame: attrNative(),

      // Scalar OOXML paragraph properties (stored verbatim; no CSS equivalent)
      keepNext: attrNative(),
      keepLines: attrNative(),
      pageBreakBefore: attrNative(),
      widowControl: attrNative(),
      contextualSpacing: attrNative(),
      bidirectional: attrNative(),
      outlineLevel: attrNative(),
      textDirection: attrNative(),
      textAlignment: attrNative(),
      suppressLineNumbers: attrNative(),
      wordWrap: attrNative(),
      overflowPunctuation: attrNative(),
      autoSpaceEastAsianText: attrNative(),
      suppressOverlap: attrNative(),
      suppressAutoHyphens: attrNative(),
      adjustRightInd: attrNative(),
      snapToGrid: attrNative(),
      mirrorIndents: attrNative(),
      kinsoku: attrNative(),
      topLinePunct: attrNative(),
      autoSpaceDE: attrNative(),
      textboxTightWrap: attrNative(),
      rightTabStop: attrNative(),
      leftTabStop: attrNative(),
      divId: attrNative(),
      tabStops: attrNative(),
      cnfStyle: attrNative(),
    };
  },

  renderHTML({
    node,
    HTMLAttributes,
  }: {
    node: { attrs: Record<string, unknown> };
    HTMLAttributes: Record<string, unknown>;
  }) {
    const styles = renderParagraphStyles(node.attrs);
    const level = (node.attrs?.level as number) ?? 1;
    const attrs = { ...HTMLAttributes };
    if (styles.length > 0) attrs.style = styles.join(";");
    return [`h${level}`, attrs, 0] as const;
  },

  renderDocx,
  parseDocx,
});
