import type { JSONContent } from "@tiptap/core";

import { Paragraph as BaseParagraph } from "./tiptap";
import {
  renderParagraphStyles,
  alignmentFromElement,
  indentFromElement,
  spacingFromElement,
  bordersFromElement,
  shadingFromElement,
} from "./utils";

/**
 * Paragraph extension with nested office-open attrs.
 *
 * Attrs mirror ParagraphPropertiesOptionsBase (alignment/indent/spacing/border/
 * shading/frame as nested objects + scalar OOXML properties). DOCX round-trip is
 * near-identity: renderDocx/parseDocx pass attrs through; CSS conversion happens
 * only in renderHTML via utils mappers.
 */

// ── DOCX serialization (near-identity: attrs mirror ParagraphPropertiesOptionsBase) ──

export function renderDocx(node: JSONContent): Record<string, unknown> {
  const attrs = (node.attrs ?? {}) as Record<string, unknown>;
  const opts: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (value !== null && value !== undefined) opts[key] = value;
  }
  return opts;
}

/** Structural/semantic keys expressed elsewhere (heading ext, list handling, run children). */
const SKIP_KEYS = new Set([
  "heading",
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
  for (const [key, value] of Object.entries(resolved)) {
    if (SKIP_KEYS.has(key)) continue;
    attrs[key] = value ?? null;
  }
  return attrs;
}

// ── Attr that stores an office-open native value (not parsed from HTML) ──

const attrNative = () => ({ default: null, parseHTML: () => null, rendered: false });

// ── Extension ──

export const Paragraph = BaseParagraph.extend({
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
    const attrs = { ...HTMLAttributes };
    if (styles.length > 0) attrs.style = styles.join(";");
    return ["p", attrs, 0] as const;
  },

  renderDocx,
  parseDocx,
});
