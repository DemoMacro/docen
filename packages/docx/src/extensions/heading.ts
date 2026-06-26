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

/** Runtime-only attrs the TableOfContents extension injects on each heading
 *  (`id` / `data-toc-id`). They are regenerated on every load, so never persist
 *  them to DOCX — skip them in renderDocx. */
const TOC_RUNTIME_KEYS = new Set(["id", "data-toc-id"]);

// ── DOCX serialization (near-identity: attrs mirror ParagraphPropertiesOptionsBase) ──

export function renderDocx(node: JSONContent): Record<string, unknown> {
  const attrs = (node.attrs ?? {}) as Record<string, unknown>;
  const opts: Record<string, unknown> = {};

  // Emit the original pStyle val verbatim. The stored styleId IS the source
  // pStyle ("Heading3" in some docs, or the numeric id "3" common in WPS /
  // Chinese Word whose styles.xml keeps styleId="3"). Deriving a HeadingLevel
  // literal from `level` would write pStyle="Heading3" while styles.xml still
  // carries styleId="3" — a mismatch Word silently drops, losing the heading
  // level. Fall back to `level` only when no styleId is present.
  const level = attrs.level as number | undefined;
  const styleId = attrs.styleId as string | undefined;
  if (styleId) opts.heading = styleId;
  else if (level) opts.heading = HEADING_COMPILE_MAP[level] ?? "Heading1";

  // Pass remaining attrs through verbatim (skip nulls + mapped fields).
  for (const [key, value] of Object.entries(attrs)) {
    if (key === "level" || key === "styleId" || TOC_RUNTIME_KEYS.has(key)) continue;
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

  // Reverse-map OOXML `heading` → Tiptap `level`. office-open lifts a
  // HeadingLevel pStyle ("Heading1".."Heading6"/"Title") into `heading`; that
  // literal IS the pStyle val, so also carry it as styleId for CSS.
  if (resolved.heading) {
    const level = HEADING_PARSE_MAP[resolved.heading as string];
    if (level) attrs.level = level;
    attrs.styleId = resolved.heading;
  }
  // An explicit OOXML `style` (non-HeadingLevel pStyle) overrides styleId.
  if (resolved.style) attrs.styleId = resolved.style;

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

      // pStyle reference (e.g. "Heading1") — same as Paragraph. renderHTML emits
      // class="docx-style-{styleId}" for the injected document CSS.
      styleId: {
        default: null,
        parseHTML: (el: HTMLElement) => {
          const m = (el.getAttribute("class") || "").match(/(?:^|\s)docx-style-(\S+)/);
          return m ? m[1] : null;
        },
      },

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
    node: { attrs: Record<string, unknown> } & {
      forEach?: (cb: (child: { isText?: boolean; type?: { name?: string } }) => void) => void;
    };
    HTMLAttributes: Record<string, unknown>;
  }) {
    // An empty heading (¶ glyph only) renders spacing.line at the natural
    // metric (no grid pitch), matching Word — same as Paragraph. Mirrors
    // measure.ts isEmptyTextblock. See renderParagraphStyles `empty`.
    let hasContent = false;
    node.forEach?.((child) => {
      if (child.isText || child.type?.name === "hardBreak" || child.type?.name === "image")
        hasContent = true;
    });
    const styles = renderParagraphStyles(node.attrs, { empty: !hasContent });
    const level = (node.attrs?.level as number) ?? 1;
    const attrs = { ...HTMLAttributes };
    const styleId = node.attrs.styleId as string | null;
    if (styleId) attrs.class = `docx-style-${styleId}`;
    if (styles.length > 0) attrs.style = styles.join(";");
    return [`h${level}`, attrs, 0] as const;
  },

  renderDocx,
  parseDocx,
});
