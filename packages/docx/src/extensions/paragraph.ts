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

// Editor-only attrs that mark a paragraph as a section's last paragraph (its pPr
// holds the OOXML sectPr). DocxManager peels them off to close a section in
// compile; they must NOT leak into ParagraphOptions.
const SECTION_ATTR_KEYS = new Set(["sectionProperties", "sectionHeaders", "sectionFooters"]);

export function renderDocx(node: JSONContent): Record<string, unknown> {
  const attrs = (node.attrs ?? {}) as Record<string, unknown>;
  const opts: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (value === null || value === undefined) continue;
    if (SECTION_ATTR_KEYS.has(key)) continue;
    // styleId (attr) → OOXML `style` (the paragraph's pStyle reference).
    if (key === "styleId") {
      opts.style = value;
      continue;
    }
    opts[key] = value;
  }
  return opts;
}

/**
 * Structural/semantic keys handled elsewhere (heading ext, list handling, run/text
 * children). NOTE: `run` is intentionally NOT skipped — ParagraphOptions.run (the
 * paragraph's default run properties: font/size/color) is kept as an attr for
 * lossless round-trip (e.g. header/footer paragraphs whose styling lives there).
 */
const SKIP_KEYS = new Set([
  "heading",
  "style",
  "bullet",
  "numbering",
  "children",
  "text",
  "thematicBreak",
]);

export function parseDocx(opts: Record<string, unknown>): Record<string, unknown> {
  const resolved = typeof opts === "string" ? { text: opts } : opts;
  const attrs: Record<string, unknown> = {};
  // OOXML `style` (the paragraph's pStyle reference, e.g. "Heading1") → styleId,
  // carried as an attr so the named style's CSS applies via class="docx-style-{id}".
  if (resolved.style) attrs.styleId = resolved.style;
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

      // pStyle reference (e.g. "Heading1", "Title", "Normal") — the named
      // paragraph style. renderHTML emits class="docx-style-{styleId}" so the
      // injected document CSS (generated from styles.xml) applies. Round-trips
      // via OOXML `style`.
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
      // Paragraph-mark run properties (pPr/rPr): per OOXML (ECMA-376) these
      // format ONLY the ¶ glyph — never applied to the paragraph's runs. Carried
      // verbatim for lossless DOCX round-trip (renderDocx emits opts.run); NOT
      // rendered to CSS. Run font/size/color come from the run's own marks or
      // the named-style / docDefaults CSS (stylesToCss).
      run: attrNative(),

      // Section properties carried on a section's LAST paragraph (OOXML sectPr
      // lives in that paragraph's pPr). DocxManager uses these to split/merge
      // sections in compile/resolve; they never render and never reach
      // ParagraphOptions (renderDocx skips them).
      sectionProperties: attrNative(),
      sectionHeaders: attrNative(),
      sectionFooters: attrNative(),

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
    const styleId = node.attrs.styleId as string | null;
    // class="docx-style-{id}" applies the named style's CSS. A pStyle-less
    // paragraph renders as the document's default paragraph style (OOXML:
    // no pStyle = default style), so mark it `docx-default` — stylesToCss emits
    // the default style's rule under that selector too.
    attrs.class = styleId ? `docx-style-${styleId}` : "docx-default";
    if (styles.length > 0) attrs.style = styles.join(";");
    return ["p", attrs, 0] as const;
  },

  renderDocx,
  parseDocx,
});
