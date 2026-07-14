import type { StylesOptions } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";
import { Heading as BaseHeading } from "@tiptap/extension-heading";
import type { Node } from "@tiptap/pm/model";

import { buildTextBlock, indexParagraphStyles } from "../converters/styles";
import type { ParseParagraphRule } from "./types";
import { docxParagraphAttrs, renderTextBlock, SECTION_ATTR_KEYS } from "./utils";

/**
 * Heading extension — a paragraph with an outline level.
 *
 * In OOXML a heading IS a paragraph: a <w:p> whose pPr carries pStyle="Heading1"
 * (or an outlineLvl), and a sectPr (when present) lives in that same <w:p>'s pPr.
 * This node therefore shares Paragraph's office-open attrs via docxParagraphAttrs
 * and adds only Tiptap's inherited `level` (1-9) plus the level↔HeadingLevel
 * mapping.
 *
 * DOCX round-trip is near-identity: renderDocx/parseDocx pass attrs through and
 * only map `level` ↔ OOXML `heading` (HeadingLevel literal). CSS conversion
 * happens solely in renderHTML via utils mappers.
 */

// HeadingLevel literals: "Heading1".."Heading9", "Title".
const HEADING_COMPILE_MAP: Record<number, string> = {
  1: "Heading1",
  2: "Heading2",
  3: "Heading3",
  4: "Heading4",
  5: "Heading5",
  6: "Heading6",
  7: "Heading7",
  8: "Heading8",
  9: "Heading9",
};

const HEADING_PARSE_MAP: Record<string, number> = {
  Heading1: 1,
  Heading2: 2,
  Heading3: 3,
  Heading4: 4,
  Heading5: 5,
  Heading6: 6,
  Heading7: 7,
  Heading8: 8,
  Heading9: 9,
  Title: 1,
};

/** Heading level (1-9) from a localized style NAME: "heading 1"/"标题 1" → 1,
 *  "title" → 1. office-open's built-in names are English ("heading 1"), but
 *  zh-CN Word labels the same styles "标题 1"; both map to the same level. */
function headingLevelFromName(name: string | undefined): number | undefined {
  if (!name) return undefined;
  const m = /^heading\s+(\d)$/i.exec(name) ?? /^标题\s*(\d)$/.exec(name);
  if (m) {
    const lvl = Number(m[1]);
    if (lvl >= 1 && lvl <= 9) return lvl;
  }
  return /^title$/i.test(name) ? 1 : undefined;
}

/** Heading level (1-9) for a paragraph, or undefined when it isn't a heading.
 *  DOCX marks a heading several ways, checked in priority order:
 *  1. office-open lifts a HeadingLevel pStyle ("Heading1".."Title") into `heading`.
 *  2. An explicit `outlineLevel` (0-8 → 1-9) — Word's outline/TOC key off this
 *     even without a heading pStyle; the Heading1-9 styles carry outlineLvl 0-8.
 *  3. A pStyle that names a heading style: directly ("Heading7", which stays on
 *     `style` because office-open's HeadingLevel type caps at 6), by localized
 *     NAME ("heading 1"/"标题 1"), or via the `basedOn` chain (a custom style
 *     "MyTitle" basedOn="Heading1"). `heading` and `style` carry the same pStyle.
 *  `resolved` accepts a full office-open ParagraphOptions (parse/compile) OR a
 *  PM-node attrs subset — the editor outline walks PM nodes whose styleId /
 *  outlineLevel mark a heading without being a `type: "heading"` node (a
 *  paragraph the user styled as "Heading 1" at runtime). Pure (no `this`):
 *  resolved + the document styles snapshot are all it reads. */
export function detectHeadingLevel(
  resolved: { heading?: string; style?: string; outlineLevel?: number },
  styles: StylesOptions | undefined,
): number | undefined {
  if (resolved.heading) {
    const lvl = HEADING_PARSE_MAP[resolved.heading];
    if (lvl) return lvl;
  }
  const outline = resolved.outlineLevel;
  if (typeof outline === "number" && outline >= 0 && outline <= 8) {
    return outline + 1;
  }
  const styleId = resolved.style;
  if (!styleId || !styles) return undefined;
  const byId = indexParagraphStyles(styles);
  const visited = new Set<string>();
  let curId: string | undefined = styleId;
  while (curId && !visited.has(curId)) {
    visited.add(curId);
    if (HEADING_PARSE_MAP[curId]) return HEADING_PARSE_MAP[curId];
    const style = byId.get(curId);
    if (!style) break;
    const lvl = headingLevelFromName(style.name);
    if (lvl) return lvl;
    curId = style.basedOn ?? undefined;
  }
  return undefined;
}

// DOCX heading paragraph → heading node. detectHeadingLevel covers office-open's
// lifted `heading` literal, an explicit outlineLevel, or a pStyle naming a
// heading style (directly, by localized name, or via basedOn). The real pStyle
// still rides on attrs.styleId (heading parseDocx carries resolved.style);
// buildTextBlock stamps the detected level when parseDocx couldn't derive it.
export const parseDocxParagraph: ParseParagraphRule = {
  match: (para, ctx) => detectHeadingLevel(para, ctx.styles) != null,
  convert: (para, ctx) => {
    const level = detectHeadingLevel(para, ctx.styles);
    if (!level) return null;
    return buildTextBlock("heading", para, ctx, level);
  },
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
  // Narrow by typeof, not `as number`: a stray string level (e.g. "7" from
  // malformed JSON) would otherwise index HEADING_COMPILE_MAP as a string.
  const level = typeof attrs.level === "number" ? attrs.level : undefined;
  const styleId = attrs.styleId as string | undefined;
  if (styleId) opts.heading = styleId;
  else if (level) opts.heading = HEADING_COMPILE_MAP[level] ?? "Heading1";

  // Pass remaining attrs through verbatim (skip nulls + mapped fields). Section
  // attrs (sectionProperties/Headers/Footers) are editor-only — DocxManager
  // closes a section off them in compile, so they must NOT reach ParagraphOptions
  // (a heading can be a section's last paragraph, same as a plain paragraph).
  for (const [key, value] of Object.entries(attrs)) {
    if (key === "level" || key === "styleId" || TOC_RUNTIME_KEYS.has(key)) continue;
    if (SECTION_ATTR_KEYS.has(key)) continue;
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
  // HeadingLevel pStyle ("Heading1".."Heading9"/"Title") into `heading`; that
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

// ── Extension ──

export const Heading = BaseHeading.extend({
  // Levels 1-9: DOCX carries Heading1-Heading9 (office-open built-in styles).
  // HTML has only h1-h6, so 7-9 render as <h6 data-heading-level="N"> (see
  // renderHTML/parseHTML) and round-trip via styleId / HEADING_*_MAP. The
  // schema's `level` attr is unvalidated, so 7-9 are valid; setHeading gates on
  // options.levels (Tiptap's Level type caps at 1-6), so applyStyle uses
  // setNode for 7-9 rather than widening options.levels (which the type forbids).
  // Override parseHTML so 7-9 (rendered as <h6 data-heading-level>) parse back
  // to their real level; the data-heading-level rule runs before the native
  // h6 rule so a plain <h6> still maps to level 6.
  parseHTML() {
    return [
      {
        tag: "h6[data-heading-level]",
        getAttrs: (el) => {
          const level = Number((el as HTMLElement).getAttribute("data-heading-level"));
          // Missing/empty/non-integer attr falls back to 6 (a plain <h6>);
          // returning NaN/0 would violate the 1-9 union and corrupt the node.
          return { level: Number.isInteger(level) && level >= 1 && level <= 9 ? level : 6 };
        },
      },
      ...[1, 2, 3, 4, 5, 6].map((level) => ({ tag: `h${level}`, attrs: { level } })),
    ];
  },
  // A heading is a paragraph in OOXML, so it shares Paragraph's office-open attrs
  // via docxParagraphAttrs — only `level` differs, and this.parent (BaseHeading)
  // provides it.
  addAttributes() {
    return { ...this.parent?.(), ...docxParagraphAttrs() };
  },

  renderHTML({ node, HTMLAttributes }: { node: Node; HTMLAttributes: Record<string, unknown> }) {
    const level = (node.attrs?.level as number) ?? 1;
    // HTML has no h7-h9: levels 7-9 render as <h6> carrying the real level in
    // data-heading-level (parseHTML reads it back). renderTextBlock stamps it.
    const tag = level >= 1 && level <= 6 ? `h${level}` : "h6";
    return renderTextBlock(node, HTMLAttributes, tag, level);
  },

  renderDocx,
  parseDocx,
  parseDocxParagraph,
});
