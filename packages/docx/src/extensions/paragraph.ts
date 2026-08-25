import type { ParagraphOptions, StylesOptions } from "@office-open/docx";
import type { JSONContent, MarkdownRendererHelpers, RenderContext } from "@tiptap/core";
import { Paragraph as BaseParagraph } from "@tiptap/extension-paragraph";
import type { Node } from "@tiptap/pm/model";

import { indexParagraphStyles } from "../style-cascade";
import { HTML_ORDERED_TEMP } from "./list-numbering";
import { docxParagraphAttrs, renderTextBlock, SECTION_ATTR_KEYS } from "./utils";

/**
 * Paragraph extension with nested office-open attrs — the single textblock node.
 *
 * Attrs mirror ParagraphPropertiesOptionsBase verbatim (alignment/indent/
 * spacing/border/shading/frame as nested objects + every scalar OOXML property
 * + heading/style/bullet/numbering/thematicBreak): a heading IS a paragraph in
 * OOXML, so its HeadingLevel pStyle rides on the `heading` attr instead of a
 * separate node, and DOCX round-trip is near-identity — renderDocx/parseDocx
 * pass attrs through; CSS conversion happens only in renderHTML via utils
 * mappers. Consumers derive the display level via detectHeadingLevel.
 */

/** HeadingLevel literals: "Heading1".."Heading9", "Title". */
export const HEADING_COMPILE_MAP: Record<number, string> = {
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
 *  PM-node attrs subset — the editor outline walks PM nodes whose style /
 *  outlineLevel mark a heading without a lifted `heading` attr (a numeric
 *  pStyle id common in WPS / Chinese Word). Pure (no `this`): resolved + the
 *  document styles snapshot are all it reads. */
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

/** The heading tag for renderHTML: levels 1-6 map to h1-h6; 7-9 fall back to
 *  a <h6 data-heading-level="N"> proxy (HTML has no h7-h9; parseHTML reads the
 *  data attr back). A non-heading paragraph renders as <p>. */
function headingTag(resolved: { heading?: string; style?: string; outlineLevel?: number }): {
  tag: string;
  level?: number;
} {
  const level = detectHeadingLevel(resolved, undefined);
  if (!level) return { tag: "p" };
  return { tag: level >= 1 && level <= 6 ? `h${level}` : "h6", level };
}

// ── DOCX serialization (near-identity: attrs mirror ParagraphPropertiesOptionsBase) ──

export function renderDocx(node: JSONContent): Record<string, unknown> {
  const attrs = (node.attrs ?? {}) as Record<string, unknown>;
  const opts: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (value === null || value === undefined) continue;
    if (SECTION_ATTR_KEYS.has(key)) continue;
    // Runtime-only attrs the TableOfContents extension injects on each heading
    // (id / data-toc-id) are regenerated on every load — never persist them.
    if (key === "id" || key === "data-toc-id") continue;
    opts[key] = value;
  }
  return opts;
}

/**
 * Structural/semantic keys handled elsewhere (run/text children — `run` is
 * intentionally NOT skipped: ParagraphOptions.run is the paragraph's default
 * run properties, kept as an attr for lossless round-trip, e.g. header/footer
 * paragraphs whose styling lives there).
 */
const SKIP_KEYS = new Set(["children", "text"]);

export function parseDocx(opts: ParagraphOptions | string): Record<string, unknown> {
  const resolved: ParagraphOptions = typeof opts === "string" ? { text: opts } : opts;
  const attrs: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(resolved)) {
    if (SKIP_KEYS.has(key)) continue;
    attrs[key] = value ?? null;
  }
  return attrs;
}

// ── Extension ──

export const Paragraph = BaseParagraph.extend({
  // A heading is a paragraph in OOXML (a <w:p> with pStyle="Heading1"), so the
  // paragraph node carries the full office-open mirror — heading/style/bullet/
  // numbering/thematicBreak included. See utils.
  addAttributes() {
    return { ...this.parent?.(), ...docxParagraphAttrs() };
  },

  // HTML round-trip: h1-h6 parse back with a lifted HeadingLevel `heading`
  // attr; a <h6 data-heading-level="N"> proxy restores levels 7-9. The proxy
  // rule runs before the native h6 rule so a plain <h6> maps to level 6.
  // A <li> maps to a flat list paragraph: nesting depth = ancestor ul/ol
  // count, and the NEAREST list element decides bullet vs ordered (an <ol>
  // wrapping a nested <ul> marks that sublist's items as bullets). Ordered
  // items carry the html-ordered placeholder reference — parseHTML
  // (converters/html.ts) rewrites each consecutive run to a fresh
  // docen-ordered-* reference so independent lists number separately.
  parseHTML() {
    return [
      {
        tag: "li",
        getAttrs: (el) => {
          let depth = 0;
          let nearest: string | null = null;
          for (let p = el.parentElement; p; p = p.parentElement) {
            const tag = p.tagName.toUpperCase();
            if (tag === "UL" || tag === "OL") {
              depth++;
              nearest ??= tag;
            }
          }
          // The nearest list is depth 1 → nesting level 0.
          const level = Math.max(0, depth - 1);
          return nearest === "OL"
            ? { numbering: { reference: HTML_ORDERED_TEMP, level } }
            : { bullet: { level } };
        },
      },
      {
        tag: "h6[data-heading-level]",
        getAttrs: (el) => {
          const level = Number((el as HTMLElement).getAttribute("data-heading-level"));
          return {
            heading:
              HEADING_COMPILE_MAP[Number.isInteger(level) && level >= 1 && level <= 9 ? level : 6],
          };
        },
      },
      ...[1, 2, 3, 4, 5, 6].map((level) => ({
        tag: `h${level}`,
        attrs: { heading: HEADING_COMPILE_MAP[level] },
      })),
      { tag: "p" },
    ];
  },

  renderHTML({ node, HTMLAttributes }: { node: Node; HTMLAttributes: Record<string, unknown> }) {
    // renderHTML has no access to the document styles, so the tag decision
    // covers the lifted `heading` attr and an explicit outlineLevel only; a
    // numeric pStyle id resolved through styles.xml (WPS/zh-CN Word) falls
    // back to <p> — its docx-style-{id} CSS class still applies the style.
    const { tag, level } = headingTag(
      node.attrs as { heading?: string; style?: string; outlineLevel?: number },
    );
    // A list paragraph marks itself (kind + depth) so generateHTML can regroup
    // consecutive list paragraphs into nested ul/ol lists — parseHTML's <li>
    // rule reads the grouping back.
    const attrs = node.attrs as { bullet?: { level?: number } | null; numbering?: unknown };
    const withList = { ...HTMLAttributes };
    if (attrs.bullet) {
      withList["data-list"] = "bullet";
      withList["data-list-level"] = String(attrs.bullet.level ?? 0);
    } else if (attrs.numbering) {
      withList["data-list"] = "ordered";
      withList["data-list-level"] = String((attrs.numbering as { level?: number }).level ?? 0);
    }
    return renderTextBlock(node, withList, tag, level);
  },

  renderDocx,
  parseDocx,

  // Markdown serialization: a heading paragraph renders as "#{level} text";
  // everything else keeps the upstream paragraph semantics (empty paragraphs
  // emit the &nbsp; empty-paragraph marker between consecutive empties).
  renderMarkdown: (node: JSONContent, h: MarkdownRendererHelpers, ctx: RenderContext): string => {
    const attrs = (node.attrs ?? {}) as { heading?: string | null };
    const level = attrs.heading ? HEADING_PARSE_MAP[attrs.heading] : undefined;
    const content = h.renderChildren(Array.isArray(node.content) ? node.content : []);
    if (level) return `${"#".repeat(level)} ${content}`;
    if (!node.content || node.content.length === 0) {
      const prev = ctx?.previousNode;
      const prevIsEmptyParagraph =
        prev?.type === "paragraph" && (!prev.content || prev.content.length === 0);
      return prevIsEmptyParagraph ? "&nbsp;" : "";
    }
    return content;
  },
});
