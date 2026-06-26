import {
  createDocument,
  resolvePageSize,
  sectionLinePitchCss,
  sectionMarginCss,
  twipsToMm,
  type SectionPropertiesOptions,
} from "@docen/docx";
import { Node } from "@docen/docx/core";

// resolvePageSize moved to the @docen/docx engine; re-exported so page-plugin's
// existing import (`from "./page-node"`) keeps working without touching it.
export { resolvePageSize };

/**
 * Page node — a fixed-height paper sheet holding a slice of the document flow.
 *
 * Editing-time structure only: it never enters DOCX. The DOCX model is flat
 * `doc > block+`; the editor wraps content into pages on import and unwraps on
 * export (see `wrapPages`/`unwrapPages`), so round-trip is transparent. The
 * paginator physically splits/merges pages as content overflows the fixed box.
 *
 * Each page carries its **section's** `sectionProperties` (page size/margin/
 * orientation + document grid), so pages in different sections render at
 * different geometries (e.g. a landscape section) — matching Word, where
 * section layout is per-section. `renderHTML` applies that geometry inline.
 *
 * See CLAUDE.md → Pagination Architecture (C-route) and CONTRIBUTING.md →
 * Pagination Conventions.
 */

/** Inline geometry styles for a page from its section properties: paper size
 *  (the fixed box the paginator measures against), margins (padding), and the
 *  document-grid line-height (Word snaps every line up to linePitch; normal
 *  paragraphs inherit it). Delegates unit/geometry mapping to the @docen/docx
 *  engine (twipsToMm/sectionMarginCss/sectionLinePitchCss) so standalone HTML
 *  export (generateHTML) renders the same geometry. */
function pageGeometryStyles(sp: SectionPropertiesOptions | null | undefined): string[] {
  const styles: string[] = [];
  const dims = resolvePageSize(sp?.page?.size);
  if (dims) {
    // Orientation is resolved in resolvePageSize (landscape swaps width/height),
    // so width/height here are the VISUAL paper edges.
    if (dims.width > 0) styles.push(`width:${twipsToMm(dims.width)}`);
    if (dims.height > 0) styles.push(`height:${twipsToMm(dims.height)}`);
  }
  const padding = sectionMarginCss(sp?.page?.margin);
  if (padding) styles.push(padding);
  styles.push(...sectionLinePitchCss(sp?.grid));
  return styles;
}

export const Page = Node.create({
  name: "page",
  // No group → a page can only be a direct child of `doc`, never nested inside
  // another block. Keeps the page structure flat and unambiguous for the
  // paginator (it walks doc.children and knows every child is a page).
  group: "",
  content: "block+",
  // Backspace/merge at a page boundary stays within the page — the paginator,
  // not the editor's delete logic, decides when content moves between pages.
  isolating: true,
  defining: true,

  addAttributes() {
    return {
      // This page's section properties (page size/margin/orientation + grid).
      // Filled by wrapPages (initial) and reflow (per page, by its section) so
      // each page renders at its own geometry. Stored as JSON in a data-* attr.
      sectionProperties: {
        default: null,
        parseHTML: (element: HTMLElement) => {
          const raw = element.getAttribute("data-section-properties");
          if (!raw) return null;
          try {
            return JSON.parse(raw);
          } catch {
            return null;
          }
        },
        rendered: false,
      },
    };
  },

  parseHTML() {
    return [{ tag: "div.docen-page" }];
  },
  renderHTML({
    node,
    HTMLAttributes,
  }: {
    node: { attrs: Record<string, unknown> };
    HTMLAttributes: Record<string, unknown>;
  }) {
    const attrs: Record<string, unknown> = { ...HTMLAttributes, class: "docen-page" };
    const styles = pageGeometryStyles(
      node.attrs.sectionProperties as SectionPropertiesOptions | null | undefined,
    );
    if (styles.length > 0) attrs.style = styles.join(";");
    return ["div", attrs, 0] as const;
  },
});

/**
 * Document schema for the editing shape: `doc > page+` instead of docx's default
 * `doc > block+`. Built from docx's `createDocument` factory so the Document
 * definition has a SINGLE source (the docx package) — the editor parameterizes
 * only the content expression, instead of `.extend`-overriding docx's Document
 * and relying on whatever attrs that happened to inherit. Registered after
 * docx's Document in `createDocxEditor`'s extension list, so Tiptap's same-name
 * override picks this `page+` variant.
 */
export const PageDocument = createDocument("page+");
