import { createDocument, type SectionPropertiesOptions } from "@docen/docx";
import { Node } from "@docen/docx/core";

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

/** twips → mm (1in = 1440tw = 25.4mm), 2dp. */
const mm = (twips: number): string => `${((twips / 1440) * 25.4).toFixed(2)}mm`;

/** Resolve a section's printable page dimensions (twips), honoring orientation.
 *  A landscape section commonly stores portrait dimensions (w<h) with
 *  `orientation: "landscape"` — Word renders it landscape, so swap width/height
 *  to make width the larger edge. Exported because BOTH the page's rendered box
 *  (renderHTML) and its measured content box (sectionContentDims in page-plugin)
 *  must agree on orientation, or a page renders one size but is packed against
 *  another and never converges. Returns null when no numeric dimensions. */
export function resolvePageSize(size: unknown): { width: number; height: number } | null {
  if (!size || typeof size !== "object") return null;
  const s = size as { width?: unknown; height?: unknown; orientation?: unknown };
  const w = typeof s.width === "number" ? s.width : undefined;
  const h = typeof s.height === "number" ? s.height : undefined;
  if (w == null || h == null) return null;
  return s.orientation === "landscape" && w < h ? { width: h, height: w } : { width: w, height: h };
}

/** Inline geometry styles for a page from its section properties: paper size,
 *  margins (as padding inside the fixed box), and the document-grid line-height
 *  (Word snaps every line up to linePitch; normal paragraphs inherit it).
 *
 *  `SectionPropertiesOptions` is office-open's OOXML type — page.size/margin
 *  values are number|measure unions (parse always yields numbers; we narrow
 *  before arithmetic). grid.type "default" means no line-grid snapping. */
function pageGeometryStyles(sp: SectionPropertiesOptions | null | undefined): string[] {
  const styles: string[] = [];
  const dims = resolvePageSize(sp?.page?.size);
  if (dims) {
    // Orientation is resolved in resolvePageSize (landscape swaps width/height),
    // so width/height here are the VISUAL paper edges. Margins are left as-is:
    // office-open already returns them rotated for a landscape section.
    if (dims.width > 0) styles.push(`width:${mm(dims.width)}`);
    if (dims.height > 0) styles.push(`height:${mm(dims.height)}`);
  }
  const margin = sp?.page?.margin;
  if (margin) {
    const sides = [margin.top, margin.right, margin.bottom, margin.left];
    if (sides.every((s): s is number => typeof s === "number")) {
      styles.push(`padding:${sides.map(mm).join(" ")}`);
    }
  }
  const grid = sp?.grid;
  // linePitch (twips) snaps lines up when the grid type is line-snapping
  // (lines/linesAndChars/snapToChars). "default" = no snapping. twips→px = /15.
  if (grid?.linePitch && grid.type !== "default") {
    const pitchPx = (grid.linePitch / 15).toFixed(2);
    // The page's own line-height = one grid line (single spacing for paragraphs
    // that don't set their own). --docen-line-pitch lets paragraph line-spacing
    // MULTIPLES resolve relative to the grid (Word w:docGrid), via
    // calc(var(--docen-line-pitch) * m) in lineSpacingToCss — so "1.5 lines" is
    // 1.5 × pitch (Word), not 1.5 × fontSize (plain CSS).
    styles.push(`line-height:${pitchPx}px`);
    styles.push(`--docen-line-pitch:${pitchPx}px`);
  }
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
