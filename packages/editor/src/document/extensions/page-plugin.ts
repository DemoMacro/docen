import {
  sectionMarginDefaults,
  scrollCaretToTop,
  type SectionPropertiesOptions,
} from "@docen/docx";
import { Extension, type Editor } from "@docen/docx/core";
import type { Node as PmNode } from "@tiptap/pm/model";
import { Plugin, PluginKey, TextSelection } from "@tiptap/pm/state";

import {
  type CellMargins,
  measureBlockHeight,
  measureParagraphLines,
  measureRowHeight,
  paragraphFloatZonesOf,
  paragraphSpacingMargins,
  resolveIndentWidth,
  resolvePaginationAttrs,
  tableColumnWidths,
  tableWidthOf,
  type FloatZone,
  type PaginationAttrs,
} from "../utils/measure";
import { resolvePageSize } from "./page-node";
import { cloneHeaderRows } from "./split-table";

/** Plugin key marking a re-flow transaction (so the update listener can tell
 *  a user edit from our own regrouping). Exported so the host's change emitter
 *  can skip reflow transactions — they re-pack content into editor-only page
 *  nodes, so the flat doc (host.getJSON) is unchanged. ProseMirror convention:
 *  a plugin tags its own transactions via its PluginKey meta. */
export const flowKey = new PluginKey("pageFlow");

export interface PagePluginOptions {
  /** Debounce (ms) before re-flowing after an edit. Default 300. */
  debounceMs?: number;
}

export interface PagePluginStorage {
  /** Force a re-flow now (bypasses debounce). The host calls this after an
   *  import or a page-geometry change, once layout has settled. */
  repaginate: () => void;
}

export function pageStorageOf(editor: Editor): PagePluginStorage {
  return (editor.storage as unknown as { docenPagePlugin: PagePluginStorage }).docenPagePlugin;
}

/** Per-page usable content height (px): the fixed page box's height minus its
 *  vertical padding. Uses getBoundingClientRect().height (sub-pixel exact), NOT
 *  clientHeight — clientHeight/offsetHeight round to whole pixels, so the
 *  page boundary wobbles ±1px between re-flows and the packer never converges
 *  (visible as table/row flicker). Sub-pixel measurement makes the boundary
 *  stable. Returns 0 until the page is laid out. */
function resolvePageContentHeight(editor: Editor): number {
  const page = editor.view.dom.querySelector<HTMLElement>(".docen-page");
  if (!page) return 0;
  const cs = getComputedStyle(page);
  const padY = (parseFloat(cs.paddingTop) || 0) + (parseFloat(cs.paddingBottom) || 0);
  return Math.max(0, page.getBoundingClientRect().height - padY);
}

/** A flattened flow item: a standalone block, or a single table row carrying
 *  its parent table (so the paginator can regroup rows back into per-page
 *  table nodes). Tables expand to rows because contenteditable cannot visually
 *  split a `<tr>` — only whole rows move between pages. */
/** Per-line geometry for a splittable textblock (from measureParagraphLines).
 *  `lineBreakOffsets[i]` = PM content offset at end of line i+1. */
type BlockLines = {
  lineHeight: number;
  lineCount: number;
  lineBreakOffsets: number[];
  block?: boolean;
  /** Per-row heights when rows are non-uniform (inline-image rows: each row is
   *  its tallest image). Present ⇒ trySplitBlock walks these instead of
   *  `lineHeight × n`. Absent ⇒ uniform text lines (`lineHeight × n` is exact). */
  lineHeights?: number[];
};

type FlatItem =
  | {
      kind: "block";
      node: PmNode;
      height: number;
      after: number;
      /** Effective OOXML pagination props (keepLines/keepNext/widowControl/
       *  pageBreakBefore) — null→OOXML default resolved. */
      pag: PaginationAttrs;
      /** Present when the block is splittable mid-paragraph (text lines, or
       *  inline-image boundaries). Absent ⇒ the block moves whole (keepLines,
       *  pageBreak atom, section-end paragraph). */
      lines?: BlockLines;
    }
  | { kind: "row"; row: PmNode; height: number; table: PmNode; isClone?: boolean };

/** Page content box + document-grid line pitch, derived deterministically from a
 *  section's OOXML geometry (twips → px) — NOT the DOM. Each page renders its
 *  section's inline width/height/padding (page-node renderHTML), so the content
 *  box is `paper size − margins`; linePitch snaps lines up when the grid type is
 *  line-snapping (lines/linesAndChars/snapToChars; "default" = no snapping).
 *  Absent geometry falls back to the engine's section-geometry defaults
 *  (@office-open/docx sectionPageSizeDefaults/sectionMarginDefaults — the same
 *  values the engine fills into an empty sectPr), so a blank doc with no sectPr
 *  still measures A4 + MS Office zh-CN "Normal" margins instead of dropping to
 *  a divergent DOM default. Twip→px matches the rendered box exactly (size/15). */
export function sectionContentDims(sp: unknown): {
  width: number;
  height: number;
  linePitchPx: number | undefined;
} {
  const s = (sp && typeof sp === "object" ? sp : {}) as SectionPropertiesOptions;
  const dims = resolvePageSize(s.page?.size);
  const twipToPx = 4 / 3 / 20; // twip → pt (÷20) → px (×4/3)
  // Margins mirror page-node renderHTML's sectionMarginCss: each side falls back
  // to the engine default (@office-open/docx sectionMarginDefaults) when absent,
  // so measure == render even for a section whose <w:pgMar> omits sides (or a
  // blank doc with no sectPr at all — page size also falls back to the engine
  // default inside resolvePageSize, so a blank doc measures A4 + MS Office
  // zh-CN "Normal" margins like the engine fills into an empty sectPr).
  const margin = s.page?.margin;
  const def = sectionMarginDefaults;
  const num = (v: unknown, d: number): number => (typeof v === "number" ? v : d);
  const mTop = num(margin?.top, def.TOP);
  const mRight = num(margin?.right, def.RIGHT);
  const mBottom = num(margin?.bottom, def.BOTTOM);
  const mLeft = num(margin?.left, def.LEFT);
  const grid = s.grid;
  // OOXML: docGrid @type omitted or "default" = NO grid — lines do NOT snap to
  // @linePitch (Word renders at the font's natural metric). Only
  // lines/linesAndChars/snapToChars snap. @type is absent on many Western docs
  // which still carry a @linePitch; treating absent as "snap" added 18pt/line
  // and inflated pagination ~60%. Must mirror
  // sectionLinePitchCss (render) so measure == render.
  const linePitchPx =
    grid?.linePitch && grid.type && grid.type !== "default" ? grid.linePitch * twipToPx : undefined;
  return {
    width: (dims.width - mLeft - mRight) * twipToPx,
    height: (dims.height - mTop - mBottom) * twipToPx,
    linePitchPx,
  };
}

/** Flatten every page's blocks (in flow order) into block/row items and measure
 *  each one's height. Blocks use Pretext (deterministic canvas measurement);
 *  a table expands to its rows (measured per `<tr>` via DOM, until row-level
 *  Pretext measurement lands). Continuation-page header clones (`splitClone`
 *  rows) are skipped so they are re-derived, not doubled, on the next re-flow.
 *  ProseMirror widgets skipped. */
function measureFlatItems(editor: Editor): FlatItem[] {
  const { view, state } = editor;
  const styles = (state.doc.attrs as { styles?: unknown }).styles;
  // A page's sectionProperties are filled by the reflow that created it (each
  // page renders its section's geometry). On the FIRST reflow after load the doc
  // is a single placeholder page whose sectionProperties are absent — the OOXML
  // sectPr rides on the section's last paragraph / doc.attrs, not the page — so
  // measuring that page read pitch=undefined and fell back to the font-natural
  // metric, under-counting every row. The next pass then re-measured with the
  // real grid pitch (pages now carry it), so row heights changed between passes
  // (28→32px) and pagination oscillated, drifting a row onto the next page vs
  // Word. Fall back to the document-level section so the first pass measures
  // with the real grid pitch; single-section docs converge in one pass.
  const docSectionProperties =
    (state.doc.attrs as { sectionProperties?: unknown }).sectionProperties ?? null;
  const pageDoms = Array.from(view.dom.children).filter(
    (el): el is HTMLElement => el instanceof HTMLElement && el.classList.contains("docen-page"),
  );
  type Raw = {
    node: PmNode;
    dom: HTMLElement | undefined;
    width: number;
    linePitchPx: number | undefined;
    pageIndex: number;
  };
  // Pass 1 — flatten every page's children into one flow, each carrying its
  // page's content width + document-grid pitch (per-page section geometry) and
  // its DOM handle. Fall back to the laid-out DOM rect only when no section
  // geometry is set.
  const raw: Raw[] = [];
  let pi = 0;
  state.doc.forEach((pageNode) => {
    if (pageNode.type.name !== "page") return;
    const curPi = pi++;
    const pageDom = pageDoms[curPi];
    if (!pageDom) return;
    const dims = sectionContentDims(pageNode.attrs.sectionProperties ?? docSectionProperties);
    const pageCs = getComputedStyle(pageDom);
    const pageRect = pageDom.getBoundingClientRect();
    const padX = (parseFloat(pageCs.paddingLeft) || 0) + (parseFloat(pageCs.paddingRight) || 0);
    const width = dims?.width ?? Math.max(0, pageRect.width - padX);
    const childDoms = Array.from(pageDom.children).filter(
      (el): el is HTMLElement =>
        el instanceof HTMLElement && !el.classList.contains("ProseMirror-widget"),
    );
    let bi = 0;
    pageNode.forEach((child) => {
      const dom = childDoms[bi++];
      if (!dom) return;
      raw.push({ node: child, dom, width, linePitchPx: dims?.linePitchPx, pageIndex: curPi });
    });
  });
  // Pass 2 — CROSS-PAGE logical merge: adjacent (in flow order, possibly across
  // a page boundary) same-splitGroup paragraphs/headings merge back into the
  // original whole paragraph BEFORE measuring. This makes re-flow idempotent —
  // a paragraph split across pages last pass is re-merged and re-split from the
  // original each pass, so the break can't nest or drift (the head re-splitting
  // under a fresh id every pass left only orphan tails and never converged).
  // Tables/leaves never carry splitGroup, so they pass through untouched.
  const logical: Raw[] = [];
  for (const r of raw) {
    const last = logical[logical.length - 1];
    const group = (r.node.attrs as { splitGroup?: string | null }).splitGroup;
    const isTextblockLike = r.node.type.name === "paragraph" || r.node.type.name === "heading";
    if (
      group != null &&
      isTextblockLike &&
      last &&
      (last.node.type.name === "paragraph" || last.node.type.name === "heading") &&
      (last.node.attrs as { splitGroup?: string | null }).splitGroup === group
    ) {
      const merged = last.node.type.create(
        last.node.attrs,
        last.node.content.append(r.node.content),
        last.node.marks,
      );
      logical[logical.length - 1] = { ...last, node: merged };
    } else {
      logical.push(r);
    }
  }
  // Pass 3 — measure. Tables expand to rows; text blocks use Pretext at the
  // paragraph's USABLE width (page width minus its own indent — resolveIndentWidth
  // mirrors the renderer, else an indented paragraph under-counts lines and
  // overflows the page, and the split offset lands mid-word); leaves fall back
  // to DOM height. Continuation header clones (splitClone rows) reserve height
  // against the next real row, never as their own flow item.
  const out: FlatItem[] = [];
  // Active float zones (page-content-relative) + accumulated Y, reset per page.
  // A square/tight float image rides its anchor paragraph's flow and overhangs
  // later paragraphs; tracking its band lets those paragraphs measure with the
  // reduced width (text wraps beside the image) — without this they under-count
  // wrapped lines and the page overflows after reflow.
  let activeFloats: FloatZone[] = [];
  let accY = 0;
  let curPi = -1;
  for (const { node: child, dom, width: pageContentWidth, linePitchPx, pageIndex } of logical) {
    if (pageIndex !== curPi) {
      curPi = pageIndex;
      activeFloats = [];
      accY = 0;
    }
    const blockTop = accY;
    if (child.type.name === "table") {
      const tableW = tableWidthOf(child, pageContentWidth);
      const colWidths = tableColumnWidths(child, tableW);
      // Table-level cell insets (w:tblCellMar): threaded into the measure ctx so
      // a cell lacking its own tcMar inherits the table default (mirrors the
      // renderer). See effectiveCellMargins in measure.ts.
      const tableCellMargins = (child.attrs as { margins?: CellMargins | null }).margins ?? null;
      let pendingClone = 0;
      let tableH = 0;
      child.forEach((row) => {
        const rh = measureRowHeight(row, colWidths, { linePitchPx, styles, tableCellMargins });
        if (row.attrs.splitClone === true) {
          pendingClone += rh;
          return;
        }
        out.push({ kind: "row", row, height: rh + pendingClone, table: child });
        tableH += rh + pendingClone;
        pendingClone = 0;
      });
      accY += tableH;
    } else {
      const after = paragraphSpacingMargins(child, styles).afterPx;
      const blockWidth = child.isTextblock
        ? resolveIndentWidth(child, pageContentWidth, styles)
        : pageContentWidth;
      // The block's own float images join the active set BEFORE measuring: the
      // block's text wraps around its own float (the image shares its lines),
      // and subsequent blocks the image overhangs must see it too. Push at
      // blockTop so the band aligns with this block's first line.
      if (child.isTextblock) {
        for (const z of paragraphFloatZonesOf(child, blockTop)) activeFloats.push(z);
      }
      // floatZones + startY: this block's lines shed width to any active float
      // overlapping them (CSS float text-wrap), measured per line via Pretext.
      const ctx = {
        domHeightOf: (n: PmNode) => (n === child ? dom?.getBoundingClientRect().height : undefined),
        linePitchPx,
        styles,
        floatZones: activeFloats,
        startY: blockTop,
      };
      // Textblock height MUST stay Pretext (model-only) — never read the DOM
      // here. A textblock's DOM height wobbles across layout passes (sub-pixel
      // font hinting) and collapses to ~0 on off-screen pages (content-visibility:
      // auto); preferring it makes boundary blocks flip pages every re-flow pass
      // (A↔B oscillation, visible as the page contents flickering up/down). This
      // is the same invariant measureBlockHeight enforces — see its doc comment.
      const height = measureBlockHeight(child, blockWidth, ctx);
      const pag = resolvePaginationAttrs(child);
      const lines = measureParagraphLines(child, blockWidth, ctx) ?? undefined;
      accY += height + after;
      out.push({ kind: "block", node: child, height: height + after, after, pag, lines });
    }
  }
  return out;
}

/** Split a textblock at content offset `off` into head + tail, both carrying
 *  `splitGroup` (so unwrapPages merges them back) and `splitPart` (head/tail).
 *  attrs are preserved verbatim — spacing stays as the original on both halves;
 *  the page clips the final paragraph's `after` at layout time, not here, so
 *  merging restores the paragraph exactly. marks ride along on the text nodes
 *  (content.cut keeps them). */
function splitTextblock(
  node: PmNode,
  off: number,
  splitGroup: string,
): { head: PmNode; tail: PmNode } {
  const head = node.type.create(
    { ...node.attrs, splitGroup, splitPart: "head" },
    node.content.cut(0, off),
    node.marks,
  );
  // The tail is a CONTINUATION of the original paragraph — its first line is a
  // mid-paragraph line, not a paragraph start — so it must NOT inherit the
  // first-line indent. Word's paragraph-internal page break leaves the continued
  // line flush; without this override the tail kept the docDefault
  // firstLineChars indent and its first line came up 2 chars short. Left/right
  // indent stays (the whole paragraph shares it); only firstLine is zeroed.
  const tail = node.type.create(
    {
      ...node.attrs,
      splitGroup,
      splitPart: "tail",
      indent: { ...node.attrs.indent, firstLine: 0, firstLineChars: 0 },
    },
    node.content.cut(off),
    node.marks,
  );
  return { head, tail };
}

/** Try to split a block so its first lines fill `remaining` px on the current
 *  page. Returns head/tail FlatItems, or null when the block must move whole:
 *  - no `lines` (a container, or a non-textblock leaf), keepLines ON, or a
 *    single line;
 *  - widow/orphan can't be satisfied (widowControl ON ⇒ ≥2 lines at each page
 *    edge; OFF ⇒ ≥1) — Word's widowControl default ON behavior (CSS orphans/
 *    widows = 2), implemented manually because C-route's physical pages can't
 *    use CSS fragmentation.
 *  Head is the page's last item (after clipped ⇒ height = N×lineHeight, no
 *  after); tail keeps the original `after` for the next page AND its remaining
 *  row geometry (`lines`), so the packer can split it again the SAME pass — a
 *  >2-page paragraph then fills every page in one re-flow instead of stranding
 *  the rest on one oversized clipped page. */
function trySplitBlock(
  item: Extract<FlatItem, { kind: "block" }>,
  remaining: number,
  splitGroup: string,
  keepNext = false,
): { head: FlatItem; tail: FlatItem } | null {
  if (!item.lines || item.pag.keepLines) return null;
  const { lineHeight, lineCount, lineBreakOffsets, block, lineHeights } = item.lines;
  if (lineCount <= 1) return null;
  // widow/orphan applies to TEXT lines (don't strand a single line alone at a
  // page edge). Inline-image rows (block: true) are self-contained blocks —
  // each can sit alone on a page (a near-full-page image is one per page), so
  // they skip widow/orphan (edge 1). Without this, a multi-image paragraph of
  // near-page-height images couldn't split (n=1 < edge=2 → move whole → clip).
  // keepNext: this paragraph follows a keepNext heading (Word's heading default),
  // so it must stay on the heading's page — relax the orphan rule to 1 line so
  // the heading isn't stranded alone (e.g. a section heading + its body).
  const edge = block ? 1 : keepNext ? 1 : item.pag.widowControl ? 2 : 1; // orphans + widows
  // How many leading rows fit in `remaining`, and their actual height. Text
  // rows are uniform (lineHeight × n); image rows are not (each row is its
  // tallest image), so walk lineHeights and stop before exceeding remaining.
  // The first row is always placed even if it alone exceeds remaining — a
  // single over-tall image row still renders on the page (clipping is the page
  // box's job); refusing it would move the whole block and strand every later
  // image on the next page.
  const headHeightOf = (k: number): number =>
    lineHeights ? lineHeights.slice(0, k).reduce((s, h) => s + h, 0) : k * lineHeight;
  let n: number;
  if (lineHeights) {
    let acc = 0;
    n = 0;
    for (let i = 0; i < lineCount; i++) {
      const h = lineHeights[i] ?? 0;
      if (n > 0 && acc + h > remaining) break;
      acc += h;
      n++;
    }
  } else {
    n = Math.floor(remaining / lineHeight);
  }
  if (n >= lineCount) return null; // whole block fits — shouldn't reach here
  if (n < edge) return null; // current page can't hold `orphans` lines
  if (lineCount - n < edge) {
    // next page would get < `widows` lines → shrink head to give the tail more
    n = lineCount - edge;
    if (n < edge) return null; // still can't satisfy orphans → move whole
  }
  if (n <= 0) return null;
  const off = lineBreakOffsets[n - 1];
  const totalH = lineHeights ? lineHeights.reduce((s, h) => s + h, 0) : lineCount * lineHeight;
  const { head, tail } = splitTextblock(item.node, off, splitGroup);
  const headItem: FlatItem = {
    kind: "block",
    node: head,
    height: headHeightOf(n), // page-last ⇒ after clipped
    after: 0,
    pag: item.pag,
  };
  // The tail keeps the remaining rows' geometry so the packer can split it
  // again THIS pass — a >2-page paragraph then fills every page in one re-flow,
  // instead of splitting off only the head and stranding the rest on one
  // oversized page (clipped). lineBreakOffsets are re-based to the tail's
  // content (which starts at `off`).
  const tailLines: BlockLines | undefined = item.lines
    ? {
        lineHeight,
        lineCount: lineCount - n,
        lineBreakOffsets: lineBreakOffsets.slice(n).map((o) => o - off),
        block,
        lineHeights: lineHeights?.slice(n),
      }
    : undefined;
  const tailItem: FlatItem = {
    kind: "block",
    node: tail,
    height: totalH - headHeightOf(n) + item.after,
    after: item.after,
    pag: item.pag,
    lines: tailLines,
  };
  return { head: headItem, tail: tailItem };
}

/** Flat caret units across all pages, in flow order: a standalone block, or a
 *  single table row. Tables expand to rows because the paginator regroups rows
 *  (never reorders or edits them), so the flat row index stays stable across a
 *  re-flow even when a table splits or merges — keeping the caret in place
 *  where the old top-level-block index broke. Continuation-page header clones
 *  (splitClone rows) are skipped: they are re-derived each re-flow and never
 *  hold the caret. */
function flatCaretUnits(doc: PmNode): Array<{ start: number; end: number }> {
  const units: Array<{ start: number; end: number }> = [];
  doc.forEach((pageNode, pageOffset) => {
    if (pageNode.type.name !== "page") return;
    pageNode.forEach((child, childOffset) => {
      const blockStart = pageOffset + 1 + childOffset;
      if (child.type.name === "table") {
        child.forEach((row, rowOffset) => {
          if (row.attrs.splitClone === true) return;
          const rowStart = blockStart + 1 + rowOffset;
          units.push({ start: rowStart, end: rowStart + row.nodeSize });
        });
      } else {
        units.push({ start: blockStart, end: blockStart + child.nodeSize });
      }
    });
  });
  return units;
}

/** A caret position as (flat unit index, offset within unit) — stable across a
 *  re-flow: rows/blocks keep their order (only their page grouping and table
 *  regrouping change), so the index+offset maps straight back. */
type CaretAnchor = { unitIndex: number; offset: number };

function caretAnchorAt(doc: PmNode, pos: number): CaretAnchor | null {
  const units = flatCaretUnits(doc);
  for (let i = 0; i < units.length; i++) {
    if (pos >= units[i].start && pos < units[i].end) {
      return { unitIndex: i, offset: pos - units[i].start };
    }
  }
  return null;
}

function restoreAnchor(doc: PmNode, anchor: CaretAnchor): number | null {
  const units = flatCaretUnits(doc);
  const u = units[anchor.unitIndex];
  if (!u) return null;
  const pos = Math.min(u.start + anchor.offset, u.end - 1);
  return Math.max(u.start + 1, pos);
}

/** Save BOTH ends of the selection so a RANGE selection survives a re-flow.
 *  Saving only `from` (as saveCaret did) collapsed every range: a debounced
 *  re-flow landing right after a double/triple-click or a drag-selection
 *  discarded `to` and restored a bare caret, so selecting text looked like it
 *  was immediately "cancelled". */
function saveSelection(
  doc: PmNode,
  selection: { from: number; to: number },
): { from: CaretAnchor | null; to: CaretAnchor | null } | null {
  // A table CellSelection (multi-cell) can't survive the re-flow: the full-doc
  // replaceWith maps it unpredictably, and restoreSelection only rebuilds a
  // TextSelection. Skip save/restore for any non-text selection so a re-flow
  // never collapses a multi-cell selection to a caret — the "selected, then
  // immediately cancelled" symptom.
  if (!(selection instanceof TextSelection)) return null;
  return {
    from: caretAnchorAt(doc, selection.from),
    to: selection.from === selection.to ? null : caretAnchorAt(doc, selection.to),
  };
}

function restoreSelection(
  doc: PmNode,
  saved: { from: CaretAnchor | null; to: CaretAnchor | null } | null,
): TextSelection | null {
  // null saved = a non-text selection (e.g. CellSelection) was skipped in
  // saveSelection; leave PM's own mapping in place instead of forcing a caret.
  if (!saved || !saved.from) return null;
  const from = restoreAnchor(doc, saved.from);
  if (from == null) return null;
  if (!saved.to) return TextSelection.create(doc, from);
  const to = restoreAnchor(doc, saved.to);
  if (to == null) return TextSelection.create(doc, from);
  return TextSelection.create(doc, from, Math.max(from, to));
}

/** Re-flow: regroup flow items (blocks + table rows) into pages so each page's
 *  content fits its fixed height. Greedy packing — a row never splits (a `<tr>`
 *  cannot visually break across pages). Dispatches nothing when the regrouping
 *  already matches the current page structure, which also terminates the
 *  re-flow → dispatch → update → re-flow cycle. */

/** Position of the FIRST pageBreak atom within a block — a descendant scan, so
 *  it also catches a break nested inside a list/table container. A top-level
 *  bulletList is NOT a textblock, so an `isTextblock`-only check missed a break
 *  at the start of a list item and left e.g. a section heading glued to the
 *  bottom of the preceding page.
 *  - "before": the atom is the block's leading inline (imported
 *    `<w:br w:type="page"/>` before the paragraph text) → the block starts a
 *    fresh page — Word breaks there, ahead of the text.
 *  - "after":  the atom is mid/trailing → the block ends a page.
 *  - null: no pageBreak atom. */
/** A paragraph holds real content beyond a pageBreak/hardBreak atom: a
 *  non-whitespace text run or an inline image. The pageBreak-carrier rule
 *  (core 0, no page space) excludes these — their images/text DO take space, so
 *  crediting them 0 stranded an image+break paragraph on its page and clipped
 *  it (pageBreak's atom has no textContent, so a textContent-only check wrongly
 *  treated an image paragraph as an empty carrier). */
function hasSubstantialContent(node: PmNode): boolean {
  let found = false;
  node.descendants((n) => {
    if ((n.isText && (n.text ?? "").replace(/\s/g, "").length > 0) || n.type.name === "image")
      found = true;
    return !found;
  });
  return found;
}
function pageBreakPosition(node: PmNode): "before" | "after" | null {
  // An empty paragraph whose only content is the pageBreak atom is the break's
  // CARRIER — Word renders it at the END of the preceding page (the break fires
  // there), never at the start of the next. So a text-less paragraph counts as
  // "after" (forcesPageBreakAfter closes the page after it), otherwise "before"
  // would strand the empty <p> atop the next page and drag the break mark down.
  const hasText = (node.textContent ?? "").replace(/\s/g, "").length > 0;
  let pos: "before" | "after" | null = null;
  let seenInline = false;
  node.descendants((n) => {
    if (pos) return false; // first pageBreak decides; skip the rest
    if (n.type.name === "pageBreak") {
      pos = !hasText || seenInline ? "after" : "before";
      return false;
    }
    if (n.isInline) seenInline = true;
    return true;
  });
  return pos;
}

/** A block that must START a new page: either a leading pageBreak atom
 *  (`<w:br w:type="page"/>` before its text) OR the OOXML `pageBreakBefore`
 *  paragraph property (`<w:pageBreakBefore/>`). The paginator closes the
 *  current page before it. */
function forcesPageBreakBefore(node: PmNode, pag?: PaginationAttrs): boolean {
  if (pag?.pageBreakBefore === true) return true;
  return pageBreakPosition(node) === "before";
}

/** A block forces a page break AFTER it when it ends a section (a paragraph
 *  carrying sectionProperties is its section's last paragraph — OOXML sectPr in
 *  its pPr — so the next block starts a fresh section/page) or when it holds a
 *  mid/trailing pageBreak atom. */
function forcesPageBreakAfter(node: PmNode): boolean {
  if ((node.attrs as { sectionProperties?: unknown }).sectionProperties != null) return true;
  return pageBreakPosition(node) === "after";
}

/** Identity key for split-grouping table rows: rows from tables already sharing
 *  a `splitGroup` (a prior split) regroup by that id; otherwise by node. */
function tableKeyOf(table: PmNode, ids: Map<PmNode, string>): string {
  const group = table.attrs.splitGroup as string | null;
  if (group) return group;
  let id = ids.get(table);
  if (!id) {
    id = `n${ids.size + 1}`;
    ids.set(table, id);
  }
  return id;
}

/** Column map of a sequence of rows: for each row, the grid column each
 *  present cell starts at, after skipping columns covered by a rowspan from an
 *  earlier row in the sequence. A vMerge-continue row resolves to ONE content
 *  cell plus an owner rowspan above it; this maps that cell to its column.
 *  Equivalent to prosemirror-tables' TableMap, kept local to avoid the dep. */
function mapRows(
  rows: PmNode[],
): Map<PmNode, Array<{ cell: PmNode; colStart: number; colspan: number }>> {
  const result = new Map<PmNode, Array<{ cell: PmNode; colStart: number; colspan: number }>>();
  const coveredUntil = new Map<number, number>(); // col -> row index covered up to (exclusive)
  let ri = 0;
  for (const row of rows) {
    const positions: Array<{ cell: PmNode; colStart: number; colspan: number }> = [];
    let col = 0;
    row.forEach((cell) => {
      const cs = (cell.attrs.colspan as number) ?? 1;
      const rs = (cell.attrs.rowspan as number) ?? 1;
      while ((coveredUntil.get(col) ?? 0) > ri) col++;
      positions.push({ cell, colStart: col, colspan: cs });
      if (rs > 1) for (let c = col; c < col + cs; c++) coveredUntil.set(c, ri + rs);
      col += cs;
    });
    result.set(row, positions);
    ri++;
  }
  return result;
}

function rowCellColumns(
  table: PmNode,
): Map<PmNode, Array<{ cell: PmNode; colStart: number; colspan: number }>> {
  const rows: PmNode[] = [];
  table.forEach((row) => rows.push(row));
  return mapRows(rows);
}

/** Grid column count of a table (furthest column reached by any row's cells). */
function tableColumnCount(table: PmNode): number {
  let max = 0;
  for (const positions of rowCellColumns(table).values()) {
    for (const { colStart, colspan } of positions) {
      const end = colStart + colspan;
      if (end > max) max = end;
    }
  }
  return max;
}

/** Per-column width (px) from the table's first row, used to give placeholder
 *  cells a colwidth matching their column so fixTables sees a self-consistent
 *  table and does not rewrite/pad it. */
function columnWidthsOf(table: PmNode, columnCount: number): number[] {
  const widths: number[] = Array.from({ length: columnCount }, () => 0);
  const firstRow = table.firstChild;
  if (!firstRow) return widths;
  let col = 0;
  firstRow.forEach((cell) => {
    const cs = (cell.attrs.colspan as number) ?? 1;
    const cw = (cell.attrs.colwidth as number[] | null) ?? [];
    for (let j = 0; j < cs && col < columnCount; j++) widths[col++] = cw[j] ?? 0;
  });
  return widths;
}

/** Rebuild continuation-page data rows that lost their rowspan anchor when the
 *  table split across pages. A vMerge-continue row resolves to one content
 *  cell; the merged columns were held by a rowspan owner. If that owner ALSO
 *  landed on this continuation page (a restart row earlier in `rows`, e.g.
 *  covering the next continue row), ProseMirror already places the content cell
 *  past the owner's rowspan — leave it. Only rows whose owner stayed on the
 *  PRIOR page have no rowspan covering their merged columns here, so the
 *  content cell collapses to column 0 and fixTables pads empty cells at the END
 *  (shifting column-3 content into column 1). For those rows, pad placeholder
 *  cells in front to restore the content cell to its source-table column.
 *  Comparing the continuation's own column map (clone header + these rows)
 *  against the source's tells the two cases apart: a content cell whose
 *  continuation column is already at/after its source column is covered by an
 *  on-page owner. Idempotent — a row rebuilt on a prior pass is met again at
 *  full width, and one absent from the source (source is itself a split
 *  segment next pass) is left as-is. */
function rebuildContinuationDataRows(
  rows: PmNode[],
  clonedHeader: PmNode[],
  source: PmNode,
  schema: PmNode["type"]["schema"],
): PmNode[] {
  const columnCount = tableColumnCount(source);
  if (!columnCount) return rows;
  const sourceMap = rowCellColumns(source);
  const continuationMap = mapRows([...clonedHeader, ...rows]);
  const colWidths = columnWidthsOf(source, columnCount);
  const cellType = schema.nodes.tableCell;
  const paraType = schema.nodes.paragraph;
  const makePlaceholder = (col: number): PmNode => {
    const attrs = colWidths[col] > 0 ? { colwidth: [colWidths[col]] } : {};
    return cellType.create(attrs, paraType.create());
  };
  return rows.map((row) => {
    let span = 0;
    row.forEach((c) => {
      span += (c.attrs.colspan as number) ?? 1;
    });
    if (span >= columnCount) return row;
    const srcPos = sourceMap.get(row)?.[0]?.colStart ?? 0;
    const contPos = continuationMap.get(row)?.[0]?.colStart ?? 0;
    if (contPos >= srcPos) return row; // owner is on this page — ProseMirror already placed it
    const cells: PmNode[] = [];
    for (let col = contPos; col < srcPos; col++) cells.push(makePlaceholder(col));
    row.forEach((c) => cells.push(c));
    return row.type.create(row.attrs, cells, row.marks);
  });
}

/** Build a table node from regrouped rows. Continuation pages (the table
 *  already placed on a prior page) clone the header rows from the group's
 *  first segment (`headerSource` — later segments have no header yet); split
 *  tables carry a `splitGroup` id so unwrapPages merges them back on export. */
function buildTableNode(
  schema: PmNode["type"]["schema"],
  table: PmNode,
  headerSource: PmNode,
  rows: PmNode[],
  isContinuation: boolean,
  splitGroup: string | null,
): PmNode {
  const cloned = isContinuation && splitGroup ? cloneHeaderRows(headerSource) : [];
  const dataRows =
    isContinuation && splitGroup
      ? rebuildContinuationDataRows(rows, cloned, headerSource, schema)
      : rows;
  let finalRows = isContinuation && splitGroup ? [...cloned, ...dataRows] : dataRows;
  if (isContinuation && splitGroup && cloned.length && dataRows.length) {
    // Sync the cloned header's cell colwidths to the column widths carried by
    // the data rows on this continuation page. prosemirror-tables' tableEditing
    // plugin runs fixTables on every transaction: it derives each column's
    // width from its cells and, if they disagree, rewrites every cell in that
    // column to the majority value. The cloned header carries the ORIGINAL
    // header row's tcW colwidths (e.g. [60]), but the data rows here can carry
    // different tcW colwidths (e.g. [72]) — DOCX rows have independent tcW —
    // so fixTables rewrote the clone 60→72 after each re-flow, the next re-flow
    // rebuilt it back to 60, and pagesEqual never held (the re-flow looped
    // forever on large multi-section documents). Aligning the clone to the data
    // rows' column widths makes the column self-consistent, so fixTables leaves
    // it alone → idempotent re-flow.
    const colWidths: number[] = [];
    for (const row of dataRows) {
      let col = 0;
      row.forEach((cell) => {
        const cs = (cell.attrs.colspan as number) ?? 1;
        const cw = (cell.attrs.colwidth as number[] | null) ?? [];
        for (let j = 0; j < cs; j++) {
          if (cw[j] != null && colWidths[col + j] == null) colWidths[col + j] = cw[j] as number;
        }
        col += cs;
      });
    }
    if (colWidths.some((w) => w > 0)) {
      finalRows = finalRows.map((row) => {
        if ((row.attrs as { splitClone?: boolean }).splitClone !== true) return row;
        const cells: PmNode[] = [];
        let col = 0;
        row.forEach((cell) => {
          const cs = (cell.attrs.colspan as number) ?? 1;
          const cw: number[] = [];
          for (let j = 0; j < cs; j++) cw.push(colWidths[col + j] ?? 0);
          cells.push(cell.type.create({ ...cell.attrs, colwidth: cw }, cell.content, cell.marks));
          col += cs;
        });
        return row.type.create(row.attrs, cells, row.marks);
      });
    }
  }
  // Always (re)set splitGroup — clear stale ids when a table now fits one page
  // (it may carry a group id from a prior, larger split).
  const attrs = { ...table.attrs, splitGroup: splitGroup ?? null };
  return schema.nodes.table.create(attrs, finalRows);
}

/** Structural equality of the page sequence (deep, incl. attrs) — gates the
 *  skip-dispatch that terminates the re-flow cycle. */
function pagesEqual(doc: PmNode, newPages: PmNode[]): boolean {
  const current: PmNode[] = [];
  doc.forEach((pn) => {
    if (pn.type.name === "page") current.push(pn);
  });
  if (current.length !== newPages.length) return false;
  return current.every((p, i) => p.eq(newPages[i]));
}

function reflow(editor: Editor, scroll = false): void {
  const view = editor.view;
  const state = view.state;
  // Skip re-flow while a table CellSelection is active: the full-doc rebuild
  // (replaceWith over the whole doc) can't reliably map a multi-cell selection,
  // so re-flowing would drop it — the "selected, then immediately cancelled"
  // symptom. The next pass re-measures once the selection is text again.
  if (!(state.selection instanceof TextSelection)) return;
  const schema = state.schema;
  // DOM-measured page height — the fallback for pages with no section geometry
  // (default editor before a page setup is applied). Pages carrying their
  // section's geometry use the deterministic sectionContentDims height, which
  // matches the page's inline width/height/padding box exactly (so multi-section
  // docs — e.g. a landscape section — pack against each section's own height).
  const domFallback = resolvePageContentHeight(editor);
  // Sub-pixel safety: Pretext's canvas measurement can drift a couple px below
  // the browser's layout (font hinting, CJK kinsoku); over many blocks/page that
  // accumulates past the fixed page box. Closing the page a hair early keeps the
  // marginal block for the next page instead of clipping it.
  const OVERFLOW_SAFETY_PX = 2;
  const segHeightOf = (section: unknown): number =>
    sectionContentDims(section)?.height ?? domFallback;

  // Image decode no longer gates re-flow: pagination measures image paragraphs
  // from node.attrs (clamped to the content width), NOT the live <img> DOM, so
  // a decoding image lands in the exact box already reserved — no height drift,
  // no page jump. (Sized images' loads are layout no-ops; unsized images still
  // converge via onMediaReady in addProseMirrorPlugins.)
  const items = measureFlatItems(editor);
  if (items.length === 0) return;

  // Per-item section properties (each page renders its section's geometry).
  // A paragraph carrying sectionProperties is its section's LAST paragraph
  // (OOXML sectPr in its pPr) and describes the section ENDING there, so its
  // section applies to itself and everything BEFORE it — propagate backwards.
  const docSectionProps =
    (state.doc.attrs as { sectionProperties?: unknown }).sectionProperties ?? null;
  const itemSection: unknown[] = [];
  let sectCursor: unknown = docSectionProps;
  for (let i = items.length - 1; i >= 0; i--) {
    const it = items[i];
    // A heading can be a section's last paragraph too (heading IS a paragraph in
    // OOXML) — read its sectionProperties the same way.
    if (
      it.kind === "block" &&
      (it.node.type.name === "paragraph" || it.node.type.name === "heading")
    ) {
      const sp = (it.node.attrs as { sectionProperties?: unknown }).sectionProperties;
      if (sp != null) sectCursor = sp;
    }
    itemSection[i] = sectCursor;
  }

  const ids = new Map<PmNode, string>();

  // Greedy pack into page segments. An item taller than the page takes a page
  // to itself (a row never splits) rather than loop forever — Word likewise
  // clips oversized content; warn so it is visible.
  const segments: { items: FlatItem[]; section: unknown }[] = [];
  let cur: FlatItem[] = [];
  let curSection: unknown = itemSection[0] ?? docSectionProps;
  // Guard: if the first page's height can't be measured yet (no section geometry
  // AND the DOM not laid out), wait — packing against 0 stacks every block on
  // its own page. Once geometry is present this is deterministic and > 0.
  if (segHeightOf(curSection) <= 0) return;
  let curHeight = segHeightOf(curSection);
  let acc = 0;
  const closeSegment = (): void => {
    if (cur.length > 0) segments.push({ items: cur, section: curSection });
    cur = [];
    acc = 0;
  };
  /** Pop cur's trailing run of keepNext blocks so they follow the block that's
   *  leaving this page onto the next one. MS Office keepNext: a keepNext
   *  paragraph stays with the following paragraph, so when that paragraph
   *  moves to the next page the keepNext run moves with it (a heading no
   *  longer strands alone at the page bottom). Returns [] when the run is
   *  empty, or when it spans all of cur — a keepNext chain that already fills
   *  the page has nothing before it to leave behind, and carrying it would
   *  re-overflow the next page and loop; per the spec a chain too tall for a
   *  page starts that page and flows on (this bound is also the loop guard). */
  const carryKeepNextChain = (): FlatItem[] => {
    let chainLen = 0;
    while (chainLen < cur.length) {
      const last = cur[cur.length - 1 - chainLen];
      if (last.kind === "block" && last.pag.keepNext) chainLen++;
      else break;
    }
    if (chainLen === 0 || chainLen >= cur.length) return [];
    return cur.splice(cur.length - chainLen, chainLen);
  };
  /** Close the current page, carrying its trailing keepNext run onto a fresh
   *  page under `section` so those blocks stay with the block that triggered
   *  the break. */
  const closeSegmentCarrying = (section: unknown): void => {
    const carried = carryKeepNextChain();
    const carriedAcc = carried.reduce((s, c) => s + c.height, 0);
    acc -= carriedAcc;
    closeSegment();
    curSection = section;
    curHeight = segHeightOf(curSection);
    for (const c of carried) cur.push(c);
    acc = carriedAcc;
  };
  // while-loop (not for): a mid-paragraph split inserts the tail back into the
  // flow as a new item, so the index advances past the head only.
  let splitSeq = 0;
  let i = 0;
  while (i < items.length) {
    const item = items[i];
    const pag = item.kind === "block" ? item.pag : undefined;
    // A block whose pageBreak atom leads (imported break-at-start) OR whose
    // pageBreakBefore property is set starts a new page — close the current
    // page before it, matching Word. The curSection/curHeight reset is handled
    // by the `cur.length === 0` block below.
    if (item.kind === "block" && cur.length > 0 && forcesPageBreakBefore(item.node, pag)) {
      closeSegment();
    }
    if (cur.length === 0) {
      curSection = itemSection[i];
      curHeight = segHeightOf(curSection);
    }
    // Core height excludes a block's trailing margin: if it lands as a page's
    // last item, that margin is clipped at the fixed page bottom (Word clips
    // the final paragraph's spacing.after the same way), so it must not be the
    // reason the block overflows to the next page. The full height (incl.
    // after) still accumulates, since a non-final block's after renders as gap
    // before the next block and must count toward the page fill.
    // A pageBreak carrier (an empty paragraph whose only content is the break
    // atom) takes no page space: Word fires the break at the carrier and starts
    // the next content on a fresh page without the carrier occupying a line.
    // Counting its empty-line height toward overflow stranded the carrier alone
    // on its own page, pushing every break's following content one page late.
    const isBreakCarrier =
      item.kind === "block" &&
      !hasSubstantialContent(item.node) &&
      pageBreakPosition(item.node) === "after";
    const core = isBreakCarrier
      ? 0
      : item.kind === "block"
        ? item.height - item.after
        : item.height;
    // Overflow check. `core > curHeight` (the item alone is taller than a page)
    // applies even on an EMPTY page — otherwise the very first item, when it is a
    // splittable multi-image paragraph, takes the whole page to itself and never
    // splits (it never re-enters the `cur.length > 0` branch). The `acc + core`
    // part only matters once the page already has content. trySplitBlock returns
    // null for unsplittable blocks (container/keepLines) → they still take a page
    // to themselves (Word clips oversized content). Rows can't split (a <tr>
    // can't visually break), so they always fall through.
    const overflow =
      cur.length > 0
        ? core > curHeight || acc + core > curHeight - OVERFLOW_SAFETY_PX
        : core > curHeight;
    if (overflow) {
      if (item.kind === "block") {
        const remaining = cur.length > 0 ? curHeight - acc - OVERFLOW_SAFETY_PX : curHeight;
        // Reuse the paragraph's existing splitGroup when it has one: a tail that
        // re-splits mid-packer keeps its own group, so EVERY piece of the same
        // original paragraph shares one id. Otherwise head(p0) + a re-split tail's
        // head'(p1) can't merge back next pass and the doc oscillates between two
        // splitGroup layouts forever — a >2-page multi-image paragraph never
        // converged. Mint a new id only for a paragraph with no group yet.
        const sg =
          (item.node.attrs as { splitGroup?: string | null }).splitGroup ?? `p${splitSeq++}`;
        // keepNext: the previous item is a keepNext heading (Word's heading
        // default) → this paragraph must stay on the heading's page; pass it so
        // trySplitBlock relaxes the orphan rule and splits instead of moving the
        // whole paragraph off the heading's page.
        const prev = cur.length > 0 ? cur[cur.length - 1] : null;
        const prevKeepNext = prev?.kind === "block" && prev.pag.keepNext;
        const split = trySplitBlock(item, remaining, sg, prevKeepNext);
        if (split) {
          // If the head's own height exceeds the current page's remaining space
          // (its first row alone is over-tall — e.g. a near-page-height image
          // arriving after other content), close the current page FIRST so the
          // head starts a fresh page instead of overflowing this one. The tail
          // still spills onto the next page via the closeSegment below.
          if (cur.length > 0 && split.head.height > remaining) closeSegmentCarrying(itemSection[i]);
          cur.push(split.head);
          acc += split.head.height;
          closeSegment();
          // The tail becomes the next item with the SAME splitGroup + section,
          // so the next iteration re-packs it on a fresh page. It is re-measured
          // next re-flow after head+tail merge back into the original paragraph
          // (idempotent). Splice itemSection too so the tail keeps the section.
          items.splice(i + 1, 0, split.tail);
          itemSection.splice(i + 1, 0, itemSection[i]);
          i++;
          continue;
        }
      }
      if (cur.length > 0) {
        closeSegmentCarrying(itemSection[i]);
      }
    }
    cur.push(item);
    acc += item.height;
    if (item.kind === "row" && item.height > curHeight) {
      console.warn(
        "[docen] table row taller than the page content area; clipped (Word would split mid-row, which contenteditable cannot).",
      );
    }
    if (item.kind === "block" && forcesPageBreakAfter(item.node)) {
      closeSegment();
    }
    i++;
  }
  closeSegment();

  // Pass 1: tables whose rows land on >1 page get a splitGroup id (for merge
  // on export and header cloning on continuation pages).
  const tableSegCount = new Map<string, number>();
  for (const seg of segments) {
    const keys = new Set<string>();
    for (const it of seg.items) if (it.kind === "row") keys.add(tableKeyOf(it.table, ids));
    for (const k of keys) tableSegCount.set(k, (tableSegCount.get(k) ?? 0) + 1);
  }
  let groupSeq = 0;
  const splitGroupForKey = new Map<string, string>();
  for (const [key, count] of tableSegCount) {
    if (count > 1) splitGroupForKey.set(key, `t${groupSeq++}`);
  }

  // Pass 2: build page nodes. Contiguous same-table rows merge into one table
  // node per page; continuation pages clone the header rows.
  const seenTables = new Set<string>();
  const firstTableOfGroup = new Map<string, PmNode>();
  const newPages = segments.map((seg) => {
    const children: PmNode[] = [];
    let pendingRows: PmNode[] = [];
    let pendingTable: PmNode | null = null;
    let pendingKey = "";
    const flush = (): void => {
      if (pendingRows.length && pendingTable) {
        const key = tableKeyOf(pendingTable, ids);
        const group = splitGroupForKey.get(key) ?? null;
        const isContinuation = seenTables.has(key);
        // The group's first segment carries the original header rows; clone
        // from it on continuation pages (later segments have no header yet).
        if (group && !firstTableOfGroup.has(group)) firstTableOfGroup.set(group, pendingTable);
        const headerSource = group ? (firstTableOfGroup.get(group) ?? pendingTable) : pendingTable;
        seenTables.add(key);
        children.push(
          buildTableNode(schema, pendingTable, headerSource, pendingRows, isContinuation, group),
        );
      }
      pendingRows = [];
      pendingTable = null;
      pendingKey = "";
    };
    for (const item of seg.items) {
      if (item.kind === "row") {
        const key = tableKeyOf(item.table, ids);
        if (key !== pendingKey) {
          flush();
          pendingTable = item.table;
          pendingKey = key;
        }
        // Clone rows reserve height (measured above) but are re-cloned by
        // buildTableNode — don't re-pack them or the header doubles up.
        if (!item.isClone) pendingRows.push(item.row);
      } else {
        flush();
        children.push(item.node);
      }
    }
    flush();
    return schema.nodes.page.create({ sectionProperties: seg.section ?? null }, children);
  });

  // Idempotence gate: if a re-flow reproduces the current page sequence exactly
  // (deep, incl. attrs), there is nothing to dispatch — this is what terminates
  // the re-flow cycle. Continuation-page clone headers must be rebuilt with
  // column widths self-consistent with their data rows (see buildTableNode), or
  // prosemirror-tables' fixTables rewrites their colwidth after every dispatch
  // and this equality never holds (the re-flow loops forever).
  if (pagesEqual(state.doc, newPages)) return;

  const savedSel = saveSelection(state.doc, state.selection);
  const tr = state.tr.replaceWith(0, state.doc.content.size, newPages);
  const sel = restoreSelection(tr.doc, savedSel);
  if (sel) tr.setSelection(sel);
  tr.setMeta(flowKey, { flow: true });
  tr.setMeta("addToHistory", false);
  view.dispatch(tr);
  // Follow the caret (editor-driven reflows only — see pendingScroll). A reflow
  // can move the caret's block to a new page; scroll it to the TOP of the
  // viewport (ProseMirror's scrollIntoView parks the caret at the bottom edge,
  // which reads wrong for a page jump). No-op if the caret is still in view.
  if (scroll) scrollCaretToTop(editor.view);
}

/**
 * PagePlugin — C-route pagination over a single contenteditable.
 *
 * The document schema is `doc > page+`; each `page` is a fixed-height box
 * (`height` + `overflow: hidden`). This plugin regroups blocks across pages so
 * nothing overflows a page's fixed box: when a page's measured content exceeds
 * its height, trailing blocks move to the next page (and a new page is added at
 * the end when needed). Re-flow is debounced (an offline pass, not per
 * keystroke) and the caret is preserved across regrouping.
 *
 * See CLAUDE.md → Pagination Architecture and CONTRIBUTING.md → Pagination
 * Conventions.
 */
export const PagePlugin = Extension.create<PagePluginOptions>({
  name: "docenPagePlugin",

  addOptions() {
    return { debounceMs: 300 };
  },

  addStorage() {
    return { repaginate: () => {} };
  },

  addProseMirrorPlugins() {
    const editor = this.editor;
    const debounceMs = this.options.debounceMs ?? 300;
    let timer: ReturnType<typeof setTimeout> | undefined;
    // Whether the pending reflow should scroll the caret back into view. OR'd
    // across every schedule() inside a debounce window, so multiple edits in a
    // window still follow the caret.
    let pendingScroll = false;

    const runRaf = (): void => {
      requestAnimationFrame(() => {
        const scroll = pendingScroll;
        pendingScroll = false;
        if (!editor.isDestroyed) reflow(editor, scroll);
      });
    };
    const schedule = (followCaret = false): void => {
      pendingScroll = pendingScroll || followCaret;
      if (timer) clearTimeout(timer);
      timer = setTimeout(runRaf, debounceMs);
    };

    // Host hook: re-flow now, once layout has settled (the host awaits fonts/
    // images and one rAF before calling this after an import or geometry change).
    pageStorageOf(editor).repaginate = (): void => {
      if (timer) clearTimeout(timer);
      runRaf();
    };

    return [
      new Plugin({
        key: flowKey,
        view() {
          return {
            update(view, prevState) {
              // Only re-flow on real doc changes; a no-op or our own flow tr
              // (doc unchanged after our dispatch converges, or the same-check
              // inside reflow short-circuits) does not cascade.
              if (view.state.doc === prevState.doc) return;
              // Follow the caret only when the selection actually moved (a user
              // edit). Programmatic doc changes (image load embedding a data URL,
              // import) keep the caret put — scrolling back to it then fights the
              // user's scroll position: a large doc loaded with the caret on
              // page 1 snaps back to page 1 on every image-load transaction while
              // the user scrolls. scrollCaretToTop is a no-op when the caret is
              // in view, so this never fights normal typing. Defer past PM's
              // scrollIntoView (it runs AFTER plugin view.update, so a
              // synchronous scroll here gets overwritten).
              const caretMoved = view.state.selection.head !== prevState.selection.head;
              if (caretMoved) requestAnimationFrame(() => scrollCaretToTop(view));
              schedule(caretMoved);
            },
            destroy() {
              if (timer) clearTimeout(timer);
            },
          };
        },
      }),
    ];
  },
});
