/**
 * The pagination feedback pass — resolve field atoms against the pages they
 * actually landed on. Runs after every layout: the walk visits each page's
 * flow items (footnote areas included), reads the field descriptor carried by
 * the laid inline atoms, and writes the resolved value back:
 *
 * - `resolved` always holds the live value (paint prefers it over the cache).
 * - `text` is rewritten for non-numbering fields so the measurement follows
 *   the value (Word measures the field result). A rewrite reports
 *   `textChanged`, and the orchestrator re-lays the same projected blocks —
 *   bounded at a few passes; dynamic PAGE/NUMPAGES atoms keep their measuring
 *   placeholder (their value is a function of the very pagination a rewrite
 *   would move).
 *
 * Edges walk paragraphs/table cells/groups; an atom seen twice on a page (a
 * split slice shares its inline array) resolves once.
 */

import {
  computePageNumberOffsets,
  type FlowPage,
  type LaidOutBlock,
  type ProjectedPageNumbering,
} from "@docen/layout";

import {
  evaluateField,
  instructionName,
  LIVE_FIELD_NAMES,
  type FieldContext,
  type FieldFrame,
} from "./fields";

/** One page slice of the pagination result handed to the resolver. */
export interface PageFieldContext {
  /** 0-based physical page index. */
  pageIndex: number;
  /** 1-based displayed page number (the section's w:pgNumType restart applied). */
  pageNumber: number;
  /** Total pages in the document. */
  pageCount: number;
  /** 1-based section number. */
  section: number;
  /** Total pages in this page's section. */
  sectionPages: number;
  /** The section's w:numFmt token (absent = decimal). */
  pageFormat?: string;
}

export interface PageFieldResolution {
  /** Pages whose atoms gained a different `resolved`/`text`. */
  changedPages: number[];
  /** Whether any measured text was rewritten (the caller re-lays). */
  textChanged: boolean;
}

/** One page of the resolve walk as the evaluator's FieldFrame. */
export const frameOfPage = (page: PageFieldContext): FieldFrame => ({
  page: page.pageNumber,
  pageCount: page.pageCount,
  section: page.section,
  sectionPages: page.sectionPages,
  ...(page.pageFormat ? { pageFormat: page.pageFormat } : {}),
});

/** The render pass's resolver policy: only the live numbering fields resolve,
 *  and nothing resolves under the field-code display (Alt+F9) — code atoms
 *  must keep their instruction text so paint shows the code, and the pass must
 *  not churn layouts for a view that is pure presentation. */
export function liveFieldResolver(
  base: Omit<FieldContext, "frame" | "sequences">,
  showFieldCodes: boolean,
): (instruction: string, page: PageFieldContext) => string | null {
  if (showFieldCodes) return () => null;
  return (instruction, page) =>
    LIVE_FIELD_NAMES.has(instructionName(instruction))
      ? evaluateField(instruction, { ...base, frame: frameOfPage(page) })
      : null;
}

export interface BoundedPageFieldResolution {
  pages: FlowPage[];
  sectionOfPage: number[];
  /** Pages whose painted field values changed in the final pass. */
  dirty?: number[];
  /** Resolve passes run (1 when the first pass was stable). */
  passes: number;
}

/** Drive the bounded resolve/re-lay loop: resolve, and when a resolution
 *  rewrote measured text, `relayout` the same projected blocks and resolve
 *  again — at most `maxPasses` times, so a value/pagination oscillation
 *  terminates (the last resolution wins). Numbering fields whose atoms keep
 *  their measuring placeholder never report `textChanged`, so the common case
 *  is a single pass. */
export function resolvePageFieldsBounded(
  pages: FlowPage[],
  sections: readonly { pageNumbering?: ProjectedPageNumbering }[],
  sectionOfPage: number[],
  resolve: (instruction: string, page: PageFieldContext) => string | null | undefined,
  relayout: () => { pages: FlowPage[]; sectionOfPage: number[] },
  maxPasses = 3,
): BoundedPageFieldResolution {
  let currentPages = pages;
  let currentSections = sectionOfPage;
  for (let pass = 1; ; pass++) {
    const outcome = resolvePageFields(currentPages, sections, currentSections, resolve);
    if (!outcome.textChanged || pass >= maxPasses) {
      return {
        pages: currentPages,
        sectionOfPage: currentSections,
        dirty: outcome.changedPages,
        passes: pass,
      };
    }
    const next = relayout();
    currentPages = next.pages;
    currentSections = next.sectionOfPage;
  }
}

/** Resolve every field atom of the laid flow against its page. `resolve`
 *  returns the value or null/undefined when this page/context cannot provide
 *  one (the atom keeps its cache). */
export function resolvePageFields(
  pages: readonly FlowPage[],
  sections: readonly { pageNumbering?: ProjectedPageNumbering }[],
  sectionOfPage: readonly number[],
  resolve: (instruction: string, page: PageFieldContext) => string | null | undefined,
): PageFieldResolution {
  const offsets = computePageNumberOffsets(sections, sectionOfPage);
  const sectionPages = new Map<number, number>();
  for (const section of sectionOfPage)
    sectionPages.set(section, (sectionPages.get(section) ?? 0) + 1);
  const changed = new Set<number>();
  let textChanged = false;

  const visit = (
    block: LaidOutBlock,
    pageIndex: number,
    page: PageFieldContext,
    seen: Set<unknown>,
  ): void => {
    if (block.kind === "paragraph") {
      for (const line of block.lines) {
        for (const item of line.items) {
          const inline = block.inline[item.inlineIndex];
          if (!inline || inline.kind !== "text" || !inline.instruction || seen.has(inline))
            continue;
          seen.add(inline);
          const value = resolve(inline.instruction, page);
          if (value == null) continue;
          if (inline.resolved !== value) {
            inline.resolved = value;
            changed.add(pageIndex);
          }
          // Dynamic page numbers keep their measuring placeholder — their
          // value depends on the pagination a rewrite would move.
          if (inline.field == null && inline.text !== value) {
            inline.text = value;
            textChanged = true;
            changed.add(pageIndex);
          }
        }
      }
      return;
    }
    if (block.kind === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          for (const item of cell.stack) visit(item.block, pageIndex, page, seen);
        }
      }
      return;
    }
    if (block.kind === "group") {
      for (const item of block.children) visit(item.block, pageIndex, page, seen);
    }
  };

  pages.forEach((page, pageIndex) => {
    const section = sectionOfPage[pageIndex] ?? 0;
    const format = sections[section]?.pageNumbering?.format;
    const context: PageFieldContext = {
      pageIndex,
      pageNumber: pageIndex + 1 + (offsets[section] ?? 0),
      pageCount: pages.length,
      section: section + 1,
      sectionPages: sectionPages.get(section) ?? 1,
      ...(format ? { pageFormat: format } : {}),
    };
    const seen = new Set<unknown>();
    for (const item of page.items) visit(item.block, pageIndex, context, seen);
    for (const item of page.footnotes?.items ?? []) visit(item.block, pageIndex, context, seen);
  });

  return { changedPages: [...changed].sort((a, b) => a - b), textChanged };
}
