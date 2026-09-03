// The docx adapter — projects office-open's DocumentOptions into
// @docen/layout's LayoutDoc. The PERSISTENCE model is the projection source
// (not the editor's Tiptap JSON subset): every body shape office-open can
// round-trip reaches the layout engine, and shapes the adapter cannot lay out
// yet (toc, sdt, textbox, altChunk, customXml, rawXml) become placeholder
// boxes instead of silently vanishing. Callers chain
// Tiptap JSON --compileDocument--> DocumentOptions --this--> LayoutDoc.
//
// Zero-DOM discipline is inherited from @docen/layout — this module is
// Node-safe (headless export) by construction.

import type {
  LayoutBlock,
  ProjectedColumns,
  ProjectedFlowBox,
  ProjectedLineNumbers,
  ProjectedPageBackground,
  ProjectedPageBorders,
  ProjectedPageFurniture,
} from "@docen/layout";
import { twipToPx } from "@docen/layout";
import type { DocumentOptions } from "@office-open/docx";

import { indexCharacterStyles } from "../style-cascade";
import type { ProjectContext } from "./project/context";
import { indexNumberings } from "./project/numbering";
import {
  projectChild,
  projectColumns,
  projectFlowBox,
  projectLineNumbers,
  projectPageBackground,
  projectPageBorders,
  projectPageFurniture,
} from "./project/page";

export { projectFlowBox } from "./project/page";

export interface ProjectedSection {
  blocks: LayoutBlock[];
  flow: ProjectedFlowBox;
  furniture: ProjectedPageFurniture;
  /** The section's page borders (w:pgBorders), absent when none. */
  pageBorders?: ProjectedPageBorders;
  /** The section's line numbering (w:lnNumType), absent when none. */
  lineNumbers?: ProjectedLineNumbers;
  /** The section's columns (w:cols), absent for a single-column section. */
  columns?: ProjectedColumns;
}

/** Project a full DocumentOptions into the engine's input: one
 *  {@link ProjectedSection} per document section plus the page background
 *  (document-wide). Sections paginate in order — see
 *  `layoutFlowSections` in @docen/layout. */
export function projectDocumentOptions(doc: DocumentOptions): {
  sections: ProjectedSection[];
  background?: ProjectedPageBackground;
} {
  const ctx: ProjectContext = {
    styles: doc.styles,
    characterStyles: indexCharacterStyles(doc.styles),
    numberings: indexNumberings(doc.numbering),
    listCounters: new Map(),
    openComments: new Set(),
    footnoteOrdinals: new Map(),
    endnoteOrdinals: new Map(),
    // The document-wide tab grid (w:defaultTabStop, twips); Word's 720 default
    // applies when settings omit it (the engine carries that fallback).
    defaultTabStopPx:
      doc.settings?.defaultTabStop != null && doc.settings.defaultTabStop > 0
        ? twipToPx(doc.settings.defaultTabStop)
        : undefined,
  };
  const sections: ProjectedSection[] = (doc.sections ?? []).map((section, i) => {
    const blocks: LayoutBlock[] = [];
    for (const child of section.children ?? []) {
      const block = projectChild(child, ctx);
      if (Array.isArray(block)) blocks.push(...block);
      else if (block) blocks.push(block);
    }
    // A non-final section's last paragraph carries the sectPr — Word paints
    // its mark row as "─────分节符(下一页)─────". The final section's sectPr
    // rides the body's end (no paragraph holds it) and shows no mark.
    const last = blocks[blocks.length - 1];
    if (i < (doc.sections?.length ?? 0) - 1 && last?.kind === "paragraph") last.sectionEnd = true;
    return {
      blocks,
      flow: projectFlowBox(section.properties),
      furniture: projectPageFurniture(section, doc),
      pageBorders: projectPageBorders(section.properties),
      lineNumbers: projectLineNumbers(section.properties),
      columns: projectColumns(section.properties),
    };
  });
  return {
    sections:
      sections.length > 0
        ? sections
        : [
            {
              blocks: [],
              flow: projectFlowBox(undefined),
              furniture: projectPageFurniture(undefined, doc),
              pageBorders: undefined,
              lineNumbers: undefined,
              columns: undefined,
            },
          ],
    background: projectPageBackground(doc),
  };
}
