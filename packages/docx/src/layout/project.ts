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
  ProjectedFlowBox,
  ProjectedPageBackground,
  ProjectedPageBorders,
  ProjectedPageFurniture,
} from "@docen/layout";
import type { DocumentOptions } from "@office-open/docx";

import { indexCharacterStyles } from "../style-cascade";
import type { ProjectContext } from "./project/context";
import { indexNumberings } from "./project/numbering";
import {
  projectChild,
  projectFlowBox,
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
  };
  const sections: ProjectedSection[] = (doc.sections ?? []).map((section) => {
    const blocks: LayoutBlock[] = [];
    for (const child of section.children ?? []) {
      const block = projectChild(child, ctx);
      if (Array.isArray(block)) blocks.push(...block);
      else if (block) blocks.push(block);
    }
    return {
      blocks,
      flow: projectFlowBox(section.properties),
      furniture: projectPageFurniture(section, doc),
      pageBorders: projectPageBorders(section.properties),
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
            },
          ],
    background: projectPageBackground(doc),
  };
}
