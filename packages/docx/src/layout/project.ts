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
import type { DocumentOptions, SectionPropertiesOptions } from "@office-open/docx";

import { indexCharacterStyles } from "../style-cascade";
import type { ProjectContext } from "./project/context";
import { isRecord, type BodyParagraph } from "./project/guards";
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
import { projectParagraph } from "./project/paragraph";
import { projectTable } from "./project/table";

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
  /** The section break type (sectPr @w:type): a "continuous" section merges
   *  onto the previous section's flow instead of opening a fresh page. */
  type?: SectionPropertiesOptions["type"];
  /** Footnote id → definition blocks (absent when document has no footnotes). */
  footnoteDefinitions?: Map<number, readonly LayoutBlock[]>;
  /** Endnote id → definition blocks (absent when document has no endnotes). */
  endnoteDefinitions?: Map<number, readonly LayoutBlock[]>;
}

function projectNoteBlocks(
  children: readonly unknown[],
  ctx: ProjectContext,
  defaultStyle: string,
): LayoutBlock[] {
  const blocks: LayoutBlock[] = [];
  for (const child of children) {
    if (typeof child === "string") {
      blocks.push(projectParagraph({ text: child, style: defaultStyle }, ctx));
    } else if (isRecord(child)) {
      if ("paragraph" in child) {
        const p = child.paragraph;
        const pObj =
          typeof p === "string"
            ? { text: p, style: defaultStyle }
            : isRecord(p) && !p.style
              ? { ...p, style: defaultStyle }
              : p;
        blocks.push(projectParagraph(pObj as BodyParagraph, ctx));
      } else if ("table" in child) {
        const t = projectTable(child.table as never, ctx);
        if (t) blocks.push(t);
      } else {
        const pObj = !child.style ? { ...child, style: defaultStyle } : child;
        blocks.push(projectParagraph(pObj as BodyParagraph, ctx));
      }
    }
  }
  return blocks;
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
  const sectionBlocks = (doc.sections ?? []).map((section) => {
    const blocks: LayoutBlock[] = [];
    for (const child of section.children ?? []) {
      const block = projectChild(child, ctx);
      if (Array.isArray(block)) blocks.push(...block);
      else if (block) blocks.push(block);
    }
    return blocks;
  });

  const footnoteDefinitions = new Map<number, readonly LayoutBlock[]>();
  for (const note of doc.footnotes ?? []) {
    if (note.id == null) continue;
    const ordinal = ctx.footnoteOrdinals.get(note.id) ?? note.id;
    const noteCtx: ProjectContext = { ...ctx, currentNoteOrdinal: ordinal };
    const noteBlocks = projectNoteBlocks(note.children ?? [], noteCtx, "FootnoteText");
    footnoteDefinitions.set(note.id, noteBlocks);
  }

  const endnoteDefinitions = new Map<number, readonly LayoutBlock[]>();
  const docEndnotes = (
    doc as unknown as { endnotes?: Array<{ id?: number; children?: unknown[] }> }
  ).endnotes;
  for (const note of docEndnotes ?? []) {
    if (note.id == null) continue;
    const ordinal = ctx.endnoteOrdinals.get(note.id) ?? note.id;
    const noteCtx: ProjectContext = { ...ctx, currentNoteOrdinal: ordinal };
    const noteBlocks = projectNoteBlocks(note.children ?? [], noteCtx, "EndnoteText");
    endnoteDefinitions.set(note.id, noteBlocks);
  }

  const fnDefs = footnoteDefinitions.size > 0 ? footnoteDefinitions : undefined;
  const enDefs = endnoteDefinitions.size > 0 ? endnoteDefinitions : undefined;

  const sections: ProjectedSection[] = (doc.sections ?? []).map((section, i) => {
    const blocks = sectionBlocks[i] ?? [];
    // A non-final section's last paragraph carries the sectPr — Word paints
    // its mark row as "─────分节符(下一页)─────". The final section's sectPr
    // rides the body's end (no paragraph holds it) and shows no mark.
    const last = blocks[blocks.length - 1];
    if (i < (doc.sections?.length ?? 0) - 1 && last?.kind === "paragraph") {
      // The mark row names the break type (Word: "分节符(连续)") — nextPage
      // collapses to true, the painter's default label.
      const type = section.properties?.type;
      last.sectionEnd =
        type === "continuous" || type === "evenPage" || type === "oddPage" ? type : true;
    }
    return {
      blocks,
      flow: {
        ...projectFlowBox(section.properties),
        // settings.xml compat: cell lines join the section's grid only when
        // the document declares w:adjustLineHeightInTable.
        adjustLinesInTable:
          typeof doc.settings?.compatibility === "object" &&
          doc.settings.compatibility.adjustLineHeightInTable === true,
      },
      furniture: projectPageFurniture(section, doc),
      pageBorders: projectPageBorders(section.properties),
      lineNumbers: projectLineNumbers(section.properties),
      columns: projectColumns(section.properties),
      type: section.properties?.type,
      footnoteDefinitions: fnDefs,
      endnoteDefinitions: enDefs,
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
