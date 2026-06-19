import type { SectionPropertiesOptions } from "@office-open/docx";

import { Node } from "../core";
import type { HeaderFooterSlots } from "../types";

/**
 * SectionBreak — block atom node marking a DOCX section boundary.
 *
 * OOXML sections (sectPr) attach to a section's last paragraph, so this node
 * marks a section END: the blocks before it (back to the previous sectionBreak
 * or doc start) form one section, whose full context this node carries:
 *
 * - `attrs.properties` — page layout (SectionPropertiesOptions: page size/margin/
 *   orientation, columns, section type, grid, …).
 * - `attrs.headers`/`attrs.footers` — that section's header/footer content, each
 *   slot (default/first/even) a JSONContent[] resolved from SectionChild[].
 *
 * The last section's context rides on `doc.attrs.sectionProperties`/
 * `sectionHeaders`/`sectionFooters` (no trailing sectionBreak). Single-section
 * documents have no sectionBreak at all.
 *
 * Not a single-node renderDocx/parseDocx carrier: DocxManager intercepts it in
 * its compile/resolve main loops to split/merge sections.
 */
interface SectionBreakAttrs {
  properties: SectionPropertiesOptions | null;
  headers: HeaderFooterSlots | null;
  footers: HeaderFooterSlots | null;
}

/** Attribute spec for a nested value stored as JSON in a data-* attr. */
const attrDataJson = (name: string) => ({
  default: null,
  rendered: false,
  parseHTML: (element: HTMLElement) => {
    const raw = element.getAttribute(name);
    if (!raw) return null;
    try {
      return JSON.parse(raw);
    } catch {
      return null;
    }
  },
});

export const SectionBreak = Node.create({
  name: "sectionBreak",
  group: "block",
  atom: true,

  addAttributes() {
    return {
      properties: attrDataJson("data-properties"),
      headers: attrDataJson("data-headers"),
      footers: attrDataJson("data-footers"),
    };
  },

  parseHTML() {
    return [{ tag: "div[data-section-break]" }];
  },

  renderHTML({ node }: { node: { attrs: Record<string, unknown> } }) {
    const { properties, headers, footers } = node.attrs as unknown as SectionBreakAttrs;
    const attrs: Record<string, unknown> = {
      "data-section-break": "",
      contenteditable: "false",
    };
    if (properties) attrs["data-properties"] = JSON.stringify(properties);
    if (headers) attrs["data-headers"] = JSON.stringify(headers);
    if (footers) attrs["data-footers"] = JSON.stringify(footers);
    return ["div", attrs];
  },
});
