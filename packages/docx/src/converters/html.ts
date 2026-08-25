import { encodeBase64 } from "@office-open/core";
import { getSchema } from "@tiptap/core";
import {
  DOMParser as ProseMirrorDOMParser,
  DOMSerializer,
  Node,
  type ParseOptions,
} from "@tiptap/pm/model";
import { parseHTML as createLinkedomDocument } from "linkedom";

import type { Extensions, JSONContent } from "../core";
import { docxExtensions } from "../core";
import { assignOrderedReferences } from "../extensions/list-numbering";
import { sectionLinePitchCss, sectionMarginCss } from "../extensions/utils";

/** Page background — mirrors parseDOCX output (office-open parse.ts): a simple
 *  color background `{ color, themeColor, … }` or a VML pattern `{ rawXml,
 *  rawMedia }`. Document-level (CT_DocumentBase), so it is read off `doc.attrs`
 *  and wraps the whole document. */
interface DocumentBackground {
  color?: string;
  rawMedia?: Array<{
    type?: string;
    data?: Uint8Array | Record<string, number>;
  }>;
}

/** JSON round-trips byte arrays as plain objects ({0:byte,…}); rebuild here. */
function toBytes(data: Uint8Array | Record<string, number> | undefined): Uint8Array | null {
  if (!data) return null;
  return data instanceof Uint8Array ? data : new Uint8Array(Object.values(data));
}

/** Page background → CSS for the root wrapper. Color renders directly; a VML
 *  pattern's first media item tiles as an image. OOXML patterns have no CSS
 *  equivalent, so DOCX (not HTML) is the fidelity source. */
function backgroundToCss(bg: DocumentBackground | undefined): string | undefined {
  const styles: string[] = [];
  if (bg?.color) styles.push(`background-color:#${bg.color}`);
  const media = bg?.rawMedia?.[0];
  const bytes = toBytes(media?.data);
  if (bytes) {
    styles.push(
      `background-image:url(data:image/${media?.type ?? "png"};base64,${encodeBase64(bytes)})`,
      "background-repeat:repeat",
    );
  }
  return styles.length ? styles.join(";") : undefined;
}

/** A section's geometry fields used for CSS (subset of SectionPropertiesOptions). */
type SectionGeometry = { page?: { margin?: unknown }; grid?: unknown } | null;

/** A run of blocks belonging to one section. OOXML attaches sectPr to a
 *  section's LAST paragraph; `properties` is that paragraph's sectionProperties
 *  (or doc.attrs.sectionProperties for the final section). */
interface JsonSection {
  properties: SectionGeometry;
  blocks: JSONContent[];
}

/** Split flat `doc > block+` into sections by section-carrying paragraphs.
 *  Mirrors DocxManager's compile-time split (converters/docx.ts): a paragraph
 *  with `sectionProperties` closes its section; trailing blocks form the final
 *  section under doc.attrs.sectionProperties. No section-carrying paragraph →
 *  a single section (backward compatible). */
function splitJsonSections(doc: JSONContent): JsonSection[] {
  const sections: JsonSection[] = [];
  let current: JSONContent[] = [];
  for (const node of doc.content ?? []) {
    current.push(node);
    const sp = (node.attrs as Record<string, unknown> | undefined)
      ?.sectionProperties as SectionGeometry;
    if (sp != null) {
      sections.push({ properties: sp, blocks: current });
      current = [];
    }
  }
  const tailProps = (doc.attrs as Record<string, unknown> | undefined)
    ?.sectionProperties as SectionGeometry;
  sections.push({ properties: tailProps ?? null, blocks: current });
  return sections;
}

/** Regroup consecutive list paragraphs (serialized as <p data-list …>) into
 *  nested ul/ol lists — the HTML shape every consumer expects. Flat depth
 *  (data-list-level) drives nesting: each level deeper nests inside the
 *  previous item's <li>; a shallower level pops back out. A non-list sibling
 *  ends the group. Mutates `parent` in place. */
function regroupLists(parent: HTMLElement, document: Document): void {
  // Open lists; entry i is the list element at depth i.
  const stack: { el: HTMLElement; kind: string }[] = [];
  const top = () => stack[stack.length - 1];
  const listTag = (kind: string) => (kind === "ordered" ? "ol" : "ul");

  for (const node of [...parent.childNodes]) {
    const p = node as HTMLElement;
    const kind = p.getAttribute?.("data-list");
    if (!kind) {
      stack.length = 0;
      continue;
    }
    const level = Number(p.getAttribute("data-list-level") ?? 0) || 0;
    while (stack.length > level + 1) stack.pop();
    // Open lists down to the paragraph's depth — each nests inside the
    // previous item's <li>. A kind flip at the current depth closes the top
    // list and opens a sibling of the new kind under the same host. A
    // top-level list inserts at the paragraph's own position so later
    // siblings keep their order.
    while (stack.length < level + 1) {
      const host = top()?.el.lastElementChild ?? top()?.el;
      const list = document.createElement(listTag(kind));
      if (host) host.appendChild(list);
      else parent.insertBefore(list, p);
      stack.push({ el: list, kind });
    }
    if (top().kind !== kind) {
      stack.pop();
      const host = top()?.el.lastElementChild ?? top()?.el;
      const list = document.createElement(listTag(kind));
      if (host) host.appendChild(list);
      else parent.insertBefore(list, p);
      stack.push({ el: list, kind });
    }
    p.removeAttribute("data-list");
    p.removeAttribute("data-list-level");
    const li = document.createElement("li");
    li.appendChild(p);
    top().el.appendChild(li);
  }
}

/**
 * Serialize Tiptap JSON to an HTML string. Renders per-section: each OOXML
 * section (CT_SectPr) becomes a `<section>` carrying its own page margin
 * (padding) and a document-grid line-height (the font's `normal` metric —
 * Word does not add the grid pitch to rendered line height), so paragraph
 * line-spacing multiples resolve against the section's font, not a fallback.
 * The document background (CT_DocumentBase, single for the whole doc) wraps all
 * sections.
 *
 * Same ProseMirror DOMSerializer pipeline as @tiptap/html, on a linkedom
 * document: happy-dom drops calc(var(…)) when re-serializing the style
 * attribute, so DOCX line-spacing survives only with linkedom.
 */
export function generateHTML(doc: JSONContent, extensions: Extensions = docxExtensions): string {
  const schema = getSchema(extensions);
  const { document } = createLinkedomDocument("<!DOCTYPE html><html><body></body></html>");

  const serializer = DOMSerializer.fromSchema(schema);
  const parts: string[] = [];
  for (const section of splitJsonSections(doc)) {
    const sp = section.properties;
    const styles: string[] = [];
    const padding = sectionMarginCss(sp?.page?.margin);
    if (padding) styles.push(padding);
    styles.push(...sectionLinePitchCss(sp?.grid));
    const sec = document.createElement("section");
    if (styles.length) sec.setAttribute("style", styles.join(";"));
    if (section.blocks.length) {
      const fragment = Node.fromJSON(schema, {
        type: "doc",
        content: section.blocks,
      }).content;
      serializer.serializeFragment(fragment, { document }, sec);
      regroupLists(sec, document);
    }
    parts.push(sec.outerHTML);
  }
  const body = parts.join("");
  const bgCss = backgroundToCss(doc.attrs?.background);
  return bgCss ? `<div style="${bgCss}">${body}</div>` : body;
}

/**
 * Parse an HTML string into Tiptap JSON. Same ProseMirror DOMParser pipeline as
 * @tiptap/html on a linkedom document. The background wrapper and section
 * containers are unknown elements (no doc/div/section node in the schema), so
 * the parser ignores their tags and extracts the content. Section geometry
 * (linePitch/margins) and the page background are section/doc-level metadata,
 * not content — they round-trip losslessly via DOCX, not HTML.
 */
export function parseHTML(
  html: string,
  extensions: Extensions = docxExtensions,
  options?: ParseOptions,
): JSONContent {
  const schema = getSchema(extensions);
  const { document } = createLinkedomDocument(`<!DOCTYPE html><html><body>${html}</body></html>`);
  const json = ProseMirrorDOMParser.fromSchema(schema).parse(document.body, options).toJSON();
  // <li> items carry the html-ordered placeholder — rewrite each run to a
  // fresh generated reference (see list-numbering).
  return assignOrderedReferences(json);
}
