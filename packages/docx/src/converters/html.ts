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
  return ProseMirrorDOMParser.fromSchema(schema).parse(document.body, options).toJSON();
}
