import { fromXml } from "xast-util-from-xml";
import { unzipSync } from "fflate";
import type { Root, Element } from "xast";
import type { JSONContent } from "@tiptap/core";
import type { DocxImportOptions } from "./option";
import type { StyleMap, StyleInfo } from "./parsing/styles";
import type { ListInfo, ListTypeMap, ImageInfo } from "./parsing/types";
import { toUint8Array, DataType } from "undio";
import { findChild } from "./utils/xml";
import { extractImages } from "./parsing/images";
import { extractHyperlinks } from "./parsing/hyperlinks";
import { parseNumberingXml } from "./parsing/numbering";
import { parseStylesXml } from "./parsing/styles";
import { processElements, type ProcessContext } from "./processors";

// Export types for use in converters
export type { StyleMap, StyleInfo, ListInfo, ListTypeMap, ImageInfo };

/**
 * Main entry point: Parse DOCX file and convert to TipTap JSON
 */
export async function parseDOCX(
  input: DataType,
  options: DocxImportOptions = {},
): Promise<JSONContent> {
  const { ignoreEmptyParagraphs = false } = options;

  // Convert input to Uint8Array
  const uint8Array = await toUint8Array(input);
  // Unzip DOCX file
  const files = unzipSync(uint8Array);

  // Extract hyperlinks and images (images are already converted to base64 data URLs)
  const hyperlinks = extractHyperlinks(files);
  const images = extractImages(files);

  // Parse document.xml
  const documentXml = files["word/document.xml"];
  if (!documentXml) {
    throw new Error("Invalid DOCX file: missing word/document.xml");
  }

  const documentXast = fromXml(new TextDecoder().decode(documentXml));

  // Build list type map and style map
  const listTypeMap = parseNumberingXml(files);
  const styleMap = parseStylesXml(files);

  // Convert document
  const content = await convertDocument(
    documentXast,
    images,
    hyperlinks,
    listTypeMap,
    styleMap,
    ignoreEmptyParagraphs,
    options,
  );

  return content;
}

/**
 * Convert document XAST to TipTap JSON
 */
async function convertDocument(
  documentXast: Root,
  images: Map<string, ImageInfo>,
  hyperlinks: Map<string, string>,
  listTypeMap: ListTypeMap,
  styleMap: StyleMap,
  ignoreEmptyParagraphs: boolean,
  options?: DocxImportOptions,
): Promise<JSONContent> {
  if (documentXast.type !== "root") {
    return { type: "doc", content: [] };
  }

  const document = findChild(documentXast, "w:document");
  if (!document) return { type: "doc", content: [] };

  const body = findChild(document, "w:body");
  if (!body) return { type: "doc", content: [] };

  const context: ProcessContext = {
    hyperlinks,
    images,
    listTypeMap,
    styleMap,
    ignoreEmptyParagraphs,
    options,
  };

  const content = await processElements(
    body.children.filter((c): c is Element => c.type === "element"),
    context,
  );

  return {
    type: "doc",
    content,
  };
}
