import { fromXml } from "xast-util-from-xml";
import { unzipSync } from "fflate";
import type { Root, Element } from "xast";
import type { JSONContent } from "@tiptap/core";
import type { DocxImportOptions } from "./options";
import type { StyleMap } from "./parsers/styles";
import type { ListTypeMap, ImageInfo } from "./parsers/types";
import { toUint8Array, DataType } from "undio";
import { findChild, findDeepChildren } from "@docen/utils";
import { extractImages } from "./parsers/images";
import { extractHyperlinks } from "./parsers/hyperlinks";
import { parseNumberingXml } from "./parsers/numbering";
import { parseStylesXml } from "./parsers/styles";
import { convertTable } from "./converters/table";
import { convertParagraph } from "./converters/paragraph";
import { convertTaskItem } from "./converters/task-list";
import { isCodeBlock, getCodeBlockLanguage } from "./converters/code-block";
import { isListItem, getListInfo } from "./converters/list";
import { isTaskItem } from "./converters/task-list";
import { isHorizontalRule } from "./converters/horizontal-rule";

/**
 * Parsing context containing all global resources from DOCX file
 */
export interface ParseContext extends DocxImportOptions {
  hyperlinks: Map<string, string>;
  images: Map<string, ImageInfo>;
  listTypeMap: ListTypeMap;
  styleMap: StyleMap;
}

/**
 * Main entry point: Parse DOCX file and convert to TipTap JSON
 */
export async function parseDOCX(
  input: DataType,
  options: DocxImportOptions = {},
): Promise<JSONContent> {
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

  // Create parsing context
  const context: ParseContext = {
    ...options,
    hyperlinks,
    images,
    listTypeMap,
    styleMap,
  };

  // Convert document
  return await convertDocument(documentXast, { context });
}

/**
 * Convert document XAST to TipTap JSON
 */
async function convertDocument(
  node: Root,
  params: { context: ParseContext },
): Promise<JSONContent> {
  if (node.type !== "root") {
    return { type: "doc", content: [] };
  }

  const document = findChild(node, "w:document");
  if (!document) return { type: "doc", content: [] };

  const body = findChild(document, "w:body");
  if (!body) return { type: "doc", content: [] };

  const content = await convertElements(
    body.children.filter((c): c is Element => c.type === "element"),
    params,
  );

  return {
    type: "doc",
    content,
  };
}

/**
 * Convert XML elements to TipTap nodes (main conversion loop)
 */
async function convertElements(
  elements: Element[],
  params: { context: ParseContext },
): Promise<JSONContent[]> {
  const result: JSONContent[] = [];

  for (let i = 0; i < elements.length; i++) {
    const element = elements[i];

    // Skip empty paragraphs if option is set
    if (
      params.context.ignoreEmptyParagraphs &&
      element.name === "w:p" &&
      isEmptyParagraph(element)
    ) {
      continue;
    }

    const node = await convertElement(element, elements, i, params);

    if (Array.isArray(node)) {
      result.push(...node);
    } else if (node) {
      result.push(node);
    }
  }

  return result;
}

/**
 * Convert single XML element to TipTap node (routing function)
 */
async function convertElement(
  element: Element,
  siblings: Element[],
  index: number,
  params: { context: ParseContext },
): Promise<JSONContent | JSONContent[] | null> {
  switch (element.name) {
    case "w:tbl":
      return await convertTable(element, params);

    case "w:p":
      // Route to different paragraph processors
      if (isCodeBlock(element)) {
        return await convertCodeBlock(element);
      }

      if (isTaskItem(element)) {
        return await convertTaskItem(element, params);
      }

      if (isListItem(element)) {
        return await convertList(element, siblings, index, params);
      }

      if (isHorizontalRule(element)) {
        return { type: "horizontalRule" };
      }

      // Default paragraph
      return await convertParagraph(element, params);

    default:
      return null;
  }
}

/**
 * Convert code block paragraph
 */
async function convertCodeBlock(element: Element): Promise<JSONContent> {
  const language = getCodeBlockLanguage(element);
  const content = extractTextFromParagraph(element);

  return {
    type: "codeBlock",
    ...(language && { attrs: { language } }),
    content,
  };
}

/**
 * Convert list (handles consecutive list items)
 */
async function convertList(
  startElement: Element,
  siblings: Element[],
  startIndex: number,
  params: { context: ParseContext },
): Promise<JSONContent> {
  const listInfo = getListInfo(startElement);
  if (!listInfo) {
    return await convertParagraph(startElement, params);
  }

  const listTypeInfo = params.context.listTypeMap.get(listInfo.numId);
  const listType = listTypeInfo?.type || "bullet";

  // Collect consecutive list items with same numId
  const items: JSONContent[] = [];
  let i = startIndex;

  while (i < siblings.length) {
    const el = siblings[i];
    if (el.name !== "w:p" || !isListItem(el)) {
      break;
    }

    const info = getListInfo(el);
    if (!info || info.numId !== listInfo.numId) {
      break;
    }

    // Convert list item paragraph
    const paragraph = await convertParagraph(el, params);
    const listItemContent = Array.isArray(paragraph) ? paragraph[0] : paragraph;

    items.push({
      type: "listItem",
      content: [listItemContent],
    });

    i++;
  }

  // Build list node
  const listNode: JSONContent = {
    type: listType === "bullet" ? "bulletList" : "orderedList",
    content: items,
  };

  if (listType === "ordered") {
    listNode.attrs = {
      type: null,
      ...(listTypeInfo?.start !== undefined && { start: listTypeInfo.start }),
    };
  }

  return listNode;
}

/**
 * Extract text content from paragraph (for code blocks)
 */
function extractTextFromParagraph(element: Element): Array<{ type: string; text: string }> {
  const content: Array<{ type: string; text: string }> = [];
  const runs = findDeepChildren(element, "w:r");

  for (const run of runs) {
    const textElement = findChild(run, "w:t");
    if (!textElement) continue;

    const textNode = textElement.children.find((c) => c.type === "text");
    if (textNode && "value" in textNode && textNode.value) {
      content.push({
        type: "text",
        text: textNode.value,
      });
    }
  }

  return content;
}

/**
 * Check if a paragraph is empty
 */
function isEmptyParagraph(element: Element): boolean {
  const runs = findDeepChildren(element, "w:r");

  for (const run of runs) {
    const textElement = findChild(run, "w:t");
    if (textElement) {
      const textNode = textElement.children.find((c) => c.type === "text");
      if (textNode && "value" in textNode && textNode.value && textNode.value.trim().length > 0) {
        return false;
      }
    }

    if (
      findChild(run, "w:drawing") ||
      findChild(run, "mc:AlternateContent") ||
      findChild(run, "w:pict")
    ) {
      return false;
    }

    const br = findChild(run, "w:br");
    if (br && br.attributes["w:type"] === "page") {
      return false;
    }
  }

  return true;
}
