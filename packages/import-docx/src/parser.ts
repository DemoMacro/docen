import { fromXml } from "xast-util-from-xml";
import { unzipSync } from "fflate";
import type { Root, Element, Text } from "xast";
import type { JSONContent } from "@tiptap/core";
import type { DocxImportOptions } from "./option";
import type { DocxImageConverter, DocxImageInfo } from "./types";
import { toUint8Array, DataType } from "undio";
import { imageMeta } from "image-meta";
import {
  convertParagraph,
  convertTable,
  isListItem,
  getListInfo,
  isCodeBlock,
  getCodeBlockLanguage,
  isHorizontalRule,
  isTaskItem,
  convertTaskItem,
} from "./converters";
import { uint8ArrayToBase64 } from "./utils/base64";
import { findChild, findDeepChildren } from "./utils/xml";
import { extractImages } from "./parsing/images";
import { extractHyperlinks } from "./parsing/hyperlinks";

interface ListInfo {
  type: "bullet" | "ordered";
  start?: number;
}
type ListTypeMap = Map<string, ListInfo>;

/**
 * Default image converter implementation
 * Embeds images as base64 data URLs
 * Exported for users who want to reuse or extend this behavior
 */
export const defaultImageConverter: DocxImageConverter = async (image) => {
  const base64 = uint8ArrayToBase64(image.data);
  return {
    src: `data:${image.contentType};base64,${base64}`,
  };
};

/**
 * Main entry point: Parse DOCX file and convert to TipTap JSON
 */
export async function parseDOCX(
  input: DataType,
  options: DocxImportOptions = {},
): Promise<JSONContent> {
  // Apply defaults
  const { convertImage = defaultImageConverter, ignoreEmptyParagraphs = false } = options;

  // Convert input to Uint8Array
  const uint8Array = await toUint8Array(input);
  // Unzip DOCX file
  const files = unzipSync(uint8Array);

  // Extract hyperlinks and images
  const hyperlinks = extractHyperlinks(files);
  const rawImages = extractImages(files);

  // Process images with custom converter
  const convertedImages = new Map<string, string>();
  for (const [rId, imageData] of rawImages.entries()) {
    try {
      // Get actual image type using image-meta
      let contentType: string;
      try {
        const meta = imageMeta(imageData);
        contentType = `image/${meta.type}`;
      } catch {
        // Fallback to png if type detection fails
        contentType = "image/png";
      }

      const imageInfo: DocxImageInfo = {
        id: rId,
        contentType,
        data: imageData,
      };

      const result = await convertImage(imageInfo);
      convertedImages.set(rId, result.src);
    } catch (error) {
      // If image conversion fails, use fallback
      console.warn(`Failed to convert image ${rId}:`, error);
      let fallbackContentType = "image/png";
      try {
        const meta = imageMeta(imageData);
        fallbackContentType = `image/${meta.type}`;
      } catch {
        // Keep default png
      }
      const fallbackBase64 = uint8ArrayToBase64(imageData);
      const fallbackUrl = `data:${fallbackContentType};base64,${fallbackBase64}`;
      convertedImages.set(rId, fallbackUrl);
    }
  }

  // Parse document.xml
  const documentXml = files["word/document.xml"];
  if (!documentXml) {
    throw new Error("Invalid DOCX file: missing word/document.xml");
  }

  const documentXast = fromXml(new TextDecoder().decode(documentXml));

  // Build list type map
  const listTypeMap = parseNumberingXml(files);

  // Convert document
  const content = await convertDocument(
    documentXast,
    convertedImages,
    hyperlinks,
    listTypeMap,
    ignoreEmptyParagraphs,
    options,
  );

  return content;
}

/**
 * Parse numbering.xml to build list type map
 */
function parseNumberingXml(files: Record<string, Uint8Array>): ListTypeMap {
  const listTypeMap = new Map<string, ListInfo>();
  const abstractNumStarts = new Map<string, number>();
  const numberingXml = files["word/numbering.xml"];
  if (!numberingXml) return listTypeMap;

  const numberingXast = fromXml(new TextDecoder().decode(numberingXml));
  const abstractNumFormats = new Map<string, string>();

  const numbering = findChild(numberingXast, "w:numbering");
  if (!numbering) return listTypeMap;

  // First pass: collect all abstractNum definitions
  const abstractNums = findDeepChildren(numbering, "w:abstractNum");
  for (const abstractNum of abstractNums) {
    const abstractNumId = abstractNum.attributes["w:abstractNumId"] as string;
    const lvl = findChild(abstractNum, "w:lvl");
    if (!lvl) continue;

    const numFmt = findChild(lvl, "w:numFmt");
    if (numFmt?.attributes["w:val"]) {
      abstractNumFormats.set(abstractNumId, numFmt.attributes["w:val"] as string);
    }

    const start = findChild(lvl, "w:start");
    if (start?.attributes["w:val"]) {
      abstractNumStarts.set(abstractNumId, parseInt(start.attributes["w:val"] as string, 10));
    }
  }

  // Second pass: map numId to list type
  const nums = findDeepChildren(numbering, "w:num");
  for (const num of nums) {
    const numId = num.attributes["w:numId"] as string;
    const abstractNumId = findChild(num, "w:abstractNumId");
    if (!abstractNumId?.attributes["w:val"]) continue;

    const abstractNumIdVal = abstractNumId.attributes["w:val"] as string;
    const numFmt = abstractNumFormats.get(abstractNumIdVal);
    if (!numFmt) continue;

    const start = abstractNumStarts.get(abstractNumIdVal);

    if (numFmt === "bullet") {
      listTypeMap.set(numId, { type: "bullet" });
    } else {
      listTypeMap.set(numId, {
        type: "ordered",
        ...(start !== undefined && { start }),
      });
    }
  }

  return listTypeMap;
}

/**
 * Convert document XAST to TipTap JSON
 */
async function convertDocument(
  documentXast: Root,
  images: Map<string, string>,
  hyperlinks: Map<string, string>,
  listTypeMap: ListTypeMap,
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

  // Process all elements in body
  const content = await processElements(
    body.children.filter((c) => c.type === "element") as Element[],
    images,
    hyperlinks,
    listTypeMap,
    ignoreEmptyParagraphs,
    options,
  );

  return {
    type: "doc",
    content,
  };
}

/**
 * Process all elements in document body
 */
async function processElements(
  elements: Element[],
  images: Map<string, string>,
  hyperlinks: Map<string, string>,
  listTypeMap: ListTypeMap,
  ignoreEmptyParagraphs: boolean,
  options?: DocxImportOptions,
): Promise<JSONContent[]> {
  const result: JSONContent[] = [];
  let i = 0;

  while (i < elements.length) {
    const element = elements[i];

    // Handle tables
    if (element.name === "w:tbl") {
      result.push(await convertTable(element, { hyperlinks, images, options }));
      i++;
      // Skip empty paragraph after table (export-docx adds these for spacing)
      if (
        i < elements.length &&
        elements[i].name === "w:p" &&
        isEmptyParagraph(elements[i] as Element)
      ) {
        i++;
      }
      continue;
    }

    // Handle paragraphs
    if (element.name === "w:p") {
      // Skip empty paragraphs if option is enabled
      if (ignoreEmptyParagraphs && isEmptyParagraph(element)) {
        i++;
        continue;
      }

      // Check for code block
      if (isCodeBlock(element)) {
        const codeBlockNodes = processCodeBlocks(elements, i);
        result.push(...codeBlockNodes);
        i += codeBlockNodes.length;
        continue;
      }

      // Check for task items (before regular list items)
      if (isTaskItem(element)) {
        const taskListNodes = processTaskLists(elements, i);
        result.push(...taskListNodes);
        i += getTaskListConsumed(elements, i);
        continue;
      }

      // Check for list items
      if (isListItem(element)) {
        const listNodes = await processLists(elements, i, {
          hyperlinks,
          images,
          listTypeMap,
          options,
        });
        result.push(...listNodes);
        i += getListConsumed(elements, i);
        continue;
      }

      // Check for horizontal rule (page break)
      if (isHorizontalRule(element)) {
        result.push({ type: "horizontalRule" });
        i++;
        continue;
      }

      // Regular paragraph
      result.push(await convertParagraph(element, { hyperlinks, images, options }));
      i++;
      continue;
    }

    i++;
  }

  return result;
}

/**
 * Process consecutive code blocks
 */
function processCodeBlocks(elements: Element[], startIndex: number): JSONContent[] {
  const result: JSONContent[] = [];
  let i = startIndex;

  while (i < elements.length) {
    const element = elements[i];
    if (element.name !== "w:p" || !isCodeBlock(element)) {
      break;
    }

    const language = getCodeBlockLanguage(element);
    const codeBlockNode: JSONContent = {
      type: "codeBlock",
      ...(language && { attrs: { language } }),
      content: extractTextFromParagraph(element),
    };

    result.push(codeBlockNode);
    i++;
  }

  return result;
}

/**
 * Process consecutive list items and group into lists
 */
async function processLists(
  elements: Element[],
  startIndex: number,
  params: {
    hyperlinks: Map<string, string>;
    images: Map<string, string>;
    listTypeMap: Map<string, { type: "bullet" | "ordered"; start?: number }>;
    options?: DocxImportOptions;
  },
): Promise<JSONContent[]> {
  const { listTypeMap } = params;
  const result: JSONContent[] = [];
  let i = startIndex;

  while (i < elements.length) {
    const element = elements[i];
    if (element.name !== "w:p" || !isListItem(element)) {
      break;
    }

    const listInfo = getListInfo(element);
    if (!listInfo) {
      break;
    }

    // Get list type from map
    const listTypeInfo = listTypeMap.get(listInfo.numId);
    const listType = listTypeInfo?.type || "bullet";

    // Collect consecutive items with same numId
    const items: JSONContent[] = [];
    while (i < elements.length) {
      const el = elements[i];
      if (el.name !== "w:p" || !isListItem(el)) {
        break;
      }

      const info = getListInfo(el);
      if (!info || info.numId !== listInfo.numId) {
        break;
      }

      // Convert list item
      const paragraph = await convertParagraph(el, params);
      const listItem = {
        type: "listItem",
        content: [paragraph],
      };
      items.push(listItem);
      i++;
    }

    // Create list node
    const listNode: JSONContent = {
      type: listType === "bullet" ? "bulletList" : "orderedList",
      content: items,
    };

    // Add start attribute for ordered lists if available
    if (listType === "ordered") {
      listNode.attrs = {
        type: null,
        ...(listTypeInfo?.start !== undefined && { start: listTypeInfo.start }),
      };
    }

    result.push(listNode);
  }

  return result;
}

/**
 * Get number of elements consumed by a list
 */
function getListConsumed(elements: Element[], startIndex: number): number {
  let count = 0;
  let i = startIndex;

  while (i < elements.length) {
    const element = elements[i];
    if (element.name !== "w:p" || !isListItem(element)) {
      break;
    }

    count++;
    i++;
  }

  return count;
}

/**
 * Process consecutive task lists
 */
function processTaskLists(elements: Element[], startIndex: number): JSONContent[] {
  const items: JSONContent[] = [];
  let i = startIndex;

  while (i < elements.length) {
    const element = elements[i];
    if (element.name !== "w:p" || !isTaskItem(element)) {
      break;
    }

    const taskItem = convertTaskItem(element);
    items.push(taskItem);
    i++;
  }

  // Return taskList wrapper containing all items
  return [
    {
      type: "taskList",
      content: items,
    },
  ];
}

/**
 * Get number of elements consumed by a task list
 */
function getTaskListConsumed(elements: Element[], startIndex: number): number {
  let count = 0;
  let i = startIndex;

  while (i < elements.length) {
    const element = elements[i];
    if (element.name !== "w:p" || !isTaskItem(element)) {
      break;
    }

    count++;
    i++;
  }

  return count;
}

/**
 * Extract text content from a paragraph (for code blocks)
 */
function extractTextFromParagraph(element: Element): Array<{ type: string; text: string }> {
  const content: Array<{ type: string; text: string }> = [];

  const runs = findDeepChildren(element, "w:r");
  for (const run of runs) {
    const textElement = findChild(run, "w:t");
    if (!textElement) continue;

    const textNode = textElement.children.find((c): c is Text => c.type === "text");
    if (textNode && textNode.value) {
      content.push({
        type: "text",
        text: textNode.value,
      });
    }
  }

  return content;
}

/**
 * Check if a paragraph is empty (has no text content or images)
 */
function isEmptyParagraph(element: Element): boolean {
  // Check if paragraph has any text runs with content
  const runs = findDeepChildren(element, "w:r");
  for (const run of runs) {
    // Check for text content
    const textElement = findChild(run, "w:t");
    if (textElement) {
      const textNode = textElement.children.find((c): c is Text => c.type === "text");
      if (textNode && textNode.value && textNode.value.trim().length > 0) {
        return false;
      }
    }

    // Check for images (w:drawing, mc:AlternateContent, or w:pict)
    if (
      findChild(run, "w:drawing") ||
      findChild(run, "mc:AlternateContent") ||
      findChild(run, "w:pict")
    ) {
      return false;
    }
  }

  // No text content or images found, paragraph is empty
  return true;
}
