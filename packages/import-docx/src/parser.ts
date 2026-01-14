import { fromXml } from "xast-util-from-xml";
import { unzipSync } from "fflate";
import type { Root, Element } from "xast";
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

interface ListInfo {
  type: "bullet" | "ordered";
  start?: number;
}
type ListTypeMap = Map<string, ListInfo>;

/**
 * Base64 lookup table for fast encoding
 */
const BASE64_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

/**
 * Convert Uint8Array to base64 string using lookup table and bitwise operations
 * Similar to base64-arraybuffer implementation but without external dependencies
 * Performance: O(n) time complexity, no stack overflow risk
 */
function uint8ArrayToBase64(bytes: Uint8Array): string {
  const len = bytes.length;
  const resultLen = Math.ceil(len / 3) * 4;
  const result = Array.from<string>({ length: resultLen });
  let resultIndex = 0;

  // Process 3 bytes at a time (24 bits -> 4 base64 chars)
  for (let i = 0; i < len; i += 3) {
    // Read 3 bytes (24 bits)
    const byte1 = bytes[i];
    const byte2 = i + 1 < len ? bytes[i + 1] : 0;
    const byte3 = i + 2 < len ? bytes[i + 2] : 0;

    // Extract 4 x 6-bit values using bitwise operations
    const index0 = byte1 >> 2;
    const index1 = ((byte1 & 0x03) << 4) | (byte2 >> 4);
    const index2 = ((byte2 & 0x0f) << 2) | (byte3 >> 6);
    const index3 = byte3 & 0x3f;

    // Encode to base64 characters using lookup table
    result[resultIndex++] = BASE64_CHARS[index0];
    result[resultIndex++] = BASE64_CHARS[index1];
    result[resultIndex++] = i + 1 < len ? BASE64_CHARS[index2] : "=";
    result[resultIndex++] = i + 2 < len ? BASE64_CHARS[index3] : "=";
  }

  return result.join("");
}

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
  // Build abstractNumId -u003e start value mapping
  const abstractNumStarts = new Map<string, number>();
  const numberingXml = files["word/numbering.xml"];
  if (!numberingXml) return listTypeMap;

  const numberingXast = fromXml(new TextDecoder().decode(numberingXml));

  // Build abstractNumId -> numFmt mapping
  const abstractNumFormats = new Map<string, string>();

  if (numberingXast.type === "root") {
    for (const child of numberingXast.children) {
      if (child.type === "element" && child.name === "w:numbering") {
        const numbering = child;

        // First pass: collect all abstractNum definitions
        for (const numChild of numbering.children) {
          if (numChild.type === "element" && numChild.name === "w:abstractNum") {
            const abstractNum = numChild;
            const abstractNumId = abstractNum.attributes["w:abstractNumId"] as string;

            // Find first level and get its format
            for (const lvlChild of abstractNum.children) {
              if (lvlChild.type === "element" && lvlChild.name === "w:lvl") {
                // numFmt is a child element, not an attribute
                for (const fmtChild of lvlChild.children) {
                  if (fmtChild.type === "element" && fmtChild.name === "w:numFmt") {
                    const numFmt = fmtChild.attributes["w:val"] as string;
                    if (numFmt) {
                      abstractNumFormats.set(abstractNumId, numFmt);
                      break;
                    }
                  }
                }
                // Extract start value if present
                for (const startChild of lvlChild.children) {
                  if (startChild.type === "element" && startChild.name === "w:start") {
                    const startVal = startChild.attributes["w:val"] as string;
                    if (startVal) {
                      abstractNumStarts.set(abstractNumId, parseInt(startVal, 10));
                    }
                    break;
                  }
                }
                // Only check first level
                break;
              }
            }
          }
        }

        // Second pass: map numId to list type
        for (const numChild of numbering.children) {
          if (numChild.type === "element" && numChild.name === "w:num") {
            const num = numChild;
            const numId = num.attributes["w:numId"] as string;

            // Find abstractNumId reference
            for (const numChild2 of num.children) {
              if (numChild2.type === "element" && numChild2.name === "w:abstractNumId") {
                const abstractNumId = numChild2.attributes["w:val"] as string;
                const numFmt = abstractNumFormats.get(abstractNumId);

                if (numFmt) {
                  // Determine list type from numFmt
                  // Common formats: bullet, decimal, lowerLetter, upperLetter, lowerRoman, upperRoman
                  const start = abstractNumStarts.get(abstractNumId);

                  if (numFmt === "bullet") {
                    listTypeMap.set(numId, {
                      type: "bullet",
                    });
                  } else {
                    // decimal, letter, and roman formats are all ordered lists
                    listTypeMap.set(numId, {
                      type: "ordered",
                      ...(start !== undefined && { start }),
                    });
                  }
                }
                break;
              }
            }
          }
        }
        break;
      }
    }
  }

  return listTypeMap;
}

/**
 * Extract images from DOCX relationships
 * Returns Map of relationship ID to raw image data (Uint8Array)
 */
function extractImages(files: Record<string, Uint8Array>): Map<string, Uint8Array> {
  const images = new Map<string, Uint8Array>();

  const relsXml = files["word/_rels/document.xml.rels"];
  if (!relsXml) return images;

  const relsXast = fromXml(new TextDecoder().decode(relsXml));

  if (relsXast.type === "root") {
    for (const child of relsXast.children) {
      if (child.type === "element" && child.name === "Relationships") {
        const relationships = child;
        for (const relChild of relationships.children) {
          if (relChild.type === "element" && relChild.name === "Relationship") {
            const rel = relChild;
            const type = rel.attributes.Type;
            const imageRelType =
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

            if (type && type === imageRelType) {
              const rId = rel.attributes.Id;
              const target = rel.attributes.Target;
              if (rId && target) {
                // Extract image from media folder
                const imagePath = "word/" + (target as string);
                const imageData = files[imagePath];
                if (imageData) {
                  images.set(rId as string, imageData);
                }
              }
            }
          }
        }
        break;
      }
    }
  }

  return images;
}

/**
 * Extract hyperlinks from DOCX relationships
 */
function extractHyperlinks(files: Record<string, Uint8Array>): Map<string, string> {
  const hyperlinks = new Map<string, string>();
  const relsXml = files["word/_rels/document.xml.rels"];
  if (!relsXml) return hyperlinks;

  const relsXast = fromXml(new TextDecoder().decode(relsXml));

  // Find Relationships element first (CRITICAL FIX)
  if (relsXast.type === "root") {
    for (const child of relsXast.children) {
      if (child.type === "element" && child.name === "Relationships") {
        const relationships = child;
        // Now iterate through Relationship elements
        for (const relChild of relationships.children) {
          if (relChild.type === "element" && relChild.name === "Relationship") {
            const rel = relChild;
            const type = rel.attributes.Type;
            const hyperlinkRelType =
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
            if (type && type === hyperlinkRelType) {
              const rId = rel.attributes.Id;
              const target = rel.attributes.Target;
              if (rId && target) {
                hyperlinks.set(rId as string, target as string);
              }
            }
          }
        }
        break;
      }
    }
  }
  return hyperlinks;
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

  // Find w:document element
  for (const child of documentXast.children) {
    if (child.type === "element" && child.name === "w:document") {
      const document = child;

      // Find w:body element
      for (const bodyChild of document.children) {
        if (bodyChild.type === "element" && bodyChild.name === "w:body") {
          const body = bodyChild;

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
      }
      break;
    }
  }

  return { type: "doc", content: [] };
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
      result.push(await convertTable(element, hyperlinks, images, options));
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
        const listNodes = await processLists(elements, i, images, hyperlinks, listTypeMap, options);
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
      result.push(await convertParagraph(element, hyperlinks, images, options));
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
  images: Map<string, string>,
  hyperlinks: Map<string, string>,
  listTypeMap: ListTypeMap,
  options?: DocxImportOptions,
): Promise<JSONContent[]> {
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
      const paragraph = await convertParagraph(el, hyperlinks, images, options);
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

  for (const child of element.children) {
    if (child.type !== "element" || child.name !== "w:r") continue;

    const run = child;
    for (const runChild of run.children) {
      if (runChild.type === "element" && runChild.name === "w:t") {
        const textNode = runChild.children.find((c) => c.type === "text");
        if (textNode && "value" in textNode) {
          content.push({
            type: "text",
            text: (textNode as { value: string }).value,
          });
        }
      }
    }
  }

  return content;
}

/**
 * Check if a paragraph is empty (has no text content or images)
 */
function isEmptyParagraph(element: Element): boolean {
  // Check if paragraph has any text runs with content
  for (const child of element.children) {
    if (child.type !== "element" || child.name !== "w:r") continue;

    const run = child;
    for (const runChild of run.children) {
      // Check for text content
      if (runChild.type === "element" && runChild.name === "w:t") {
        const textNode = runChild.children.find((c) => c.type === "text");
        if (textNode && "value" in textNode) {
          const text = (textNode as { value: string }).value;
          // If there's any non-whitespace text, paragraph is not empty
          if (text.trim().length > 0) {
            return false;
          }
        }
      }

      // Check for images (w:drawing, mc:AlternateContent, or w:pict)
      if (runChild.type === "element") {
        if (
          runChild.name === "w:drawing" ||
          runChild.name === "mc:AlternateContent" ||
          runChild.name === "w:pict"
        ) {
          // Paragraph contains an image, not empty
          return false;
        }
      }
    }
  }

  // No text content or images found, paragraph is empty
  return true;
}
