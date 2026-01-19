import type { Element } from "xast";
import type { JSONContent } from "@tiptap/core";
import type { DocxImportOptions } from "./options";
import type { StyleMap } from "./parsers/styles";
import type { ImageInfo } from "./parsers/types";
import { convertTable } from "./converters/table";
import { convertParagraph } from "./converters/paragraph";
import {
  isCodeBlock,
  getCodeBlockLanguage,
  isHorizontalRule,
  isListItem,
  getListInfo,
  isTaskItem,
} from "./converters";
import { findChild, findDeepChildren } from "./utils/xml";

export interface ProcessContext {
  hyperlinks: Map<string, string>;
  images: Map<string, ImageInfo>;
  listTypeMap: Map<string, { type: "bullet" | "ordered"; start?: number }>;
  styleMap: StyleMap;
  ignoreEmptyParagraphs: boolean;
  options?: DocxImportOptions;
}

export interface ProcessResult {
  nodes: JSONContent[];
  consumed: number;
}

type ElementProcessor = (
  elements: Element[],
  index: number,
  context: ProcessContext,
) => Promise<ProcessResult>;

/**
 * Extract text content from paragraph (for code blocks)
 */

const extractTextFromParagraph = (element: Element): Array<{ type: string; text: string }> => {
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
};

/**
 * Process table element
 */

const processTable: ElementProcessor = async (elements, index, context) => {
  const table = await convertTable(elements[index], {
    hyperlinks: context.hyperlinks,
    images: context.images,
    options: context.options,
    styleMap: context.styleMap,
  });

  let consumed = 1;

  // Skip empty paragraph after table
  if (
    index + 1 < elements.length &&
    elements[index + 1].name === "w:p" &&
    isEmptyParagraph(elements[index + 1] as Element)
  ) {
    consumed++;
  }

  return { nodes: [table], consumed };
};

/**
 * Process consecutive code blocks
 */

const processCodeBlocks: ElementProcessor = async (elements, index) => {
  const result: JSONContent[] = [];
  let i = index;

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

  return { nodes: result, consumed: i - index };
};

/**
 * Process consecutive list items and group into lists
 */

const processLists: ElementProcessor = async (elements, index, context) => {
  const { listTypeMap } = context;
  const result: JSONContent[] = [];
  let i = index;

  while (i < elements.length) {
    const element = elements[i];
    if (element.name !== "w:p" || !isListItem(element)) {
      break;
    }

    const listInfo = getListInfo(element);
    if (!listInfo) break;

    const listTypeInfo = listTypeMap.get(listInfo.numId);
    const listType = listTypeInfo?.type || "bullet";

    const items: JSONContent[] = [];

    while (i < elements.length) {
      const el = elements[i];
      if (el.name !== "w:p" || !isListItem(el)) break;

      const info = getListInfo(el);
      if (!info || info.numId !== listInfo.numId) break;

      const paragraph = await convertParagraph(el, context);
      const listItemContent = Array.isArray(paragraph) ? paragraph[0] : paragraph;

      items.push({
        type: "listItem",
        content: [listItemContent],
      });

      i++;
    }

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

    result.push(listNode);
  }

  return { nodes: result, consumed: i - index };
};

/**
 * Process consecutive task lists
 */

const processTaskLists: ElementProcessor = async (elements, index) => {
  const items: JSONContent[] = [];
  let i = index;

  while (i < elements.length) {
    const element = elements[i];
    if (element.name !== "w:p" || !isTaskItem(element)) {
      break;
    }

    const { convertTaskItem } = await import("./converters");
    const taskItem = convertTaskItem(element);
    items.push(taskItem);
    i++;
  }

  return {
    nodes: [
      {
        type: "taskList",
        content: items,
      },
    ],
    consumed: i - index,
  };
};

/**
 * Process horizontal rule
 */

const processHorizontalRule: ElementProcessor = async () => {
  return { nodes: [{ type: "horizontalRule" }], consumed: 1 };
};

/**
 * Process regular paragraph
 */

const processParagraph: ElementProcessor = async (elements, index, context) => {
  const paragraphResult = await convertParagraph(elements[index], context);

  if (Array.isArray(paragraphResult)) {
    return { nodes: paragraphResult, consumed: 1 };
  }

  return { nodes: [paragraphResult], consumed: 1 };
};

/**
 * Check if a paragraph is empty
 */

const isEmptyParagraph = (element: Element): boolean => {
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
};

/**
 * Get element processor based on element type
 */

const getElementProcessor = (element: Element): ElementProcessor | null => {
  if (element.name === "w:tbl") {
    return processTable;
  }

  if (element.name === "w:p") {
    if (isCodeBlock(element)) {
      return processCodeBlocks;
    }

    if (isTaskItem(element)) {
      return processTaskLists;
    }

    if (isListItem(element)) {
      return processLists;
    }

    if (isHorizontalRule(element)) {
      return processHorizontalRule;
    }

    return processParagraph;
  }

  return null;
};

/**
 * Process all elements in document body
 */

export const processElements = async (
  elements: Element[],
  context: ProcessContext,
): Promise<JSONContent[]> => {
  const result: JSONContent[] = [];
  let i = 0;

  while (i < elements.length) {
    const element = elements[i];

    const processor = getElementProcessor(element);

    if (!processor) {
      i++;
      continue;
    }

    if (element.name === "w:p" && context.ignoreEmptyParagraphs && isEmptyParagraph(element)) {
      i++;
      continue;
    }

    const { nodes, consumed } = await processor(elements, i, context);
    result.push(...nodes);
    i += consumed;
  }

  return result;
};
