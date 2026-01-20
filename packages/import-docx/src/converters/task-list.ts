import type { Element, Text } from "xast";
import type { JSONContent } from "@tiptap/core";
import type { ParseContext } from "../parser";
import type { StyleInfo } from "../parsers/styles";
import { findChild } from "@docen/utils";
import { extractRuns, extractAlignment } from "./text";

const CHECKBOX_UNCHECKED = "☐";
const CHECKBOX_CHECKED = "☑";

/**
 * Get first text node from element
 */
function getFirstTextNode(node: Element): Text | null {
  const run = findChild(node, "w:r");
  if (!run) return null;

  const textElement = findChild(run, "w:t");
  if (!textElement) return null;

  const textNode = textElement.children.find((c): c is Text => c.type === "text");
  return (textNode?.value && textNode) || null;
}

/**
 * Check if a paragraph is a task item
 */
export function isTaskItem(node: Element): boolean {
  const textNode = getFirstTextNode(node);
  if (!textNode) return false;

  const text = textNode.value;
  return text.startsWith(CHECKBOX_UNCHECKED) || text.startsWith(CHECKBOX_CHECKED);
}

/**
 * Get the checked state from a task item
 */
export function getTaskItemChecked(node: Element): boolean {
  const textNode = getFirstTextNode(node);
  return textNode?.value.startsWith(CHECKBOX_CHECKED) || false;
}

/**
 * Convert a task item to TipTap JSON
 */
export async function convertTaskItem(
  node: Element,
  params: { context: ParseContext; styleInfo?: StyleInfo },
): Promise<JSONContent> {
  const checked = getTaskItemChecked(node);

  return {
    type: "taskItem",
    attrs: { checked },
    content: [await convertTaskItemParagraph(node, params)],
  };
}

/**
 * Convert task list (handles consecutive task items)
 */
export async function convertTaskList(
  _node: Element,
  params: {
    context: ParseContext;
    styleInfo?: StyleInfo;
    siblings: Element[];
    index: number;
    processedIndices: Set<number>;
  },
): Promise<JSONContent> {
  const { siblings, index, processedIndices } = params;

  // Collect consecutive task items
  const items: JSONContent[] = [];
  let i = index;

  while (i < siblings.length) {
    const el = siblings[i];
    if (el.name !== "w:p" || !isTaskItem(el)) {
      break;
    }

    // Mark this index as processed
    processedIndices.add(i);

    // Convert task item
    const taskItem = await convertTaskItem(el, {
      context: params.context,
      styleInfo: params.styleInfo,
    });
    items.push(taskItem);

    i++;
  }

  // Build task list node
  return {
    type: "taskList",
    content: items,
  };
}

/**
 * Convert a task item paragraph, removing the checkbox symbol
 */
async function convertTaskItemParagraph(
  node: Element,
  params: { context: ParseContext; styleInfo?: StyleInfo },
): Promise<JSONContent> {
  const { context, styleInfo } = params;
  const runs = await extractRuns(node, { context, styleInfo });

  // Remove checkbox text from the first text run
  if (runs.length > 0 && runs[0].type === "text") {
    const firstRun = runs[0] as { text: string; marks?: Array<{ type: string }> };
    const text = firstRun.text;
    if (text.startsWith(CHECKBOX_UNCHECKED) || text.startsWith(CHECKBOX_CHECKED)) {
      const remainingText = text.substring(2).trimStart();
      if (remainingText) {
        firstRun.text = remainingText;
      } else {
        runs.shift(); // Remove first run if no remaining text
      }
    }
  }

  const attrs = extractAlignment(node);

  return {
    type: "paragraph",
    ...(attrs && { attrs }),
    content: runs.length ? runs : undefined,
  };
}
