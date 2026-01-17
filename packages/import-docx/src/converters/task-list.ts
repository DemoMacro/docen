import type { Element, Text } from "xast";
import type { JSONContent } from "@tiptap/core";
import { findChild } from "../utils/xml";
import { extractAlignment } from "./text";

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
export function convertTaskItem(node: Element): JSONContent {
  const checked = getTaskItemChecked(node);

  return {
    type: "taskItem",
    attrs: { checked },
    content: [convertTaskItemParagraph(node)],
  };
}

/**
 * Convert a task item paragraph, removing the checkbox symbol
 */
function convertTaskItemParagraph(node: Element): JSONContent {
  const content: JSONContent[] = [];
  let firstTextSkipped = false;

  for (const child of node.children) {
    if (child.type !== "element" || child.name !== "w:r") continue;

    // Handle checkbox run
    if (!firstTextSkipped) {
      const textElement = findChild(child, "w:t");
      const textNode = textElement?.children.find((c): c is Text => c.type === "text");

      if (textNode?.value) {
        const text = textNode.value;
        if (text.startsWith(CHECKBOX_UNCHECKED) || text.startsWith(CHECKBOX_CHECKED)) {
          firstTextSkipped = true;
          const remainingText = text.substring(2).trimStart();
          if (remainingText) {
            content.push({ type: "text", text: remainingText });
          }
          continue;
        }
      }
    }

    // Convert regular run
    const marks = extractMarksFromRun(child);
    const textElement = findChild(child, "w:t");
    const textNode = textElement?.children.find((c): c is Text => c.type === "text");

    if (textNode?.value) {
      const textNodeData: {
        type: string;
        text: string;
        marks?: Array<{ type: string }>;
      } = {
        type: "text",
        text: textNode.value,
      };

      if (marks.length) textNodeData.marks = marks;
      content.push(textNodeData);
    }
  }

  const attrs = extractAlignment(node);

  return {
    type: "paragraph",
    ...(attrs && { attrs }),
    content: content.length ? content : undefined,
  };
}

/**
 * Extract marks from a run (bold, italic, etc.)
 */
function extractMarksFromRun(run: Element): Array<{ type: string }> {
  const marks: Array<{ type: string }> = [];
  const rPr = findChild(run, "w:rPr");
  if (!rPr) return marks;

  if (findChild(rPr, "w:b")) marks.push({ type: "bold" });
  if (findChild(rPr, "w:i")) marks.push({ type: "italic" });
  if (findChild(rPr, "w:u")) marks.push({ type: "underline" });
  if (findChild(rPr, "w:strike")) marks.push({ type: "strike" });

  return marks;
}
