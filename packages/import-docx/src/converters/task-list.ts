import type { Element, Text } from "xast";
import type { JSONContent } from "@tiptap/core";
import { findChild } from "../utils/xml";
import { extractAlignment } from "./text";

/**
 * Checkbox symbols used in DOCX
 */
const CHECKBOX_UNCHECKED = "☐";
const CHECKBOX_CHECKED = "☑";

/**
 * Check if a paragraph is a task item
 */
export function isTaskItem(node: Element): boolean {
  // Get the first text run
  const run = findChild(node, "w:r");
  if (!run) return false;

  const textElement = findChild(run, "w:t");
  if (!textElement) return false;

  const textNode = textElement.children.find((c): c is Text => c.type === "text");
  if (!textNode || !textNode.value) return false;

  const text = textNode.value;
  return text.startsWith(CHECKBOX_UNCHECKED) || text.startsWith(CHECKBOX_CHECKED);
}

/**
 * Get the checked state from a task item
 */
export function getTaskItemChecked(node: Element): boolean {
  const run = findChild(node, "w:r");
  if (!run) return false;

  const textElement = findChild(run, "w:t");
  if (!textElement) return false;

  const textNode = textElement.children.find((c): c is Text => c.type === "text");
  if (!textNode || !textNode.value) return false;

  return textNode.value.startsWith(CHECKBOX_CHECKED);
}

/**
 * Convert a task item to TipTap JSON
 * This removes the checkbox symbol from the text
 */
export function convertTaskItem(node: Element): JSONContent {
  const checked = getTaskItemChecked(node);

  // Convert the paragraph, but we need to remove the checkbox from the first text run
  const paragraph = convertTaskItemParagraph(node);

  return {
    type: "taskItem",
    attrs: {
      checked,
    },
    content: [paragraph],
  };
}

/**
 * Convert a task item paragraph, removing the checkbox symbol
 */
function convertTaskItemParagraph(node: Element): JSONContent {
  const content: JSONContent[] = [];
  let firstTextSkipped = false;

  for (const child of node.children) {
    if (child.type === "element" && child.name === "w:r") {
      // Check if this is the checkbox run
      let isCheckboxRun = false;

      if (!firstTextSkipped) {
        const textElement = findChild(child, "w:t");
        if (textElement) {
          const textNode = textElement.children.find((c): c is Text => c.type === "text");
          if (textNode && textNode.value) {
            const text = textNode.value;
            if (text.startsWith(CHECKBOX_UNCHECKED) || text.startsWith(CHECKBOX_CHECKED)) {
              isCheckboxRun = true;
              firstTextSkipped = true;

              // Extract the remaining text after the checkbox
              const remainingText = text.substring(2).trimStart();
              if (remainingText.length > 0) {
                content.push({
                  type: "text",
                  text: remainingText,
                });
              }
            }
          }
        }
      }

      if (!isCheckboxRun) {
        // Convert run to text marks
        const marks = extractMarksFromRun(child);

        const textElement = findChild(child, "w:t");
        if (textElement) {
          const textNode = textElement.children.find((c): c is Text => c.type === "text");
          if (textNode && textNode.value) {
            const textNodeData = {
              type: "text",
              text: textNode.value,
            } as {
              type: string;
              text: string;
              marks?: Array<{ type: string }>;
            };

            if (marks.length > 0) {
              textNodeData.marks = marks;
            }

            content.push(textNodeData);
          }
        }
      }
    }
  }

  // Extract paragraph alignment
  const attrs = extractAlignment(node);

  return {
    type: "paragraph",
    ...(attrs && { attrs }),
    content: content.length > 0 ? content : undefined,
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
