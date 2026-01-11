import type { Element } from "xast";
import type { JSONContent } from "@tiptap/core";

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
  for (const child of node.children) {
    if (child.type === "element" && child.name === "w:r") {
      for (const runChild of child.children) {
        if (runChild.type === "element" && runChild.name === "w:t") {
          const textNode = runChild.children.find((c) => c.type === "text");
          if (textNode && "value" in textNode) {
            const text = (textNode as { value: string }).value;
            // Check if text starts with checkbox symbol
            return (
              text.startsWith(CHECKBOX_UNCHECKED) ||
              text.startsWith(CHECKBOX_CHECKED)
            );
          }
        }
      }
      break;
    }
  }
  return false;
}

/**
 * Get the checked state from a task item
 */
export function getTaskItemChecked(node: Element): boolean {
  for (const child of node.children) {
    if (child.type === "element" && child.name === "w:r") {
      for (const runChild of child.children) {
        if (runChild.type === "element" && runChild.name === "w:t") {
          const textNode = runChild.children.find((c) => c.type === "text");
          if (textNode && "value" in textNode) {
            const text = (textNode as { value: string }).value;
            return text.startsWith(CHECKBOX_CHECKED);
          }
        }
      }
      break;
    }
  }
  return false;
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
        for (const runChild of child.children) {
          if (runChild.type === "element" && runChild.name === "w:t") {
            const textNode = runChild.children.find((c) => c.type === "text");
            if (textNode && "value" in textNode) {
              const text = (textNode as { value: string }).value;
              if (
                text.startsWith(CHECKBOX_UNCHECKED) ||
                text.startsWith(CHECKBOX_CHECKED)
              ) {
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
      }

      if (!isCheckboxRun) {
        // Convert run to text marks
        const marks = extractMarksFromRun(child);

        for (const runChild of child.children) {
          if (runChild.type === "element" && runChild.name === "w:t") {
            const textNode = runChild.children.find((c) => c.type === "text");
            if (textNode && "value" in textNode) {
              const textNodeData = {
                type: "text",
                text: (textNode as { value: string }).value,
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

  for (const child of run.children) {
    if (child.type === "element" && child.name === "w:rPr") {
      const rPr = child;

      // Check for bold
      for (const prop of rPr.children) {
        if (prop.type === "element" && prop.name === "w:b") {
          marks.push({ type: "bold" });
          break;
        }
      }

      // Check for italic
      for (const prop of rPr.children) {
        if (prop.type === "element" && prop.name === "w:i") {
          marks.push({ type: "italic" });
          break;
        }
      }

      // Check for underline
      for (const prop of rPr.children) {
        if (prop.type === "element" && prop.name === "w:u") {
          marks.push({ type: "underline" });
          break;
        }
      }

      // Check for strike
      for (const prop of rPr.children) {
        if (prop.type === "element" && prop.name === "w:strike") {
          marks.push({ type: "strike" });
          break;
        }
      }

      break;
    }
  }

  return marks;
}

/**
 * Extract paragraph alignment
 */
function extractAlignment(
  node: Element,
): { textAlign: "left" | "right" | "center" | "justify" } | undefined {
  for (const child of node.children) {
    if (child.type === "element" && child.name === "w:pPr") {
      const pPr = child;
      for (const prop of pPr.children) {
        if (prop.type === "element" && prop.name === "w:jc") {
          const align = prop.attributes["w:val"] as string;
          if (align === "both") return { textAlign: "justify" };
          if (align === "center") return { textAlign: "center" };
          if (align === "right") return { textAlign: "right" };
          if (align === "left") return { textAlign: "left" };
        }
      }
    }
  }
  return undefined;
}
