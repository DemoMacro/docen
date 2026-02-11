import { TextRun, type IParagraphOptions } from "docx";
import { TaskItemNode } from "@docen/extensions/types";
import { convertText, convertHardBreak } from "./text";
import { CHECKBOX_SYMBOLS } from "../utils";

/**
 * Convert TipTap task item node to paragraph options with checkbox
 *
 * This converter only handles data transformation from node content to DOCX format properties.
 * It returns pure data objects (IParagraphOptions), not DOCX instances.
 *
 * @param node - TipTap task item node
 * @returns Paragraph options (pure data object) with checkbox
 */
export function convertTaskItem(node: TaskItemNode): IParagraphOptions {
  if (!node.content || node.content.length === 0) {
    return {};
  }

  // Convert the first paragraph in the task item
  const firstParagraph = node.content[0];
  if (firstParagraph.type === "paragraph") {
    // Add checkbox based on checked state
    const isChecked = node.attrs?.checked || false;
    const checkboxText = isChecked
      ? CHECKBOX_SYMBOLS.checked + " "
      : CHECKBOX_SYMBOLS.unchecked + " ";

    // Convert paragraph content to text runs
    const children =
      firstParagraph.content?.flatMap((contentNode) => {
        if (contentNode.type === "text") {
          return convertText(contentNode);
        } else if (contentNode.type === "hardBreak") {
          return convertHardBreak(contentNode.marks);
        }
        return [];
      }) || [];

    // Add checkbox as first text run
    const checkboxRun = new TextRun({ text: checkboxText });

    return {
      children: [checkboxRun, ...children],
    };
  }

  // Fallback to empty paragraph options
  return {};
}
