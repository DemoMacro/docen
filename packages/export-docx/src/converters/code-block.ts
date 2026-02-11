import { TextRun, type IParagraphOptions } from "docx";
import { CodeBlockNode, TextNode } from "@docen/extensions/types";
import { DEFAULT_CODE_FONT } from "../utils";
import { convertText } from "./text";

/**
 * Convert TipTap codeBlock node to DOCX paragraph options
 *
 * This converter only handles data transformation from node.attrs to DOCX format properties.
 * It returns pure data objects (IParagraphOptions), not DOCX instances.
 *
 * @param node - TipTap codeBlock node
 * @returns DOCX paragraph options (pure data object)
 */
export function convertCodeBlock(node: CodeBlockNode): IParagraphOptions {
  // If no content, return empty paragraph options with code font
  if (!node.content || node.content.length === 0) {
    return {
      children: [
        new TextRun({
          text: "",
          font: DEFAULT_CODE_FONT,
        }),
      ],
    };
  }

  // Process each text node through convertText to preserve formatting
  const textRuns = node.content.flatMap((contentNode) => {
    if (contentNode.type === "text") {
      return convertText(contentNode as TextNode);
    }
    return [];
  });

  // Return paragraph options with processed text runs
  return {
    children: textRuns.length > 0 ? textRuns : [new TextRun({ text: "", font: DEFAULT_CODE_FONT })],
  };
}
