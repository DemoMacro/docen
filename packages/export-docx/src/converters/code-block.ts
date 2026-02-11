import { Paragraph, TextRun } from "docx";
import { CodeBlockNode, TextNode } from "@docen/extensions/types";
import { DEFAULT_CODE_FONT } from "../utils";
import { convertText } from "./text";

/**
 * Convert TipTap codeBlock node to DOCX Paragraph
 *
 * @param node - TipTap codeBlock node
 * @returns DOCX Paragraph object with code styling
 */
export function convertCodeBlock(node: CodeBlockNode): Paragraph {
  // If no content, return empty paragraph with code font
  if (!node.content || node.content.length === 0) {
    return new Paragraph({
      children: [
        new TextRun({
          text: "",
          font: DEFAULT_CODE_FONT,
        }),
      ],
    });
  }

  // Process each text node through convertText to preserve formatting
  const textRuns = node.content.flatMap((contentNode) => {
    if (contentNode.type === "text") {
      return convertText(contentNode as TextNode);
    }
    return [];
  });

  // Create paragraph with processed text runs
  return new Paragraph({
    children: textRuns.length > 0 ? textRuns : [new TextRun({ text: "", font: DEFAULT_CODE_FONT })],
  });
}
