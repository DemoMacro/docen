import { type IParagraphOptions } from "docx";
import { convertText, convertHardBreak } from "./text";
import { BlockquoteNode } from "@docen/extensions/types";

/**
 * Convert TipTap blockquote node to array of paragraph options
 *
 * This converter only handles data transformation from node content to DOCX format properties.
 * It returns pure data objects (IParagraphOptions[]), not DOCX instances.
 *
 * @param node - TipTap blockquote node
 * @returns Array of paragraph options (pure data objects)
 */
export function convertBlockquote(node: BlockquoteNode): IParagraphOptions[] {
  if (!node.content) return [];

  return node.content.map((contentNode) => {
    if (contentNode.type === "paragraph") {
      // Convert paragraph content
      const children =
        contentNode.content?.flatMap((node) => {
          if (node.type === "text") {
            return convertText(node);
          } else if (node.type === "hardBreak") {
            return convertHardBreak(node.marks);
          }
          return [];
        }) || [];

      // Return paragraph options with blockquote styling
      return {
        children,
        indent: {
          left: 720,
        },
        border: {
          left: {
            style: "single",
          },
        },
      };
    }

    // Handle other content types within blockquote
    // For now, return empty paragraph options as fallback
    return {};
  });
}
