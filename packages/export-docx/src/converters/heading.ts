import { Paragraph, HeadingLevel } from "docx";
import { HeadingNode } from "../types";
import { convertText, convertHardBreak } from "./text";

/**
 * Convert pixels to TWIPs (Twentieth of a Point)
 * 1 inch = 1440 TWIPs, 1px ≈ 15 TWIPs (at 96 DPI: 1px = 0.75pt = 15 TWIP)
 */
function pxToTwip(px: number): number {
  return Math.round(px * 15);
}

/**
 * Convert TipTap heading node to DOCX paragraph
 *
 * @param node - TipTap heading node
 * @returns DOCX Paragraph object
 */
export function convertHeading(node: HeadingNode): Paragraph {
  // Get heading level
  const level: HeadingNode["attrs"]["level"] = node?.attrs?.level;

  // Convert content using shared text converter
  const children =
    node.content?.flatMap((contentNode) => {
      if (contentNode.type === "text") {
        return convertText(contentNode);
      } else if (contentNode.type === "hardBreak") {
        return convertHardBreak(contentNode.marks);
      }
      return [];
    }) || [];

  // Map to DOCX heading levels
  const headingMap: Record<
    HeadingNode["attrs"]["level"],
    (typeof HeadingLevel)[keyof typeof HeadingLevel]
  > = {
    1: HeadingLevel.HEADING_1,
    2: HeadingLevel.HEADING_2,
    3: HeadingLevel.HEADING_3,
    4: HeadingLevel.HEADING_4,
    5: HeadingLevel.HEADING_5,
    6: HeadingLevel.HEADING_6,
  };

  // Build paragraph options
  let paragraphOptions: {
    children: ReturnType<typeof convertText>[];
    heading: (typeof HeadingLevel)[keyof typeof HeadingLevel];
    indent?: {
      left?: number;
      right?: number;
      firstLine?: number;
    };
    spacing?: {
      before?: number;
      after?: number;
    };
  } = {
    children,
    heading: headingMap[level],
  };

  // Handle paragraph style attributes from node.attrs
  if (node.attrs) {
    const { indentLeft, indentRight, indentFirstLine, spacingBefore, spacingAfter } = node.attrs;

    // Convert indentation to DOCX format
    if (indentLeft || indentRight || indentFirstLine) {
      paragraphOptions = {
        ...paragraphOptions,
        indent: {
          ...(indentLeft && { left: pxToTwip(indentLeft) }),
          ...(indentRight && { right: pxToTwip(indentRight) }),
          ...(indentFirstLine && { firstLine: pxToTwip(indentFirstLine) }),
        },
      };
    }

    // Convert spacing to DOCX format
    if (spacingBefore || spacingAfter) {
      paragraphOptions = {
        ...paragraphOptions,
        spacing: {
          ...(spacingBefore && { before: pxToTwip(spacingBefore) }),
          ...(spacingAfter && { after: pxToTwip(spacingAfter) }),
        },
      };
    }
  }

  // Create heading paragraph
  const paragraph = new Paragraph(paragraphOptions);

  return paragraph;
}
