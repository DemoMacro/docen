import { Paragraph, HeadingLevel, TextRun, ExternalHyperlink } from "docx";
import { HeadingNode } from "../types";
import { convertTextNodes, convertText } from "./text";
import { applyParagraphStyleAttributes } from "../utils";

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
  const children = convertTextNodes(node.content).filter(
    (item): item is TextRun | ExternalHyperlink => item !== undefined,
  );

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
    alignment?: "left" | "right" | "center" | "both";
  } = {
    children,
    heading: headingMap[level],
  };

  // Handle paragraph style attributes from node.attrs
  if (node.attrs) {
    paragraphOptions = applyParagraphStyleAttributes(paragraphOptions, node.attrs);
  }

  // Create heading paragraph
  const paragraph = new Paragraph(paragraphOptions);

  return paragraph;
}
