import { HeadingLevel, TextRun, ExternalHyperlink, type IParagraphOptions } from "docx";
import { HeadingNode } from "@docen/extensions/types";
import { convertTextNodes } from "./text";
import { applyParagraphStyleAttributes } from "../utils";

/**
 * Convert TipTap heading node to DOCX paragraph options
 *
 * This converter only handles data transformation from node.attrs to DOCX format properties.
 * It returns pure data objects (IParagraphOptions), not DOCX instances.
 *
 * @param node - TipTap heading node
 * @returns DOCX paragraph options (pure data object)
 */
export function convertHeading(node: HeadingNode): IParagraphOptions {
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

  // Build paragraph options with heading level
  let paragraphOptions: IParagraphOptions = {
    children,
    heading: headingMap[level],
  };

  // Handle paragraph style attributes from node.attrs
  if (node.attrs) {
    paragraphOptions = applyParagraphStyleAttributes(paragraphOptions, node.attrs);
  }

  return paragraphOptions;
}
