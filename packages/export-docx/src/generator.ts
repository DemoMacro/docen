import { JSONContent } from "@tiptap/core";
import {
  Document,
  Paragraph,
  TextRun,
  Packer,
  OutputType,
  OutputByType,
  INumberingOptions,
  ILevelsOptions,
  IPropertiesOptions,
  FileChild,
  LevelFormat,
  AlignmentType,
  convertInchesToTwip,
  TableOfContents,
  IParagraphStyleOptions,
  IParagraphOptions,
  Table,
} from "docx";
import { type DocxExportOptions } from "./options";
import { calculateEffectiveContentWidth } from "./utils";
import { convertParagraph } from "./converters/paragraph";
import { convertHeading } from "./converters/heading";
import { convertBlockquote } from "./converters/blockquote";
import { convertImage } from "./converters/image";
import { convertTable } from "./converters/table";
import { convertCodeBlock } from "./converters/code-block";
import { convertList } from "./converters/list";
import { convertTaskList } from "./converters/task-list";
import { convertTaskItem } from "./converters/task-item";
import { convertHorizontalRule } from "./converters/horizontal-rule";
import { convertDetailsSummary } from "./converters/details";
import { convertHardBreak } from "./converters/text";
import type {
  ParagraphNode,
  HeadingNode,
  BlockquoteNode,
  CodeBlockNode,
  ImageNode,
  TableNode,
  TaskListNode,
  TaskItemNode,
  OrderedListNode,
  BulletListNode,
  HorizontalRuleNode,
  DetailsSummaryNode,
} from "@docen/extensions/types";

/**
 * Convert TipTap JSONContent to DOCX format
 *
 * @param docJson - TipTap document JSON
 * @param options - Export options
 * @returns Promise with DOCX in specified format
 */
export async function generateDOCX<T extends OutputType>(
  docJson: JSONContent,
  options: DocxExportOptions<T>,
): Promise<OutputByType[T]> {
  const {
    // Document metadata
    title,
    subject,
    creator,
    keywords,
    description,
    lastModifiedBy,
    revision,

    // Styling
    styles,

    // Table of contents
    tableOfContents,

    // Document options
    sections,
    fonts,
    hyphenation,
    compatibility,
    customProperties,
    evenAndOddHeaderAndFooters,
    defaultTabStop,

    // Export options
    outputType,
  } = options;

  // Convert document content
  const children = await convertDocument(docJson, { options });

  // Create table of contents if configured
  const tocElement = tableOfContents
    ? new TableOfContents(tableOfContents.title, {
        ...tableOfContents.run,
      })
    : null;

  // Collect ordered list start values for numbering options
  const numberingOptions = createNumberingOptions(docJson);

  // Build styles - merge user styles with auto-generated image/table styles
  const additionalParagraphStyles: IParagraphStyleOptions[] = [];

  // Add image style if configured
  if (options.image?.style) {
    additionalParagraphStyles.push(options.image.style);
  }

  // Add table style if configured
  if (options.table?.style) {
    additionalParagraphStyles.push(options.table.style);
  }

  // Add code block style if configured
  if (options.code?.style) {
    additionalParagraphStyles.push(options.code.style);
  }

  const mergedStyles = styles
    ? {
        ...styles,
        ...(additionalParagraphStyles.length > 0 && {
          paragraphStyles: [...(styles.paragraphStyles || []), ...additionalParagraphStyles],
        }),
      }
    : {};

  // Build document sections - merge user config with generated content
  const documentSections = sections
    ? sections.map((section, index) => {
        const sectionChildren: FileChild[] = [];

        // Add table of contents to first section if configured
        if (index === 0 && tocElement) {
          sectionChildren.push(tocElement);
        }

        // Add main content to first section
        if (index === 0) {
          sectionChildren.push(...children);
        }

        return {
          ...section,
          ...(sectionChildren.length > 0 ? { children: sectionChildren } : {}),
        };
      })
    : [
        {
          children: tocElement ? [tocElement, ...children] : children,
        },
      ];

  // Build document options
  const docOptions: IPropertiesOptions = {
    // Sections - required
    sections: documentSections,

    // Metadata
    title: title || "Document",
    subject: subject || "",
    creator: creator || "",
    keywords: keywords || "",
    description: description || "",
    lastModifiedBy: lastModifiedBy || "",
    revision: revision || 1,

    // Styling
    styles: mergedStyles,
    numbering: numberingOptions,

    // Optional properties - only include if provided
    ...(fonts && fonts.length > 0 && { fonts }),
    ...(hyphenation && { hyphenation }),
    ...(compatibility && { compatibility }),
    ...(customProperties && customProperties.length > 0 && { customProperties }),
    ...(evenAndOddHeaderAndFooters !== undefined && { evenAndOddHeaderAndFooters }),
    ...(defaultTabStop !== undefined && { defaultTabStop }),
  };

  const doc = new Document(docOptions);

  return Packer.pack(doc, outputType || "arraybuffer") as Promise<OutputByType[T]>;
}

/**
 * Convert document content to DOCX elements
 */
export async function convertDocument(
  node: JSONContent,
  params: {
    options: DocxExportOptions;
  },
): Promise<FileChild[]> {
  const elements: FileChild[] = [];

  if (!node || !Array.isArray(node.content)) {
    return elements;
  }

  // Pre-calculate effective content width once for all images
  const effectiveContentWidth = calculateEffectiveContentWidth(params.options);

  for (const childNode of node.content) {
    const element = await convertNode(childNode, params.options, effectiveContentWidth);
    if (Array.isArray(element)) {
      elements.push(...element);
    } else if (element) {
      elements.push(element);

      // Insert empty paragraph between adjacent tables to prevent merging
      if (
        childNode.type === "table" &&
        elements.length >= 2 &&
        elements[elements.length - 2] instanceof Table
      ) {
        elements.push(new Paragraph({}));
      }
    }
  }

  return elements;
}

/**
 * Convert a single node to DOCX element(s)
 *
 * This function implements a three-layer architecture:
 * 1. Data Transformation: Convert node.attrs → IParagraphOptions (pure data)
 * 2. Style Application: Apply styleId references (if configured)
 * 3. Object Creation: Create actual DOCX instances (Paragraph, Table, etc.)
 */
export async function convertNode(
  node: JSONContent,
  options: DocxExportOptions,
  effectiveContentWidth: number,
): Promise<FileChild | FileChild[] | null> {
  if (!node || !node.type) {
    return null;
  }

  // Layer 1: Data Transformation (node.attrs → IParagraphOptions)
  const dataResult = await convertNodeData(node, options, effectiveContentWidth);

  // Handle Table and FileChild[] (already final objects)
  if (dataResult instanceof Table) {
    return dataResult;
  }

  // Handle arrays of IParagraphOptions - convert to Paragraph objects
  if (Array.isArray(dataResult)) {
    // For arrays, convert each IParagraphOptions to Paragraph
    const styleId = getStyleIdByNodeType(node.type, options);
    return dataResult.map((paragraphOptions: IParagraphOptions): FileChild => {
      const styledOptions = applyStyleReference(paragraphOptions, styleId);
      // We know this won't be a Table since we started with IParagraphOptions
      return new Paragraph(styledOptions);
    });
  }

  // Layer 2: Style Application (apply styleId if configured)
  const styleId = getStyleIdByNodeType(node.type, options);
  const styledOptions = applyStyleReference(dataResult, styleId);

  // Layer 3: Object Creation (create DOCX instance)
  return createDOCXObject(styledOptions);
}

/**
 * Layer 1: Data Transformation
 *
 * Convert node data to DOCX format properties.
 * Returns pure data objects (IParagraphOptions) or arrays, not DOCX instances.
 * This layer does NOT handle styleId references.
 */
async function convertNodeData(
  node: JSONContent,
  options: DocxExportOptions,
  effectiveContentWidth: number,
): Promise<IParagraphOptions | IParagraphOptions[] | Table | FileChild[]> {
  switch (node.type) {
    case "paragraph":
      return await convertParagraph(node as ParagraphNode, {
        image: {
          maxWidth: effectiveContentWidth,
          options: options.image?.run,
          handler: options.image?.handler,
        },
      });

    case "heading":
      return convertHeading(node as HeadingNode);

    case "blockquote":
      return convertBlockquote(node as BlockquoteNode);

    case "codeBlock":
      return convertCodeBlock(node as CodeBlockNode);

    case "image":
      // Image is special: returns paragraph options wrapping an ImageRun
      const imageRun = await convertImage(node as ImageNode, {
        maxWidth: effectiveContentWidth,
        options: options.image?.run,
        handler: options.image?.handler,
      });
      return { children: [imageRun] };

    case "table":
      return await convertTable(node as TableNode, {
        options: options.table,
      });

    case "bulletList":
      return await convertList(node as BulletListNode, {
        listType: "bullet",
      });

    case "orderedList":
      return await convertList(node as OrderedListNode, {
        listType: "ordered",
      });

    case "taskList":
      return convertTaskList(node as TaskListNode);

    case "taskItem":
      return convertTaskItem(node as TaskItemNode);

    case "hardBreak":
      // Wrap hardBreak in a paragraph
      return { children: [convertHardBreak()] };

    case "horizontalRule":
      return convertHorizontalRule(node as HorizontalRuleNode, {
        options: options.horizontalRule,
      });

    case "details":
      // Flatten details: expand summary and content directly into document flow
      return await convertDetails(node, options, effectiveContentWidth);

    case "detailsSummary":
      return convertDetailsSummary(node as DetailsSummaryNode, {
        options: options.details,
      });

    default:
      // Unknown node type, return a paragraph with text
      return {
        children: [new TextRun({ text: `[Unsupported: ${node.type}]` })],
      };
  }
}

/**
 * Helper to convert details node (needs to recursively call convertNode)
 */
async function convertDetails(
  node: JSONContent,
  options: DocxExportOptions,
  effectiveContentWidth: number,
): Promise<FileChild[]> {
  const elements: FileChild[] = [];
  if (node.content) {
    for (const child of node.content) {
      const element = await convertNode(child, options, effectiveContentWidth);
      if (Array.isArray(element)) {
        elements.push(...element);
      } else if (element) {
        elements.push(element);
      }
    }
  }
  return elements;
}

/**
 * Create a single ordered list level configuration
 */
const createOrderedListLevel = (start?: number): ILevelsOptions => ({
  level: 0,
  format: LevelFormat.DECIMAL,
  text: "%1.",
  alignment: AlignmentType.START,
  start: start ?? 1,
  style: {
    paragraph: {
      indent: {
        left: convertInchesToTwip(0.5),
        hanging: convertInchesToTwip(0.25),
      },
    },
  },
});

/**
 * Create a numbering reference configuration
 */
const createNumberingReference = (start?: number) => ({
  reference: start && start !== 1 ? `ordered-list-start-${start}` : "ordered-list",
  levels: [createOrderedListLevel(start)],
});

/**
 * Create numbering options for the document
 */
function createNumberingOptions(docJson: JSONContent): INumberingOptions {
  // Collect all unique ordered list start values
  const orderedListStarts = new Set<number>();

  function collectListStarts(node: JSONContent) {
    if (node.type === "orderedList" && node.attrs?.start) {
      orderedListStarts.add(node.attrs.start);
    }
    if (node.content) {
      for (const child of node.content) {
        collectListStarts(child);
      }
    }
  }

  collectListStarts(docJson);

  // Build numbering options
  const bulletLevel: ILevelsOptions = {
    level: 0,
    format: LevelFormat.BULLET,
    text: "•",
    alignment: AlignmentType.START,
    style: {
      paragraph: {
        indent: {
          left: convertInchesToTwip(0.5),
          hanging: convertInchesToTwip(0.25),
        },
      },
    },
  };

  // Create the final numbering options
  const numberingOptions: Array<{
    reference: string;
    levels: ILevelsOptions[];
  }> = [
    {
      reference: "bullet-list",
      levels: [bulletLevel],
    },
    createNumberingReference(1),
  ];

  // Add options for custom start values
  orderedListStarts.forEach((start) => {
    if (start !== 1) {
      numberingOptions.push(createNumberingReference(start));
    }
  });

  return { config: numberingOptions };
}

/**
 * Get style ID for a specific node type from export options
 *
 * This is a centralized mapping of node types to their configured style IDs.
 * Style references are applied separately from data transformation.
 *
 * @param nodeType - The type of TipTap node
 * @param options - Export options containing style configurations
 * @returns Style ID string if configured, undefined otherwise
 */
function getStyleIdByNodeType(nodeType: string, options: DocxExportOptions): string | undefined {
  const styleMap: Record<string, string | undefined> = {
    codeBlock: options.code?.style?.id,
    image: options.image?.style?.id,
    // Note: table, heading, paragraph, blockquote, etc. don't use styleId
    // They rely on direct formatting from node.attrs
  };

  return styleMap[nodeType];
}

/**
 * Apply style reference to paragraph options
 *
 * This function handles the final step of adding a style ID reference to
 * paragraph options. It's called after data transformation is complete.
 *
 * @param paragraphOptions - Paragraph options from converter
 * @param styleId - Style ID to apply (optional)
 * @returns Paragraph options with style ID applied if provided
 */
function applyStyleReference(
  paragraphOptions: IParagraphOptions,
  styleId?: string,
): IParagraphOptions {
  if (!styleId) {
    return paragraphOptions;
  }

  // Style is just a string reference, applied at the end
  return {
    ...paragraphOptions,
    style: styleId,
  };
}

/**
 * Create a DOCX object from paragraph options
 *
 * This is the final step that creates actual DOCX instances from
 * pure data objects.
 *
 * @param options - Paragraph options or table
 * @returns DOCX Paragraph or Table instance
 */
function createDOCXObject(options: IParagraphOptions | Table): Paragraph | Table {
  if (options instanceof Table) {
    return options;
  }
  return new Paragraph(options);
}
