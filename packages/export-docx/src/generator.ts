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
import { convertListItem } from "./converters/list-item";
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
  ListItemNode,
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
 */
export async function convertNode(
  node: JSONContent,
  options: DocxExportOptions,
  effectiveContentWidth: number,
): Promise<FileChild | FileChild[] | null> {
  if (!node || !node.type) {
    return null;
  }

  switch (node.type) {
    case "paragraph":
      return await convertParagraph(node as ParagraphNode, {
        image: {
          maxWidth: effectiveContentWidth,
        },
      });

    case "heading":
      return convertHeading(node as HeadingNode);

    case "blockquote":
      return convertBlockquote(node as BlockquoteNode);

    case "codeBlock":
      return convertCodeBlock(node as CodeBlockNode);

    case "image":
      // Convert image node to ImageRun and wrap in Paragraph with style
      const imageRun = await convertImage(node as ImageNode, {
        maxWidth: effectiveContentWidth,
      });

      // Build paragraph options with style reference if configured
      const imageParagraphOptions: IParagraphOptions = options.image?.style
        ? {
            children: [imageRun],
            style: options.image.style.id,
          }
        : {
            children: [imageRun],
          };

      return new Paragraph(imageParagraphOptions);

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

    case "listItem":
      return convertListItem(node as ListItemNode, {
        options: undefined,
      });

    case "taskItem":
      return convertTaskItem(node as TaskItemNode);

    case "hardBreak":
      // Wrap hardBreak in a paragraph
      return new Paragraph({ children: [convertHardBreak()] });

    case "horizontalRule":
      return convertHorizontalRule(node as HorizontalRuleNode, {
        options: options.horizontalRule,
      });

    case "details":
      // Flatten details: expand summary and content directly into document flow
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

    case "detailsSummary":
      return convertDetailsSummary(node as DetailsSummaryNode, {
        options: options.details,
      });

    // detailsContent is automatically expanded when details is processed

    default:
      // Unknown node type, return a paragraph with text
      return new Paragraph({
        children: [new TextRun({ text: `[Unsupported: ${node.type}]` })],
      });
  }
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
