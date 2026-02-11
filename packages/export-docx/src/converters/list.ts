import { BulletListNode, OrderedListNode, ListItemNode } from "@docen/extensions/types";
import { type IParagraphOptions } from "docx";
import { convertListItem } from "./list-item";

export interface ListOptions {
  numbering: {
    reference: string;
    level: number;
  };
  start?: number;
}

export function convertBulletList(): ListOptions {
  return {
    numbering: {
      reference: "bullet-list",
      level: 0,
    },
  };
}

export function convertOrderedList(node: OrderedListNode): ListOptions {
  // Note: The start attribute should be handled at the document level
  // when creating numbering options. We return the start value
  // so the main export function can create appropriate numbering references.
  const start = node.attrs?.start || 1;

  return {
    numbering: {
      reference: "ordered-list",
      level: 0,
    },
    start,
  };
}

/**
 * Convert list nodes (bullet or ordered) with proper numbering
 *
 * This converter only handles data transformation from node content to DOCX format properties.
 * It returns pure data objects (IParagraphOptions[]), not DOCX instances.
 */
export async function convertList(
  node: BulletListNode | OrderedListNode,
  params: {
    listType: "bullet" | "ordered";
  },
): Promise<IParagraphOptions[]> {
  const { listType } = params;

  if (!node.content) {
    return [];
  }

  const elements: IParagraphOptions[] = [];

  // Get list options
  const listOptions =
    listType === "bullet" ? convertBulletList() : convertOrderedList(node as OrderedListNode);

  // Determine numbering reference based on start value
  let numberingReference = listOptions.numbering.reference;
  if (listType === "ordered" && listOptions.start && listOptions.start !== 1) {
    numberingReference = `ordered-list-start-${listOptions.start}`;
  }

  // Convert list items
  for (const item of node.content) {
    if (item.type === "listItem") {
      const paragraphOptions = await convertListItem(item as ListItemNode, {
        options: {
          numbering: {
            reference: numberingReference,
            level: 0,
          },
        },
      });
      elements.push(paragraphOptions);
    }
  }

  return elements;
}
