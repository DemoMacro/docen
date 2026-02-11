import { type IParagraphOptions } from "docx";
import { TaskListNode, TaskItemNode } from "@docen/extensions/types";
import { convertTaskItem } from "./task-item";

/**
 * Convert TipTap task list node to array of paragraph options
 *
 * This converter only handles data transformation from node content to DOCX format properties.
 * It returns pure data objects (IParagraphOptions[]), not DOCX instances.
 *
 * @param node - TipTap task list node
 * @returns Array of paragraph options (pure data objects) with checkboxes
 */
export function convertTaskList(node: TaskListNode): IParagraphOptions[] {
  if (!node.content || node.content.length === 0) {
    return [];
  }

  // Convert each task item in the list
  return node.content
    .filter((item) => item.type === "taskItem")
    .map((item) => convertTaskItem(item as TaskItemNode));
}
