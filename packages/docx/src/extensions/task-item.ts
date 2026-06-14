import { TaskItem as TaskItemBase } from "./tiptap";

/**
 * TaskItem extension — owns the DOCX expression of a task-list checkbox.
 *
 * DOCX has no native task list, but a clickable checkbox is reversible via an
 * inline content-control SDT (w14:checkbox). Each task paragraph carries a
 * leading checkbox SDT tagged "docen-task" so resolve can tell task items apart
 * from ordinary paragraphs that happen to contain an SDT. The checked state
 * round-trips through the SDT; DocxManager injects/strips the SDT at the
 * paragraph boundary and rebuilds the taskList/taskItem tree.
 */

/** SDT tag marking our task-item checkbox content control. */
export const TASK_CHECKBOX_TAG = "docen-task";

/**
 * Inline checkbox SDT (w14:checkbox) for a task item, as a ParagraphChild.
 * Tagged so resolve can distinguish task items from ordinary SDT-bearing
 * paragraphs.
 */
export function createTaskCheckbox(checked: boolean): Record<string, unknown> {
  return { sdt: { properties: { tag: TASK_CHECKBOX_TAG, checkbox: { checked } } } };
}

/** True if an inline ParagraphChild is our task checkbox SDT. */
export function isTaskCheckbox(child: unknown): boolean {
  if (typeof child !== "object" || child === null || !("sdt" in child)) return false;
  const props = (child as { sdt: { properties?: { tag?: string } } }).sdt?.properties;
  return props?.tag === TASK_CHECKBOX_TAG;
}

/** Read the checked state of a task checkbox SDT child (false if not one). */
export function readCheckboxState(child: unknown): boolean {
  if (!isTaskCheckbox(child)) return false;
  return Boolean(
    (child as { sdt: { properties?: { checkbox?: { checked?: boolean } } } }).sdt?.properties
      ?.checkbox?.checked,
  );
}

// DocxManager injects/strips the checkbox SDT via the helpers above; the
// extension itself carries no DOCX attrs of its own.
export { TaskItemBase as TaskItem };
