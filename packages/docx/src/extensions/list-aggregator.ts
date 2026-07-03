import type { ParagraphOptions } from "@office-open/docx";

import { buildTextBlock } from "../converters/styles";
import type { JSONContent } from "../core";
import { Extension } from "../core";
import { detectHeadingLevel } from "./heading";
import * as taskItemExt from "./task-item";
import type { ParseAggregatorRule, ResolveContext } from "./types";

/**
 * ListAggregator — a plain Extension owning the DOCX → Tiptap list rebuild.
 *
 * DOCX stores lists as a flat run of paragraphs referencing a numbering
 * definition (or carrying a `bullet` marker); Tiptap models a nested
 * bulletList/orderedList/taskList tree. This module owns the whole rebuild —
 * classify (detectList), tree-build (buildListTree), and list-item paragraph
 * resolution — as one algorithm spanning all three list kinds. It declares a
 * parseDocxAggregator rule; DocxManager runs the generic group-by-predicate
 * loop and hands the run of list paragraphs to build.
 *
 * Task items carry a leading checkbox SDT (docen-task); the checkbox predicate
 * + readCheckboxState live in task-item.ts (shared with compile).
 * resolveListItemParagraph reuses the heading/paragraph buildTextBlock path so
 * a list item can itself be a heading.
 */

interface ListInfo {
  kind: "bullet" | "ordered" | "task";
  level: number;
  reference?: string;
  start?: number;
  checked: boolean;
}

/** Classify a paragraph as a list item, or null if it isn't one. */
function detectList(para: ParagraphOptions, ctx: ResolveContext): ListInfo | null {
  const p = para as unknown as Record<string, unknown>;
  const numbering = p.numbering as { reference?: string; level?: number } | undefined;
  const bullet = p.bullet as { level?: number } | undefined;

  let kind: "bullet" | "ordered";
  let level: number;
  let reference: string | undefined;
  let start: number | undefined;

  if (numbering) {
    reference = numbering.reference;
    level = numbering.level ?? 0;
    const cfg = reference ? ctx.numberingLookup?.get(reference) : undefined;
    // A config whose format isn't "bullet" → ordered; otherwise this is the
    // built-in default-bullet numbering (parse may tag numId=1 as numbering
    // when its abstractNum resolves), so degrade to bullet.
    if (cfg && cfg.format && cfg.format !== "bullet") {
      kind = "ordered";
      start = cfg.start;
    } else {
      kind = "bullet";
      // Keep the source reference: a custom bullet abstractNum (e.g. a
      // Wingdings glyph) needs its original definition to round-trip the
      // marker; buildListTree carries it on the list node for compile.
    }
  } else if (bullet) {
    kind = "bullet";
    level = bullet.level ?? 0;
  } else {
    return null;
  }

  // Task items carry a leading inline checkbox SDT tagged "docen-task".
  const first = (p.children as unknown[] | undefined)?.[0];
  const isTask = taskItemExt.isTaskCheckbox(first);

  return {
    kind: isTask ? "task" : kind,
    level,
    reference,
    start,
    checked: taskItemExt.readCheckboxState(first),
  };
}

/** Resolve a list-item paragraph to a Tiptap paragraph/heading node, stripping
 *  the list marker (bullet/numbering) and the leading task checkbox — those are
 *  expressed at the list/item level, not inside the paragraph. */
function resolveListItemParagraph(
  para: ParagraphOptions,
  info: ListInfo,
  ctx: ResolveContext,
): JSONContent {
  const resolved = typeof para === "string" ? ({ text: para } as ParagraphOptions) : para;
  const headingLevel = detectHeadingLevel(resolved, ctx.styles);
  const nodeType = headingLevel ? "heading" : "paragraph";
  // Task: drop the leading checkbox SDT (its state lives in taskItem.attrs).
  // attrs still come from the original `resolved`; only the content source is
  // the stripped paragraph (buildTextBlock's contentPara override).
  const stripped = info.kind === "task" ? stripTaskCheckbox(resolved) : resolved;
  return buildTextBlock(nodeType, resolved, ctx, headingLevel, stripped);
}

/** Return a copy of `para` with its leading docen-task checkbox SDT removed. */
function stripTaskCheckbox(para: ParagraphOptions): ParagraphOptions {
  const children = (para as unknown as Record<string, unknown>).children;
  if (Array.isArray(children) && children.length > 0 && taskItemExt.isTaskCheckbox(children[0])) {
    return { ...(para as object), children: children.slice(1) } as ParagraphOptions;
  }
  return para;
}

/**
 * Rebuild nested Tiptap lists from a flat run of list paragraphs. Stack-based:
 * each frame is an active list at a given depth; the `key` (level:type:
 * reference) decides whether a paragraph continues the top list, starts a nested
 * list, or splits off a new sibling list.
 */
function buildListTree(
  group: { para: ParagraphOptions; info: ListInfo }[],
  ctx: ResolveContext,
): JSONContent[] {
  const topLevel: JSONContent[] = [];
  const stack: {
    level: number;
    key: string;
    listNode: JSONContent;
    currentItem: JSONContent;
  }[] = [];

  for (const { para, info } of group) {
    const listType =
      info.kind === "ordered" ? "orderedList" : info.kind === "task" ? "taskList" : "bulletList";
    const itemType = info.kind === "task" ? "taskItem" : "listItem";
    const key = `${info.level}:${listType}:${info.reference ?? ""}`;

    // Pop frames that are deeper than this item, or at the same depth but a
    // different list (level/type/reference change → new list).
    while (stack.length > 0) {
      const top = stack[stack.length - 1];
      if (top.level > info.level || (top.level === info.level && top.key !== key)) {
        stack.pop();
        continue;
      }
      break;
    }

    const itemPara = resolveListItemParagraph(para, info, ctx);
    const newItem: JSONContent = { type: itemType, content: [itemPara] };
    if (itemType === "taskItem") newItem.attrs = { checked: info.checked };

    const top = stack[stack.length - 1];
    if (top && top.level === info.level && top.key === key) {
      // Same list continues — append a new item.
      (top.listNode.content as JSONContent[]).push(newItem);
      top.currentItem = newItem;
    } else {
      // New list (top-level or nested under the current item).
      const newList: JSONContent = { type: listType, content: [newItem] };
      const listAttrs: Record<string, unknown> = {};
      // Only level-0 ordered lists carry `start`; deeper levels restart at 1.
      if (
        listType === "orderedList" &&
        info.level === 0 &&
        typeof info.start === "number" &&
        info.start !== 1
      ) {
        listAttrs.start = info.start;
      }
      // Carry the source abstractNum reference so the marker round-trips.
      if (info.reference) listAttrs.numbering = info.reference;
      if (Object.keys(listAttrs).length > 0) newList.attrs = listAttrs;
      if (top) {
        (top.currentItem.content as JSONContent[]).push(newList);
      } else {
        topLevel.push(newList);
      }
      stack.push({ level: info.level, key, listNode: newList, currentItem: newItem });
    }
  }

  return topLevel;
}

// DOCX list run (numbering/bullet paragraphs) → nested bulletList/orderedList/
// taskList tree. belongs classifies a list item; build re-derives each item's
// ListInfo and rebuilds the tree (filter guards a belongs/build mismatch).
export const parseDocxAggregator: ParseAggregatorRule = {
  belongs: (para, ctx) => detectList(para, ctx) != null,
  build: (group, ctx) => {
    const items = group
      .map((para) => ({ para, info: detectList(para, ctx) }))
      .filter((x): x is { para: ParagraphOptions; info: ListInfo } => x.info != null);
    return buildListTree(items, ctx);
  },
};

export const ListAggregator = Extension.create({ name: "listAggregator", parseDocxAggregator });
