/**
 * HTML → Tiptap JSON through the docx schema's paste rules.
 *
 * The docx schema has no DOM rendering path (the canvas paints); parseHTML
 * exists for exactly one job: turning pasted/fetched HTML into document JSON.
 * This helper is DOM-provider agnostic — the editor passes a native
 * DOMParser body, specs pass a linkedom body — so the mapping rules have one
 * implementation and two DOM sources.
 *
 * @module
 */

import { DOMParser as ProseMirrorDOMParser, type Schema } from "@tiptap/pm/model";

import type { JSONContent } from "../core";
import { assignOrderedReferences } from "./list-numbering";

/**
 * Flatten nested ul/ol structure into sibling <li> elements carrying
 * data-docen-level / data-docen-kind. The schema's list representation is a
 * FLAT paragraph with a level attr (no list tree), and a paragraph rule cannot
 * host the nested block content of li>ul>li — the ProseMirror DOMParser
 * collapses such structure instead of splitting it. Rewriting the DOM first
 * keeps every nested item as its own paragraph; the Paragraph li rule reads
 * the two attributes back (nesting depth + the nearest list's kind, which
 * decides bullet vs ordered).
 */
function flattenListNesting(root: HTMLElement): void {
  const topLevel = [...root.querySelectorAll("ul, ol")].filter(
    (list) => !list.parentElement?.closest("ul, ol"),
  );
  for (const list of topLevel) {
    const items: HTMLElement[] = [];
    const walk = (parent: Element, level: number, ordered: boolean): void => {
      for (const li of parent.children) {
        if (li.tagName.toUpperCase() !== "LI") continue;
        const nested: Element[] = [];
        for (const child of li.children) {
          const tag = child.tagName.toUpperCase();
          if (tag === "UL" || tag === "OL") nested.push(child);
        }
        for (const child of nested) child.remove();
        li.setAttribute("data-docen-level", String(level));
        li.setAttribute("data-docen-kind", ordered ? "ol" : "ul");
        items.push(li as HTMLElement);
        for (const child of nested) {
          walk(child, level + 1, child.tagName.toUpperCase() === "OL");
        }
      }
    };
    walk(list, 0, list.tagName.toUpperCase() === "OL");
    const frag = root.ownerDocument?.createDocumentFragment();
    if (!frag) continue;
    for (const li of items) frag.appendChild(li);
    list.replaceWith(frag);
  }
}

/**
 * Parse a parsed-HTML `<body>` into Tiptap JSON with the given schema's
 * parseDOM rules (the extensions' parseHTML). Ordered-list items carried as
 * `<ol><li>` get fresh generated numbering references — one per consecutive
 * run — via assignOrderedReferences, matching what markdown list parsing
 * produces.
 */
export function parseHTMLBody(body: HTMLElement, schema: Schema): JSONContent {
  flattenListNesting(body);
  const doc = ProseMirrorDOMParser.fromSchema(schema).parse(body);
  return assignOrderedReferences(doc.toJSON());
}
