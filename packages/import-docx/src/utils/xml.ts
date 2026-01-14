import type { Element, Parent } from "xast";

/**
 * XML element traversal utilities
 */

/**
 * Find first child element with given name
 */
export function findChild(element: Parent, name: string): Element | undefined {
  for (const child of element.children) {
    if (child.type === "element" && child.name === name) {
      return child;
    }
  }
  return undefined;
}

/**
 * Find descendant element with given name (depth-first search)
 */
export function findDeepChild(element: Element, name: string): Element | undefined {
  for (const child of element.children) {
    if (child.type === "element" && child.name === name) {
      return child;
    }

    if (child.type === "element") {
      const found = findDeepChild(child, name);
      if (found) return found;
    }
  }

  return undefined;
}

/**
 * Find all descendant elements with given name
 */
export function findDeepChildren(element: Element, name: string): Element[] {
  const results: Element[] = [];

  for (const child of element.children) {
    if (child.type === "element" && child.name === name) {
      results.push(child);
    }

    if (child.type === "element") {
      results.push(...findDeepChildren(child, name));
    }
  }

  return results;
}
