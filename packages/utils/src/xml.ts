/**
 * XML utility functions for DOCX processing
 * Provides helper functions for traversing and parsing XML trees (xast)
 */

import type { Element, Root } from "xast";

/**
 * Find direct child element with specified name
 * @param node - Parent XML element or root node
 * @param name - Child element name to find (can include namespace prefix, e.g., "w:p")
 * @returns Child element if found, null otherwise
 *
 * @example
 * const paragraph = findChild(document, "w:p");
 */
export function findChild(node: Root | Element, name: string): Element | null {
  if (!node.children) return null;

  for (const child of node.children) {
    if (child.type === "element" && child.name === name) {
      return child as Element;
    }
  }

  return null;
}

/**
 * Find deep descendant element with specified name (recursive)
 * Searches through all descendants, not just direct children
 * @param node - Root XML element
 * @param name - Descendant element name to find
 * @returns Descendant element if found, null otherwise
 *
 * @example
 * const textElement = findDeepChild(run, "w:t");
 */
export function findDeepChild(node: Root | Element, name: string): Element | null {
  if (!node.children) return null;

  for (const child of node.children) {
    if (child.type === "element") {
      if (child.name === name) {
        return child as Element;
      }

      // Recursively search in children
      const found = findDeepChild(child as Element, name);
      if (found) return found;
    }
  }

  return null;
}

/**
 * Find all deep descendant elements with specified name (recursive)
 * @param node - Root XML element
 * @param name - Descendant element name to find
 * @returns Array of matching descendant elements
 *
 * @example
 * const allTextRuns = findDeepChildren(paragraph, "w:r");
 */
export function findDeepChildren(node: Root | Element, name: string): Element[] {
  const results: Element[] = [];

  if (!node.children) return results;

  for (const child of node.children) {
    if (child.type === "element") {
      if (child.name === name) {
        results.push(child as Element);
      }

      // Recursively search in children
      results.push(...findDeepChildren(child as Element, name));
    }
  }

  return results;
}

/**
 * Parse TWIP attribute value from element attributes
 * TWIP = Twentieth of a Point (1 inch = 1440 TWIPs)
 * @param attributes - Element attributes object
 * @param name - Attribute name to parse
 * @returns TWIP value as string, or undefined if not found
 *
 * @example
 * const leftIndent = parseTwipAttr(pPr.attributes, "w:left");
 */
export function parseTwipAttr(
  attributes: Record<string, any> | { [key: string]: string | undefined },
  name: string,
): string | undefined {
  const value = attributes[name];
  if (!value) return undefined;

  // Validate it's a number
  const num = parseInt(value as string, 10);
  if (isNaN(num)) return undefined;

  return value as string;
}
