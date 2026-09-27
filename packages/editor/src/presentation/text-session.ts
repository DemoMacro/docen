import type { SlideChild, TableCellOptions } from "@docen/pptx";

import { cellTextOf, shapeTextOf, writeCellText, writeShapeText } from "./commands";

/** A shape's text body source, or a table cell source. Both project through
 *  the same paragraph/run model and accept the same plain-text commit. */
export type TextEditSource =
  | { kind: "shape"; child: Extract<SlideChild, { shape: unknown }>["shape"] }
  | { kind: "cell"; cell: TableCellOptions };

/** Plain text currently painted by the source; null marks a non-text target. */
export function textOf(source: TextEditSource): string {
  return source.kind === "shape" ? (shapeTextOf(source.child) ?? "") : cellTextOf(source.cell);
}

/** Commit one plain-text session. The command helpers preserve the first
 *  paragraph run so editing text does not discard its base formatting. */
export function writeText(source: TextEditSource, text: string): void {
  if (source.kind === "shape") writeShapeText(source.child, text);
  else writeCellText(source.cell, text);
}
