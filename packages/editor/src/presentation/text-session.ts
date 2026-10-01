import type { SlideChild, TableCellOptions } from "@docen/pptx";

import { formatTextModel, textModelOf, writeTextModel, type TextModelSource } from "./text-model";

/** A shape's text body source, or a table cell source. Both project through
 *  the same paragraph/run model and accept the same plain-text commit. */
export type TextEditSource =
  | { kind: "shape"; child: Extract<SlideChild, { shape: unknown }>["shape"] }
  | { kind: "cell"; cell: TableCellOptions };

/** Plain text currently painted by the source; null marks a non-text target. */
export function textOf(source: TextEditSource): string {
  return textModelOf(modelSourceOf(source));
}

/** Commit one plain-text session. The command helpers preserve the first
 *  paragraph run; the model writer preserves every other styled run. */
export function writeText(source: TextEditSource, text: string): void {
  const model = modelSourceOf(source);
  writeTextModel(model, text, textModelOf(model));
}

/** Apply a formatting command to the textarea selection. */
export function formatText(
  source: TextEditSource,
  start: number,
  end: number,
  name: string,
  value?: string,
): boolean {
  return formatTextModel(modelSourceOf(source), start, end, name, value);
}

function modelSourceOf(source: TextEditSource): TextModelSource {
  return source.kind === "shape" ? { kind: "shape", shape: source.child } : source;
}
