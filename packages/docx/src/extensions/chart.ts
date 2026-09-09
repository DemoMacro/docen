import type { ChartOptions, ParagraphChild } from "@office-open/docx";

import type { JSONContent } from "../core";
import { Node } from "../core";
import type { ParseInlineRule } from "./types";

/**
 * chart — inline atom carrying a chart drawing (a:graphic > c:chart). The
 * whole ChartOptions (the ChartSpaceOptions model office-open parses out of
 * the chart part, plus the anchor fields: transformation/floating/altText/)
 * rides attrs.chart verbatim; the renderer's chart painter draws it. The
 * engine node is UI-free; a chart has no editable text body, so the node
 * stays an atom.
 */

type ChartBranch = Extract<ParagraphChild, { chart: ChartOptions }>;

// ParagraphChild `{ chart: ChartOptions }` → chart node (verbatim attrs).
export const parseDocxInline: ParseInlineRule<ChartBranch> = {
  match: (child): child is ChartBranch => "chart" in child,
  convert: (child) => ({ type: "chart", attrs: { chart: child.chart } }),
};

// chart node → office-open ParagraphChild `{ chart: ChartOptions }`.
export function renderDocx(node: JSONContent): Record<string, unknown> | null {
  const chart = (node.attrs as Record<string, unknown> | undefined)?.chart;
  return chart ? { chart } : null;
}

export const Chart = Node.create({
  name: "chart",
  inline: true,
  group: "inline",
  atom: true,
  draggable: true,

  addAttributes() {
    return {
      chart: {
        default: null,
        rendered: false,
      },
    };
  },

  parseHTML() {
    return [{ tag: "div[data-chart]" }];
  },

  renderDocx,
  parseDocxInline,
});
