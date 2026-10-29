import type { LayoutDrawingMember } from "@docen/layout";
// @vitest-environment happy-dom
import type { IGroup } from "leafer-ui";
import { describe, expect, it } from "vitest";

Object.defineProperty(globalThis, "CanvasRenderingContext2D", {
  value: class {},
  configurable: true,
});
Object.defineProperty(globalThis, "Path2D", {
  value: class {},
  configurable: true,
});
Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
  value: () => ({
    setTransform: () => undefined,
    measureText: (text: string) => ({ width: text.length * 7 }),
  }),
  configurable: true,
});

const { Group, Text } = await import("leafer-ui");
const { paintChartMember } = await import("./chart");

type PaintedText = { text?: unknown };

function nodes(tree: IGroup): unknown[] {
  return [tree, ...(tree.children ?? []).flatMap((child) => nodes(child as IGroup))];
}

function texts(tree: IGroup): string[] {
  return nodes(tree)
    .filter((node): node is PaintedText => node instanceof Text)
    .map((node) => node.text)
    .filter((text): text is string => typeof text === "string");
}

describe("chart data-label overrides", () => {
  it("renders custom, deleted, and point-formatted labels", () => {
    const tree = new Group();
    paintChartMember(tree, {
      kind: "chart",
      x: 0,
      y: 0,
      width: 400,
      height: 200,
      chart: {
        type: "column",
        categories: ["A", "B", "C"],
        series: [
          {
            values: [0.5, 2.34, 3],
            formatCode: "0.00",
            dataLabels: {
              showVal: true,
              labels: [
                { index: 0, text: { text: "Custom" } },
                { index: 1, delete: true },
                { index: 2, numberFormat: "0.0%" },
              ],
            },
          },
        ],
      },
    } satisfies Extract<LayoutDrawingMember, { kind: "chart" }>);
    const labels = texts(tree);

    expect(labels).toContain("Custom");
    expect(labels).not.toContain("2.34");
    expect(labels).toContain("300%");
  });
});
