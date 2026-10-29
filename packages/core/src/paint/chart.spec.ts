import type { LayoutDrawingMember } from "@docen/layout";
// @vitest-environment happy-dom
import type { IGroup } from "leafer-ui";
import { describe, expect, it } from "vitest";

type PaintedPath = { path?: string; stroke?: string };
type PaintedText = { text?: unknown };
type PaintedPoint = { x?: number };

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

const { Ellipse, Group, Path, Text } = await import("leafer-ui");
const { paintChartMember } = await import("./chart");

function nodes(tree: IGroup): unknown[] {
  return [tree, ...(tree.children ?? []).flatMap((child) => nodes(child as IGroup))];
}

function paths(tree: IGroup): PaintedPath[] {
  return nodes(tree).filter((node): node is PaintedPath => node instanceof Path);
}

function textOf(tree: IGroup): string[] {
  return nodes(tree)
    .filter((node): node is PaintedText => node instanceof Text)
    .map((node) => node.text)
    .filter((text): text is string => typeof text === "string");
}

function member(chart: unknown): Extract<LayoutDrawingMember, { kind: "chart" }> {
  return { kind: "chart", x: 0, y: 0, width: 400, height: 200, chart };
}

describe("paintChartMember", () => {
  it("keeps indexed blank values in their category slots", () => {
    const tree = new Group();
    paintChartMember(
      tree,
      member({
        type: "line",
        categories: ["A", "B", "C"],
        series: [{ values: [null, { index: 1, value: 4 }, { index: 2, value: 6 }] }],
      }),
    );
    const series = paths(tree).find((path) => path.stroke === "#70AD47");
    expect(series?.path).toBe("M 233 71.71428571428572 L 337.6666666666667 34.571428571428584");
  });

  it("reads indexed categories and series names", () => {
    const tree = new Group();
    paintChartMember(
      tree,
      member({
        type: "line",
        categories: [
          { index: 0, text: "First" },
          { index: 1, text: "Second" },
        ],
        series: [
          { name: { index: 0, text: "Revenue" }, values: [1, 2] },
          { name: { index: 1, text: "Cost" }, values: [2, 1] },
        ],
      }),
    );
    const labels = textOf(tree);
    expect(labels).toContain("First");
    expect(labels).toContain("Second");
    expect(labels).toContain("Revenue");
    expect(labels).toContain("Cost");
  });

  it("aligns scatter values without dropping blank slots", () => {
    const tree = new Group();
    paintChartMember(
      tree,
      member({
        type: "scatter",
        series: [
          {
            xValues: [
              { index: 0, value: 0 },
              { index: 1, value: 2 },
            ],
            values: [null, { index: 1, value: 3 }],
          },
        ],
      }),
    );
    const points = nodes(tree).filter((node): node is PaintedPoint => node instanceof Ellipse);
    expect(points).toHaveLength(1);
    expect(points[0]?.x).toBeCloseTo(357.95454545454544);
  });
});
