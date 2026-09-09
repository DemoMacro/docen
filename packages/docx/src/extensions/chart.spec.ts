import type { DocumentOptions, ParagraphChild } from "@office-open/docx";
import { generateDocumentSync, parseDocumentSync } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { compileDocument, docxExtensions, resolveDocument } from "../index";

/**
 * chart — the inline atom carries the whole ChartOptions (ChartSpaceOptions
 * model + anchor fields) verbatim on attrs; resolve/compile are identity
 * passes around it, so the chart data survives the editor round-trip
 * byte-faithfully and the renderer's chart painter consumes it directly.
 */

function roundTrip(children: ParagraphChild[]) {
  const doc: DocumentOptions = {
    sections: [{ children: [{ paragraph: { children } }] }],
  };
  const json = resolveDocument(doc, docxExtensions);
  const compiled = compileDocument(json, docxExtensions);
  const paragraph = compiled.sections[0]!.children[0] as {
    paragraph: { children?: ParagraphChild[] };
  };
  return {
    json,
    node: (
      json.content?.[0] as { content?: { type: string; attrs?: object }[] } | undefined
    )?.content?.find((n) => n.type === "chart"),
    child: paragraph.paragraph.children?.find((c) => "chart" in c),
  };
}

describe("chart", () => {
  it("carries the chart model verbatim through resolve and compile", () => {
    const chartOptions = {
      type: "column" as const,
      title: "Quarterly",
      categories: ["Q1", "Q2", "Q3", "Q4"],
      series: [{ name: "Sales", values: [120, 90, 140, 160] }],
      transformation: { width: 4572000, height: 2743200 },
    };
    const { node, child } = roundTrip([{ chart: chartOptions }]);
    expect(node?.type).toBe("chart");
    expect((node?.attrs as { chart?: unknown } | undefined)?.chart).toEqual(chartOptions);
    expect(child).toEqual({ chart: chartOptions });
  });

  it("keeps floating and series detail fields intact", () => {
    const chartOptions = {
      type: "pie" as const,
      categories: ["A", "B"],
      series: [
        {
          name: "Share",
          values: [60, 40],
          dataLabels: { showPercent: true },
        },
      ],
      holeSize: 50,
      legendPosition: "right" as const,
      showLegend: true,
      transformation: { width: 3048000, height: 2286000 },
    };
    const { child } = roundTrip([{ chart: chartOptions }]);
    expect(child).toEqual({ chart: chartOptions });
  });

  it("surfaces as the chart node from a real OPC package round-trip", () => {
    const chartOptions = {
      type: "column" as const,
      title: "Quarterly",
      categories: ["Q1", "Q2"],
      series: [{ name: "Sales", values: [120, 90] }],
      transformation: { width: 4572000, height: 2743200 },
    };
    const binary = generateDocumentSync({
      sections: [{ children: [{ paragraph: { children: [{ chart: chartOptions }] } }] }],
    });
    const parsed = parseDocumentSync(new Uint8Array(binary as Buffer));
    const json = resolveDocument(parsed, docxExtensions);
    const types: string[] = [];
    const walk = (n: { type?: string; content?: unknown[] }): void => {
      if (n.type) types.push(n.type);
      for (const c of n.content ?? []) walk(c as { type?: string; content?: unknown[] });
    };
    walk(json);
    expect(types).toContain("chart");
    // The chart branch rides the chart node (not the passthrough atom).
    expect(types).not.toContain("inlinePassthrough");
  });
});
