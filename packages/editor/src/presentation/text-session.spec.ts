import type { SlideChild, TableCellOptions } from "@docen/pptx";
import { describe, expect, it } from "vitest";

import { textOf, writeText, type TextEditSource } from "./text-session";

describe("TextEditSource", () => {
  it("reads and writes shape text without dropping its first-run style", () => {
    const shape = {
      textBody: {
        paragraphs: [{ children: [{ text: "before", bold: true }] }],
      },
    } as unknown as Extract<SlideChild, { shape: unknown }>["shape"];
    const source: TextEditSource = { kind: "shape", child: shape };
    expect(textOf(source)).toBe("before");
    writeText(source, "after");
    expect(textOf(source)).toBe("after");
    const paragraph = shape.textBody?.paragraphs?.[0];
    const run = typeof paragraph === "object" ? paragraph.children?.[0] : undefined;
    expect(run).toMatchObject({
      text: "after",
      bold: true,
    });
  });

  it("reads and writes cell text", () => {
    const cell = { children: ["before"] } as TableCellOptions;
    const source: TextEditSource = { kind: "cell", cell };
    expect(textOf(source)).toBe("before");
    writeText(source, "after");
    expect(textOf(source)).toBe("after");
  });
});
