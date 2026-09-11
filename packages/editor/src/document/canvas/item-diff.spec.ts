import type { FlowItem } from "@docen/layout";
import { describe, expect, it } from "vitest";

import { diffFlowItems } from "./item-diff";

const item = (yPx: number, block: unknown, xPx?: number): FlowItem => ({
  yPx,
  block: block as FlowItem["block"],
  ...(xPx != null ? { xPx } : {}),
});

describe("diffFlowItems", () => {
  it("returns null without a baseline or when counts diverge", () => {
    const next = [item(0, { t: "a" })];
    expect(diffFlowItems(undefined, next)).toBeNull();
    expect(diffFlowItems([], next)).toBeNull();
    expect(diffFlowItems([item(0, { t: "a" }), item(10, { t: "b" })], next)).toBeNull();
  });

  it("keeps every item whose block is deep-equal, whatever the y drift", () => {
    const prev = [item(0, { t: "a" }), item(10, { t: "b" })];
    // yPx moved down (content reflowed above) — keep items only translate.
    const next = [item(24, { t: "a" }), item(40, { t: "b" })];
    expect(diffFlowItems(prev, next)).toEqual([
      { kind: "keep", index: 0 },
      { kind: "keep", index: 1 },
    ]);
  });

  it("repaints an item whose block changed", () => {
    const prev = [item(0, { t: "a" }), item(10, { t: "b" })];
    const next = [item(0, { t: "a" }), item(10, { t: "b2" })];
    expect(diffFlowItems(prev, next)).toEqual([
      { kind: "keep", index: 0 },
      { kind: "repaint", index: 1 },
    ]);
  });

  it("treats a column-x change as a repaint even with an equal block", () => {
    const prev = [item(0, { t: "a" }, 0)];
    const next = [item(0, { t: "a" }, 300)];
    expect(diffFlowItems(prev, next)).toEqual([{ kind: "repaint", index: 0 }]);
  });

  it("equates an absent column x with 0", () => {
    const prev = [item(0, { t: "a" })];
    const next = [item(0, { t: "a" }, 0)];
    expect(diffFlowItems(prev, next)).toEqual([{ kind: "keep", index: 0 }]);
  });

  it("repaints a paragraph carrying floating drawings even when unchanged", () => {
    const floats = { kind: "paragraph", drawings: [{ behind: false }] };
    const prev = [item(0, floats), item(10, { t: "b" })];
    const next = [item(0, { ...floats, drawings: [{ behind: false }] }), item(10, { t: "b" })];
    expect(diffFlowItems(prev, next)).toEqual([
      { kind: "repaint", index: 0 },
      { kind: "keep", index: 1 },
    ]);
  });

  it("repaints a group whose nested paragraph carries floating drawings", () => {
    const group = {
      kind: "group",
      children: [{ yPx: 0, block: { kind: "paragraph", drawings: [{}] } }],
    };
    const prev = [item(0, group)];
    const next = [item(0, JSON.parse(JSON.stringify(group)))];
    expect(diffFlowItems(prev, next)).toEqual([{ kind: "repaint", index: 0 }]);
  });
});
