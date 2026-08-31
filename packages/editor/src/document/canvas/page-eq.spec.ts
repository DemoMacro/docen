import type { FlowPage } from "@docen/layout";
import { describe, expect, it } from "vitest";

import { deepEq, dirtyPagesOf } from "./page-eq";

const page = (...blocks: unknown[]): FlowPage => ({
  items: blocks.map((block, i) => ({ yPx: i * 10, block: block as never })),
});

describe("deepEq", () => {
  it("short-circuits identical references", () => {
    const shared = { a: [1, 2] };
    expect(deepEq(shared, shared)).toBe(true);
  });

  it("compares primitives, arrays, and nested objects", () => {
    expect(deepEq(1, 1)).toBe(true);
    expect(deepEq("a", "a")).toBe(true);
    expect(deepEq("a", "b")).toBe(false);
    expect(deepEq([1, { x: 2 }], [1, { x: 2 }])).toBe(true);
    expect(deepEq([1, { x: 2 }], [1, { x: 3 }])).toBe(false);
    expect(deepEq({ a: 1, b: 2 }, { b: 2, a: 1 })).toBe(true);
    expect(deepEq({ a: 1 }, { a: 1, b: 2 })).toBe(false);
    expect(deepEq(null, null)).toBe(true);
    expect(deepEq(null, {})).toBe(false);
  });
});

describe("dirtyPagesOf", () => {
  it("marks everything dirty without a baseline", () => {
    expect(dirtyPagesOf(undefined, [page({ t: "a" }), page({ t: "b" })])).toEqual([true, true]);
  });

  it("keeps unchanged pages and marks edited ones", () => {
    const prev = [page({ t: "a" }), page({ t: "b" }), page({ t: "c" })];
    // Same middle text, new string instance; page 0 text edited.
    const next = [page({ t: "a2" }), page({ t: "b" }), page({ t: "c" })];
    expect(dirtyPagesOf(prev, next)).toEqual([true, false, false]);
  });

  it("marks a page whose placement moved (yPx) even with equal content", () => {
    const prev = [{ items: [{ yPx: 0, block: {} as never }] }];
    const next = [{ items: [{ yPx: 5, block: {} as never }] }];
    expect(dirtyPagesOf(prev, next)).toEqual([true]);
  });

  it("marks trailing pages dirty when the page count grows or shrinks", () => {
    const two = [page({ t: "a" }), page({ t: "b" })];
    expect(dirtyPagesOf(two, [page({ t: "a" }), page({ t: "b" }), page({ t: "c" })])).toEqual([
      false,
      false,
      true,
    ]);
    expect(dirtyPagesOf(two, [page({ t: "a" })])).toEqual([false]);
  });
});
