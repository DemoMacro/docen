import { describe, expect, it } from "vitest";

import { blockRuleOf, enterRuleOf, inlineRuleOf, isHyphenRun } from "./markdown-input";

describe("blockRuleOf", () => {
  it("maps # runs to heading levels 1-6", () => {
    expect(blockRuleOf(" ", "#")).toEqual({ kind: "heading", level: 1 });
    expect(blockRuleOf(" ", "######")).toEqual({ kind: "heading", level: 6 });
    expect(blockRuleOf(" ", "##")).toEqual({ kind: "heading", level: 2 });
  });

  it("rejects 7 hashes and anything with extra characters", () => {
    expect(blockRuleOf(" ", "#######")).toBeNull();
    expect(blockRuleOf(" ", "#a")).toBeNull();
    expect(blockRuleOf(" ", "# ")).toBeNull();
    expect(blockRuleOf(" ", "a#")).toBeNull();
  });

  it("maps > to a quote only as a lone character", () => {
    expect(blockRuleOf(" ", ">")).toEqual({ kind: "quote" });
    expect(blockRuleOf(" ", ">>")).toBeNull();
    expect(blockRuleOf(" ", "a>")).toBeNull();
  });

  it("maps - * + to bullets", () => {
    expect(blockRuleOf(" ", "-")).toEqual({ kind: "bullet" });
    expect(blockRuleOf(" ", "*")).toEqual({ kind: "bullet" });
    expect(blockRuleOf(" ", "+")).toEqual({ kind: "bullet" });
    expect(blockRuleOf(" ", "--")).toBeNull();
  });

  it("maps N. and N) to ordered lists starting at N", () => {
    expect(blockRuleOf(" ", "1.")).toEqual({ kind: "ordered", start: 1 });
    expect(blockRuleOf(" ", "3)")).toEqual({ kind: "ordered", start: 3 });
    expect(blockRuleOf(" ", "99.")).toEqual({ kind: "ordered", start: 99 });
    expect(blockRuleOf(" ", "0.")).toBeNull();
    expect(blockRuleOf(" ", "a.")).toBeNull();
    expect(blockRuleOf(" ", "1")).toBeNull();
  });

  it("only fires on a typed space", () => {
    expect(blockRuleOf("#", "")).toBeNull();
    expect(blockRuleOf("-", "-")).toBeNull();
    expect(blockRuleOf("x", "#")).toBeNull();
  });
});

describe("enterRuleOf", () => {
  it("maps a bare fence to a Code paragraph without a language", () => {
    expect(enterRuleOf("```")).toEqual({ kind: "code", language: null });
  });

  it("carries the fence info string as the language", () => {
    expect(enterRuleOf("```ts")).toEqual({ kind: "code", language: "ts" });
    expect(enterRuleOf("```typescript")).toEqual({ kind: "code", language: "typescript" });
  });

  it("rejects two backticks and multi-word info strings", () => {
    expect(enterRuleOf("``")).toBeNull();
    expect(enterRuleOf("```a b")).toBeNull();
    expect(enterRuleOf("```a b```")).toBeNull();
  });

  it("maps horizontal-rule lines of -, *, _", () => {
    expect(enterRuleOf("---")).toEqual({ kind: "hr" });
    expect(enterRuleOf("-----")).toEqual({ kind: "hr" });
    expect(enterRuleOf("***")).toEqual({ kind: "hr" });
    expect(enterRuleOf("___")).toEqual({ kind: "hr" });
  });

  it("rejects short runs and trailing characters", () => {
    expect(enterRuleOf("--")).toBeNull();
    expect(enterRuleOf("-_-")).toBeNull();
    expect(enterRuleOf("*** ")).toBeNull();
  });
});

describe("inlineRuleOf", () => {
  it("closes bold before italic (double delimiter wins)", () => {
    expect(inlineRuleOf("*", "x **a*")).toEqual({
      mark: "bold",
      openStart: 2,
      openLen: 2,
      innerLen: 1,
    });
    expect(inlineRuleOf("*", "**a*")).toEqual({
      mark: "bold",
      openStart: 0,
      openLen: 2,
      innerLen: 1,
    });
    expect(inlineRuleOf("_", "x __a_")).toEqual({
      mark: "bold",
      openStart: 2,
      openLen: 2,
      innerLen: 1,
    });
  });

  it("closes single-delimiter italic and measures multi-char inner runs", () => {
    expect(inlineRuleOf("*", "x *a")).toEqual({
      mark: "italic",
      openStart: 2,
      openLen: 1,
      innerLen: 1,
    });
    expect(inlineRuleOf("*", "x *em text")).toEqual({
      mark: "italic",
      openStart: 2,
      openLen: 1,
      innerLen: 7,
    });
    expect(inlineRuleOf("_", "x _a")).toEqual({
      mark: "italic",
      openStart: 2,
      openLen: 1,
      innerLen: 1,
    });
  });

  it("closes strike only on doubled tildes", () => {
    expect(inlineRuleOf("~", "x ~~a~")).toEqual({
      mark: "strike",
      openStart: 2,
      openLen: 2,
      innerLen: 1,
    });
    expect(inlineRuleOf("~", "x ~a")).toBeNull();
  });

  it("closes inline code on single backticks", () => {
    expect(inlineRuleOf("`", "x `c")).toEqual({
      mark: "code",
      openStart: 2,
      openLen: 1,
      innerLen: 1,
    });
  });

  it("keeps intraword delimiters literal (no start/whitespace behind)", () => {
    expect(inlineRuleOf("*", "a*b*")).toBeNull();
    expect(inlineRuleOf("_", "a_b_")).toBeNull();
  });

  it("refuses whitespace-bounded or empty inner runs", () => {
    expect(inlineRuleOf("*", "* a")).toBeNull();
    expect(inlineRuleOf("*", "*a ")).toBeNull();
    expect(inlineRuleOf("*", "***")).toBeNull();
  });

  it("ignores characters outside the delimiter set", () => {
    expect(inlineRuleOf("x", "x *a")).toBeNull();
    expect(inlineRuleOf(" ", "x *a")).toBeNull();
  });
});

describe("isHyphenRun", () => {
  it("matches a run of hyphens from the paragraph start", () => {
    expect(isHyphenRun("-")).toBe(true);
    expect(isHyphenRun("--")).toBe(true);
  });

  it("leaves empty and ordinary text to autocorrect", () => {
    expect(isHyphenRun("")).toBe(false);
    expect(isHyphenRun("a-")).toBe(false);
    expect(isHyphenRun("-a")).toBe(false);
    expect(isHyphenRun(" -")).toBe(false);
  });
});
