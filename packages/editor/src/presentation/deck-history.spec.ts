import { describe, expect, it } from "vitest";

import { DeckHistory } from "./deck-history";

describe("DeckHistory", () => {
  it("runs undo and redo in commit order", () => {
    const history = new DeckHistory();
    const calls: string[] = [];
    history.push({ undo: () => calls.push("a-"), redo: () => calls.push("a+") });
    history.push({ undo: () => calls.push("b-"), redo: () => calls.push("b+") });
    history.undo();
    history.undo();
    history.redo();
    expect(calls).toEqual(["b-", "a-", "a+"]);
    expect(history.canUndo).toBe(true);
    expect(history.canRedo).toBe(true);
  });

  it("truncates the redo branch and honors its limit", () => {
    const history = new DeckHistory(2);
    history.push({ undo: () => {}, redo: () => {} });
    history.push({ undo: () => {}, redo: () => {} });
    history.undo();
    history.push({ undo: () => {}, redo: () => {} });
    expect(history.canUndo).toBe(true);
    expect(history.canRedo).toBe(false);
    history.push({ undo: () => {}, redo: () => {} });
    history.undo();
    history.undo();
    expect(history.canUndo).toBe(false);
  });
});
