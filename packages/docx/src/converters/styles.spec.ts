import type { StylesOptions } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { quickStyles } from "./styles";

/**
 * Quick Styles gallery entries: the `run` each entry carries is the style's
 * *effective* character formatting — the basedOn chain merged with the
 * document defaults filling the gaps — so a preview renders the style the way
 * the document does, not its raw (often sparse) entry.
 */

const stylesWith = (paragraphStyles: StylesOptions["paragraphStyles"]): StylesOptions => ({
  default: { document: { run: { size: 11 } } },
  paragraphStyles,
});

describe("quickStyles effective run", () => {
  it("merges the basedOn chain and fills the gaps from docDefaults", () => {
    const entries = quickStyles(
      stylesWith([
        { id: "Normal", name: "Normal", run: { font: "Calibri" } },
        { id: "Heading1", name: "heading 1", basedOn: "Normal", run: { size: 16, bold: true } },
      ]),
    );
    const h1 = entries.find((e) => e.id === "Heading1");
    // size from the style itself, font inherited from Normal, both present in
    // one preview-ready run.
    expect(h1?.run).toEqual({ size: 16, bold: true, font: "Calibri" });
    const normal = entries.find((e) => e.id === "Normal");
    expect(normal?.run).toEqual({ font: "Calibri", size: 11 });
  });

  it("lists only quickFormat styles by uiPriority, falling back to all", () => {
    const flagged = quickStyles(
      stylesWith([
        { id: "B", name: "B", quickFormat: true, uiPriority: 20 },
        { id: "A", name: "A", quickFormat: true, uiPriority: 10 },
        { id: "C", name: "C" },
      ]),
    );
    expect(flagged.map((e) => e.id)).toEqual(["A", "B"]);

    const unflagged = quickStyles(
      stylesWith([
        { id: "B", name: "B" },
        { id: "A", name: "A" },
      ]),
    );
    expect(unflagged.map((e) => e.id)).toEqual(["B", "A"]);
  });

  it("returns no entries for a styles-less document", () => {
    expect(quickStyles(null)).toEqual([]);
  });
});
