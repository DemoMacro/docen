import { projectDocument } from "@docen/docx/layout";
import {
  fieldLabelOf,
  type FlowPage,
  type LaidOutParagraph,
  type LayoutInline,
} from "@docen/layout";
import type { DocumentOptions, SectionChild } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { liveFieldResolver, resolvePageFieldsBounded } from "./field-resolve";
import type { FieldContext } from "./fields";

// Regression: Alt+F9 (the field-code display) must paint the instruction for
// page-dependent fields too. The projection emits code atoms with no dynamic
// marker; the render resolve pass must leave them alone (no `resolved`, no
// measured-text rewrite, no re-layout), so the painted label is the code.

/** A synthetic laid paragraph over the projected inline atoms — the paint
 *  label reads the atom plus the measured item text, both preserved here. */
const para = (inline: readonly LayoutInline[]): LaidOutParagraph => ({
  kind: "paragraph",
  heightPx: 20,
  beforePx: 0,
  afterPx: 0,
  inline,
  lines: [
    {
      yPx: 0,
      heightPx: 20,
      naturalPx: 20,
      endInlineIndex: inline.length - 1,
      items: inline.map((atom, index) => ({
        kind: "text" as const,
        inlineIndex: index,
        text: atom.kind === "text" ? atom.text : "",
        xPx: index * 10,
        widthPx: 10,
      })),
    },
  ],
});

const pageOf = (inline: readonly LayoutInline[]): FlowPage => ({
  items: [{ yPx: 0, block: para(inline) }],
});

const projectedAtoms = (showFieldCodes: boolean): LayoutInline[] => {
  const children: SectionChild[] = [
    {
      paragraph: {
        children: [
          { complexField: { instruction: "PAGE \\* MERGEFORMAT", result: "3" } },
          { complexField: { instruction: "NUMPAGES", result: "9" } },
          { simpleField: { instruction: "SECTION", cachedValue: "2" } },
          { simpleField: { instruction: "SECTIONPAGES", cachedValue: "4" } },
          { simpleField: { instruction: "AUTHOR", cachedValue: "作者" } },
        ],
      },
    },
  ];
  const doc: DocumentOptions = { sections: [{ children }] };
  const blocks = projectDocument(doc, undefined, showFieldCodes).sections[0]!.blocks;
  const block = blocks[0];
  if (block?.kind !== "paragraph") throw new Error("expected paragraph");
  return [...block.inline];
};

const base: Omit<FieldContext, "frame" | "sequences"> = { now: new Date(2026, 8, 9) };

/** The page frame a two-section document would hand the resolve pass (the
 *  field's own page: section 2 restarting at 5, lowercase Roman). */
const page = { pageIndex: 1, pageCount: 6, pageNumber: { fmt: "lowerRoman", offset: 4 } };

const labelsOf = (atoms: readonly LayoutInline[]): string[] =>
  atoms.map((atom) => {
    if (atom.kind !== "text") throw new Error("expected text atom");
    return fieldLabelOf(atom, page, atom.text);
  });

describe("field-code view (Alt+F9)", () => {
  it("paints every field's instruction and never resolves or re-lays", () => {
    const atoms = projectedAtoms(true);
    // The projection carries the code as the measured text, no dynamic marker.
    expect(atoms.map((atom) => (atom.kind === "text" ? atom.text : ""))).toEqual([
      "PAGE \\* MERGEFORMAT",
      "NUMPAGES",
      "SECTION",
      "SECTIONPAGES",
      "AUTHOR",
    ]);

    const pages = [pageOf(atoms)];
    let relayouts = 0;
    const result = resolvePageFieldsBounded(
      pages,
      [{ pageNumbering: { start: 5, format: "lowerRoman" } }],
      [0],
      liveFieldResolver(base, true),
      () => {
        relayouts += 1;
        return { pages, sectionOfPage: [0] };
      },
    );
    expect(result.passes).toBe(1);
    expect(relayouts).toBe(0);
    expect(result.dirty).toEqual([]);
    for (const atom of atoms) {
      if (atom.kind !== "text") continue;
      expect(atom.resolved).toBeUndefined();
      expect(atom.field).toBeUndefined();
    }
    expect(labelsOf(atoms)).toEqual([
      "PAGE \\* MERGEFORMAT",
      "NUMPAGES",
      "SECTION",
      "SECTIONPAGES",
      "AUTHOR",
    ]);
  });

  it("resolves and paints live numbers again when the code view turns off", () => {
    const atoms = projectedAtoms(false);
    let pages = [pageOf(atoms)];
    let relayouts = 0;
    // SECTION/SECTIONPAGES rewrite their measured text on the first pass, so
    // the bounded loop re-lays the same atoms and settles on the second —
    // one re-layout, then stable.
    const result = resolvePageFieldsBounded(
      pages,
      [{ pageNumbering: { start: 5, format: "lowerRoman" } }],
      [0],
      liveFieldResolver(base, false),
      () => {
        relayouts += 1;
        pages = [pageOf(atoms)];
        return { pages, sectionOfPage: [0] };
      },
    );
    expect(result.passes).toBe(2);
    expect(relayouts).toBe(1);
    // PAGE/NUMPAGES resolve against this page (restart 5, lowercase Roman);
    // SECTION/SECTIONPAGES against the single section.
    expect(atoms[0]).toMatchObject({ text: "0", field: "page", resolved: "v" });
    expect(atoms[1]).toMatchObject({ text: "0", field: "numPages", resolved: "1" });
    expect(atoms[2]).toMatchObject({ text: "i", resolved: "i" });
    expect(atoms[3]).toMatchObject({ text: "i", resolved: "i" });
    // AUTHOR is not a live field — its cached text stays.
    expect(atoms[4]).toMatchObject({ text: "作者" });
    expect((atoms[4] as { resolved?: string }).resolved).toBeUndefined();
    expect(labelsOf(atoms)).toEqual(["v", "1", "i", "i", "作者"]);
  });
});
