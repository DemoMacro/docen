import type { DocumentOptions, SectionChild } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { compileDocument, docxExtensions, resolveDocument } from "../index";

/**
 * VML text box (w:pict > v:shape > v:textbox) round-trips: the box's own data
 * (structured VML style + residual paragraph options) rides attrs verbatim on
 * both legs; the contained block stream resolves to editable nodes and
 * compiles back through the shared block walk.
 */

function roundTrip(children: SectionChild[]) {
  const doc: DocumentOptions = { sections: [{ children }] };
  const json = resolveDocument(doc, docxExtensions);
  const compiled = compileDocument(json, docxExtensions);
  return { json, child: compiled.sections[0].children[0] };
}

describe("textbox", () => {
  it("carries the box data verbatim and compiles content back", () => {
    const style = { width: "200pt", height: "100pt" } as const;
    const { json, child } = roundTrip([
      {
        textbox: {
          text: "caption",
          style,
          children: [{ paragraph: { text: "in-box" } }, { paragraph: { text: "more" } }],
        },
      },
    ]);
    const node = json.content?.[0] as {
      type: string;
      attrs?: { textbox?: unknown };
      content?: { type: string }[];
    };
    expect(node.type).toBe("textbox");
    expect(node.attrs?.textbox).toEqual({ text: "caption", style });
    expect(node.content?.map((c) => c.type)).toEqual(["paragraph", "paragraph"]);
    expect(child).toEqual({
      textbox: {
        text: "caption",
        style,
        children: [{ paragraph: "in-box" }, { paragraph: "more" }],
      },
    });
  });

  it("resolves a nested table through the block walk", () => {
    const { json, child } = roundTrip([
      {
        textbox: {
          style: { width: "300pt", height: "80pt" },
          children: [
            { paragraph: { text: "head" } },
            {
              table: {
                rows: [{ cells: [{ children: [{ paragraph: { text: "cell" } }] }] }],
              },
            },
          ],
        },
      },
    ]);
    const node = json.content?.[0] as { content?: { type: string }[] };
    expect(node.content?.map((c) => c.type)).toEqual(["paragraph", "table"]);
    const box = (
      child as {
        textbox: {
          children: {
            table: { rows: { cells: { children: { paragraph: string }[] }[] }[] };
          }[];
        };
      }
    ).textbox;
    // The table compiles with its projected widths re-added (shared table
    // behavior) — probe the structure, not the whole object.
    expect(box.children[1].table.rows[0].cells[0].children[0].paragraph).toBe("cell");
  });

  it("keeps an empty box resolvable with a placeholder paragraph", () => {
    const { json } = roundTrip([{ textbox: { style: { width: "100pt", height: "40pt" } } }]);
    const node = json.content?.[0] as { type: string; content?: unknown[] };
    expect(node.type).toBe("textbox");
    expect(node.content?.length).toBeGreaterThan(0);
  });

  // NOTE: no byte-level generate→parse test here on purpose. office-open
  // stringifies a textbox as `w:p > w:r > w:pict` (run-level per
  // EG_RunInnerContent) while its parser only claims a pict that is a DIRECT
  // child of w:p, so a generated file reads back as a pict blob — the
  // textbox branch type also lacks the shape detail fields (filled,
  // insetmode, shape id) a lossless promotion needs. Symmetrizing that pair
  // is its own batch; this suite pins the resolve/compile contract only.
});
