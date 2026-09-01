import { describe, expect, it } from "vitest";

import { generateDOCXSync, parseDOCX, type JSONContent } from "../index";

// w:highlight val is the fixed ST_HighlightColor enumeration — a hex value
// there makes Word refuse the file. The Highlight mark's renderDocx maps every
// attr form onto a legal encoding: palette tokens and palette RGBs emit
// w:highlight tokens, off-palette colors ride as character shading (w:shd).

function docWithHighlight(color: string): JSONContent {
  return {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "probe",
            marks: [{ type: "highlight", attrs: { color } }],
          },
        ],
      },
    ],
  };
}

/** The first run's highlight-mark color (undefined = no highlight mark). */
function highlightColor(json: JSONContent): string | undefined {
  for (const child of json.content ?? []) {
    if (child.type === "text") {
      const mark = (child.marks ?? []).find((m) => m.type === "highlight");
      if (mark) return mark.attrs?.color as string;
    }
    const hit = highlightColor(child);
    if (hit !== undefined) return hit;
  }
  return undefined;
}

/** The first run's textStyle shading attr (the off-palette carrier). */
function shadingFill(json: JSONContent): unknown {
  for (const child of json.content ?? []) {
    if (child.type === "text") {
      const mark = (child.marks ?? []).find((m) => m.type === "textStyle");
      if (mark?.attrs?.shading) return (mark.attrs.shading as Record<string, unknown>).fill;
    }
    const hit = shadingFill(child);
    if (hit !== undefined) return hit;
  }
  return undefined;
}

/** The first run's textStyle highlight attr (the "none" cancel carrier). */
function textStyleHighlight(json: JSONContent): unknown {
  for (const child of json.content ?? []) {
    if (child.type === "text") {
      const mark = (child.marks ?? []).find((m) => m.type === "textStyle");
      if (mark) return mark.attrs?.highlight;
    }
    const hit = textStyleHighlight(child);
    if (hit !== undefined) return hit;
  }
  return undefined;
}

describe("highlight mark export mapping", () => {
  it("emits a palette token and parses back as the mark", () => {
    const json = parseDOCX(generateDOCXSync(docWithHighlight("yellow")) as Uint8Array);
    expect(highlightColor(json)).toBe("yellow");
  });

  it("maps a palette RGB (pasted #hex form) back onto its token", () => {
    const json = parseDOCX(generateDOCXSync(docWithHighlight("FFFF00")) as Uint8Array);
    expect(highlightColor(json)).toBe("yellow");
  });

  it("maps the CSS rgb() form pasted HTML carries onto its token", () => {
    const json = parseDOCX(generateDOCXSync(docWithHighlight("rgb(255, 255, 0)")) as Uint8Array);
    expect(highlightColor(json)).toBe("yellow");
  });

  it("routes an off-palette color through character shading, never w:highlight", () => {
    const json = parseDOCX(generateDOCXSync(docWithHighlight("00CCFF")) as Uint8Array);
    expect(highlightColor(json)).toBeUndefined();
    expect(shadingFill(json)).toBe("00CCFF");
  });

  it("round-trips highlight none through TextStyle without a mark", () => {
    // <w:highlight w:val="none"/> cancels an inherited highlight — TextStyle
    // carries it verbatim; the mark must stay off (same class as underline
    // val="none").
    const json = parseDOCX(
      generateDOCXSync({
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "text",
                text: "probe",
                marks: [{ type: "textStyle", attrs: { highlight: "none" } }],
              },
            ],
          },
        ],
      }) as Uint8Array,
    );
    expect(highlightColor(json)).toBeUndefined();
    expect(textStyleHighlight(json)).toBe("none");
  });
});
