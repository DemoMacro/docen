// @vitest-environment happy-dom
import { describe, expect, it } from "vitest";

import {
  applyRecipientsRow,
  mergeFieldDisplay,
  mergeFieldName,
  mergeFieldSeed,
  parseRecipients,
} from "./mail-merge";

describe("parseRecipients", () => {
  it("parses a pasted TSV grid (the Excel clipboard shape)", () => {
    expect(parseRecipients("Name\tCity\n甲乙丙丁\t司南\n戊己庚辛\t度量衡")).toEqual({
      headers: ["Name", "City"],
      rows: [
        ["甲乙丙丁", "司南"],
        ["戊己庚辛", "度量衡"],
      ],
    });
  });

  it("parses CSV with quoted cells and doubled quotes", () => {
    expect(parseRecipients('Name,City\n"Smith, John","Rome ""eternal"""\nJo,Paris')).toEqual({
      headers: ["Name", "City"],
      rows: [
        ["Smith, John", 'Rome "eternal"'],
        ["Jo", "Paris"],
      ],
    });
  });

  it("pads short records and drops blank lines", () => {
    expect(parseRecipients("A,B\none\n\ntwo,x,y\n")).toEqual({
      headers: ["A", "B"],
      rows: [
        ["one", ""],
        ["two", "x"],
      ],
    });
  });

  it("needs a header row plus at least one record", () => {
    expect(parseRecipients("")).toBeNull();
    expect(parseRecipients("A,B")).toBeNull();
    expect(parseRecipients("A,B\n")).toBeNull();
    expect(parseRecipients(",\nx,y")).toBeNull();
  });
});

describe("merge fields", () => {
  it("reads the field name from Word's instruction shapes", () => {
    expect(mergeFieldName("MERGEFIELD Name")).toBe("Name");
    expect(mergeFieldName(" MERGEFIELD Name \\* MERGEFORMAT ")).toBe("Name");
    expect(mergeFieldName("PAGE")).toBeNull();
  });

  it("the seed round-trips through the field name", () => {
    const seed = mergeFieldSeed("Name");
    const data = JSON.parse(String((seed.attrs as { data?: string }).data)) as {
      simpleField: { instruction: string; cachedValue: string };
    };
    expect(data.simpleField.cachedValue).toBe(mergeFieldDisplay("Name"));
    expect(mergeFieldName(data.simpleField.instruction)).toBe("Name");
  });
});

describe("applyRecipientsRow", () => {
  const recipients = { headers: ["Name", "City"], rows: [["甲乙丙丁", "司南"]] };
  const doc = {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          mergeFieldSeed("Name"),
          { type: "text", text: " " },
          mergeFieldSeed("City"),
          mergeFieldSeed("Missing"),
        ],
      },
      {
        type: "inlinePassthrough",
        attrs: { data: "not json" },
      },
    ],
  };

  it("swaps the merge fields to the row's values and keeps unknown chevrons", () => {
    const merged = applyRecipientsRow(doc, recipients, 0);
    const paragraph = merged.content?.[0];
    const texts = (paragraph?.content ?? []).map(
      (node) =>
        (
          JSON.parse(String((node.attrs as { data?: string } | undefined)?.data ?? "{}")) as {
            simpleField?: { cachedValue?: string };
          }
        ).simpleField?.cachedValue ?? node.text,
    );
    expect(texts).toEqual(["甲乙丙丁", " ", "司南", "«Missing»"]);
    // The source document is untouched (the merge is a clone).
    const original = doc.content?.[0]?.content?.[0]?.attrs as { data?: string } | undefined;
    const source = JSON.parse(String(original?.data ?? "{}")) as {
      simpleField?: { cachedValue?: string };
    };
    expect(source.simpleField?.cachedValue).toBe(mergeFieldDisplay("Name"));
  });

  it("leaves opaque payloads and non-field nodes alone", () => {
    const merged = applyRecipientsRow(doc, recipients, 0);
    const opaque = merged.content?.[1]?.attrs as { data?: string } | undefined;
    expect(opaque?.data).toBe("not json");
  });

  it("matches field names case-insensitively", () => {
    const doc2 = { type: "doc", content: [mergeFieldSeed("name")] };
    const merged = applyRecipientsRow(doc2, recipients, 0);
    const field = merged.content?.[0]?.attrs as { data?: string } | undefined;
    const data = JSON.parse(String(field?.data ?? "{}")) as {
      simpleField: { cachedValue: string };
    };
    expect(data.simpleField.cachedValue).toBe("甲乙丙丁");
  });
});
