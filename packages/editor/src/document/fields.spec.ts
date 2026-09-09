import { describe, expect, it } from "vitest";

import { defaultInstruction, evaluateField, fieldRef, formatDate, instructionName } from "./fields";

describe("formatDate", () => {
  // Fixed date: 2026-03-09 is a Monday; 14:05:09.
  const d = new Date(2026, 2, 9, 14, 5, 9);

  it("expands the common Word picture tokens", () => {
    expect(formatDate(d, "yyyy/M/d")).toBe("2026/3/9");
    expect(formatDate(d, "yyyy年M月d日")).toBe("2026年3月9日");
    expect(formatDate(d, "H:mm:ss")).toBe("14:05:09");
    expect(formatDate(d, "hh:mm AM/PM")).toBe("02:05 PM");
    expect(formatDate(d, "yyMMdd")).toBe("260309");
  });

  it("resolves longer tokens before their prefixes (MM before M, dddd before d)", () => {
    expect(formatDate(d, "M/d/yyyy")).toBe("3/9/2026");
    expect(formatDate(d, "MM/dd")).toBe("03/09");
    expect(formatDate(d, "dddd")).toBe("星期一");
  });

  it("passes unknown characters through verbatim", () => {
    expect(formatDate(d, "yyyy!")).toBe("2026!");
  });
});

describe("evaluateField", () => {
  const ctx = {
    now: new Date(2026, 8, 9, 10, 30, 0),
    core: {
      creator: "作者",
      title: "标题",
      created: "2024-01-02T03:04:05Z",
      modified: "2025-05-06T07:08:09Z",
    },
    words: 42,
    chars: 300,
  };

  it("formats date fields with their picture", () => {
    expect(evaluateField(`DATE \\@ "yyyy/M/d"`, ctx)).toBe("2026/9/9");
    expect(evaluateField("DATE", ctx)).toBe("2026/9/9");
    expect(evaluateField(`TIME \\@ "H:mm"`, ctx)).toBe("10:30");
  });

  it("formats storedate/createdate from the core properties", () => {
    expect(evaluateField(`SAVEDATE \\@ "yyyy"`, ctx)).toBe("2025");
    // The default picture carries the clock time too — timezone-shifted, so
    // only the date part is pinned here.
    expect(evaluateField("CREATEDATE", ctx)).toMatch(/^2024\/1\/2 \d?\d:\d\d$/);
  });

  it("reads the document-information fields from the core properties", () => {
    expect(evaluateField("AUTHOR", ctx)).toBe("作者");
    expect(evaluateField("TITLE", ctx)).toBe("标题");
    expect(evaluateField("SUBJECT", ctx)).toBeNull();
  });

  it("counts words and characters", () => {
    expect(evaluateField("NUMWORDS", ctx)).toBe("42");
    expect(evaluateField("NUMCHARS", ctx)).toBe("300");
    expect(evaluateField("NUMWORDS", { ...ctx, words: undefined })).toBeNull();
  });

  it("leaves dynamic and unknown fields to the caller (null)", () => {
    expect(evaluateField("PAGE", ctx)).toBeNull();
    expect(evaluateField("NUMPAGES", ctx)).toBeNull();
    expect(evaluateField("SEQ 图 \\* ARABIC", ctx)).toBeNull();
  });
});

describe("fieldRef", () => {
  it("reads the simple field's instruction and cached value", () => {
    expect(fieldRef({ simpleField: { instruction: "DATE", cachedValue: "2026" } })).toEqual({
      kind: "simpleField",
      instruction: "DATE",
      result: "2026",
    });
  });

  it("reads the complex field's result", () => {
    expect(fieldRef({ complexField: { instruction: "PAGE", result: "3" } })).toEqual({
      kind: "complexField",
      instruction: "PAGE",
      result: "3",
    });
  });

  it("reads the checkbox's checked flag", () => {
    expect(fieldRef({ formField: { checkBox: { checked: true } } })).toEqual({
      kind: "formField",
      checked: true,
    });
    expect(fieldRef({ formField: {} })).toEqual({ kind: "formField", checked: false });
  });

  it("returns null for non-field branches and malformed entries", () => {
    expect(fieldRef({ footnoteReference: 1 })).toBeNull();
    expect(fieldRef({ simpleField: "not-an-object" })).toBeNull();
  });
});

describe("catalog helpers", () => {
  it("defaultInstruction maps a field name to its picture-carrying default", () => {
    expect(defaultInstruction("DATE")).toBe('DATE \\@ "yyyy/M/d"');
    expect(defaultInstruction("PAGE")).toBe("PAGE");
    expect(defaultInstruction("UNKNOWN")).toBe("UNKNOWN");
  });

  it("instructionName extracts the leading name of an instruction", () => {
    expect(instructionName('DATE \\@ "yyyy"')).toBe("DATE");
    expect(instructionName("seq 图")).toBe("SEQ");
  });
});
