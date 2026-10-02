import { describe, expect, it } from "vitest";

import {
  evaluateField,
  FIELD_EVALUATORS,
  parseFieldInstruction,
  resolveField,
  type FieldContext,
} from "./fields";

// Evaluator registry specs — pure functions with independent oracle values.
// The context mirrors what the render pass and the update commands build; a
// missing context slice must yield null (the caller keeps the cache).

const now = new Date(2026, 8, 9, 10, 30, 0); // 2026-09-09 10:30, a Wednesday

const base: FieldContext = {
  now,
  core: {
    creator: "作者甲",
    title: "年度报告",
    subject: "财务",
    keywords: "报告, 年度",
    description: "备注文本",
    created: "2024-01-02T03:04:05Z",
    modified: "2025-05-06T07:08:09Z",
    lastPrinted: "2025-06-07T08:09:10Z",
    revision: 7,
  },
  filename: "报告.docx",
  words: 42,
  chars: 300,
  customProperties: new Map([["项目代码", "PRJ-1"]]),
  bookmarks: new Map([
    ["_Ref00000001", { text: "图 1", page: 2 }],
    ["目标", { text: "第三章", page: 6 }],
  ]),
  sequences: new Map([["图", 3]]),
  frame: {
    page: 5,
    pageCount: 12,
    section: 2,
    sectionPages: 4,
    pageFormat: "lowerRoman",
  },
};

describe("parseFieldInstruction", () => {
  it("splits the name, positional args, and switch values", () => {
    expect(parseFieldInstruction('DATE \\@ "yyyy/M/d"')).toEqual({
      name: "DATE",
      raw: 'DATE \\@ "yyyy/M/d"',
      args: [],
      switches: { "@": "yyyy/M/d" },
    });
    expect(parseFieldInstruction("SEQ 图 \\* ARABIC")).toMatchObject({
      name: "SEQ",
      args: ["图"],
      switches: { "*": "ARABIC" },
    });
    expect(parseFieldInstruction("REF _Ref00000001 \\h")).toMatchObject({
      name: "REF",
      args: ["_Ref00000001"],
      switches: { h: "" },
    });
  });

  it("keeps a quoted argument whole and uppercases the name", () => {
    expect(parseFieldInstruction('ref "my bookmark" \\h')).toMatchObject({
      name: "REF",
      args: ["my bookmark"],
      switches: { h: "" },
    });
  });
});

describe("field evaluator registry", () => {
  it("covers every field the catalog advertises", () => {
    const names = Object.keys(FIELD_EVALUATORS);
    for (const name of ["PAGE", "NUMPAGES", "DATE", "TIME", "AUTHOR", "TITLE"]) {
      expect(names).toContain(name);
    }
  });

  it("renders the numbering fields from the page frame", () => {
    expect(resolveField("PAGE", base)).toBe("v");
    expect(resolveField("NUMPAGES", base)).toBe("12");
    expect(resolveField("SECTION", base)).toBe("ii");
    expect(resolveField("SECTIONPAGES", base)).toBe("iv");
  });

  it("leaves page-dependent fields unresolved with no page frame", () => {
    expect(resolveField("PAGE", { ...base, frame: undefined })).toBeNull();
    expect(resolveField("NUMPAGES", { ...base, frame: undefined })).toBeNull();
    expect(resolveField("SECTION", { ...base, frame: { pageCount: 3 } })).toBeNull();
  });

  it("formats the date/time fields from the injected clock and pictures", () => {
    expect(resolveField(`DATE \\@ "yyyy/M/d"`, base)).toBe("2026/9/9");
    expect(resolveField("TIME", base)).toBe("10:30:00");
    expect(resolveField(`SAVEDATE \\@ "yyyy"`, base)).toBe("2025");
    expect(resolveField("PRINTDATE", base)).toBe("2025/6/7");
    expect(resolveField("CREATEDATE", { ...base, core: {} })).toBeNull();
  });

  it("reads the document-information fields from the core properties", () => {
    expect(resolveField("AUTHOR", base)).toBe("作者甲");
    expect(resolveField("TITLE", base)).toBe("年度报告");
    expect(resolveField("SUBJECT", base)).toBe("财务");
    expect(resolveField("KEYWORDS", base)).toBe("报告, 年度");
    expect(resolveField("COMMENTS", base)).toBe("备注文本");
    expect(resolveField("SUBJECT", { ...base, core: {} })).toBeNull();
  });

  it("reads FILENAME and REVNUM", () => {
    expect(resolveField("FILENAME", base)).toBe("报告.docx");
    expect(resolveField("REVNUM", base)).toBe("7");
    expect(resolveField("FILENAME", { ...base, filename: undefined })).toBeNull();
  });

  it("resolves REF/PAGEREF from the bookmark table", () => {
    expect(resolveField("REF _Ref00000001 \\h", base)).toBe("图 1");
    expect(resolveField("PAGEREF 目标 \\h", base)).toBe("vi");
    expect(resolveField("REF 不存在", base)).toBeNull();
  });

  it("formats PAGEREF with the target bookmark's own section numbering", () => {
    // The bookmark sits on a section whose pages restart at 5 in lowercase
    // Roman; the field itself is on a decimal page 1. The target's display
    // number and format win.
    const ctx = {
      ...base,
      frame: { page: 1, pageCount: 9 },
      bookmarks: new Map([["目标", { text: "第三章", page: 5, pageFormat: "lowerRoman" }]]),
    };
    expect(resolveField("PAGEREF 目标 \\h", ctx)).toBe("v");
    // No bookmark format → the field's own section format is the fallback.
    const fallback = {
      ...ctx,
      frame: { page: 1, pageCount: 9, pageFormat: "upperRoman" },
      bookmarks: new Map([["目标", { text: "第三章", page: 5 }]]),
    };
    expect(resolveField("PAGEREF 目标 \\h", fallback)).toBe("V");
  });

  it("guards malformed revision numbers", () => {
    expect(resolveField("REVNUM", { ...base, revision: Number.NaN, core: {} })).toBeNull();
    expect(resolveField("REVNUM", { ...base, revision: undefined, core: {} })).toBeNull();
    expect(
      resolveField("REVNUM", { ...base, revision: undefined, core: { revision: "abc" } }),
    ).toBeNull();
    expect(resolveField("REVNUM", { ...base, revision: undefined, core: { revision: "12" } })).toBe(
      "12",
    );
  });

  it("resolves SEQ per occurrence ordinal", () => {
    expect(resolveField("SEQ 图 \\* ARABIC", base)).toBe("3");
    expect(resolveField("SEQ 表", base)).toBeNull();
  });

  it("formats SEQ per the \\* switch and prefixes the \\s chapter number", () => {
    const ctx: FieldContext = {
      ...base,
      sequences: new Map([["Figure", 4]]),
      chapters: new Map([[1, "2"]]),
      captionSeparators: new Map([["Figure", ":"]]),
    };
    // Word's caption format list: the switch word is case-insensitive and its
    // casing picks the output case (Roman/ROMAN → I, roman → i).
    expect(resolveField("SEQ Figure \\* ARABIC", ctx)).toBe("4");
    expect(resolveField("SEQ Figure \\* Arabic", ctx)).toBe("4");
    expect(resolveField("SEQ Figure \\* roman", ctx)).toBe("iv");
    expect(resolveField("SEQ Figure \\* Roman", ctx)).toBe("IV");
    expect(resolveField("SEQ Figure \\* ROMAN", ctx)).toBe("IV");
    expect(resolveField("SEQ Figure \\* alphabetic", ctx)).toBe("d");
    expect(resolveField("SEQ Figure \\* Alphabetic", ctx)).toBe("D");
    expect(resolveField("SEQ Figure \\* ALPHABETIC", ctx)).toBe("D");
    // No switch (or an unknown one) is Word's ARABIC default.
    expect(resolveField("SEQ Figure", ctx)).toBe("4");
    expect(resolveField("SEQ Figure \\* BOGUS", ctx)).toBe("4");
    // "Include chapter number": chapter + separator + formatted sequence.
    expect(resolveField("SEQ Figure \\* ARABIC \\s 1", ctx)).toBe("2:4");
    // A level with no chapter in context keeps the plain number; a malformed
    // level is ignored entirely (Word's invalid-switch behavior).
    expect(resolveField("SEQ Figure \\* ARABIC \\s 2", ctx)).toBe("4");
    expect(resolveField("SEQ Figure \\* ARABIC \\s 1.5", ctx)).toBe("4");
    // No caption separator setting → Word's hyphen default.
    expect(
      resolveField("SEQ Figure \\* ARABIC \\s 1", { ...ctx, captionSeparators: undefined }),
    ).toBe("2-4");
  });

  it("maps INFO arguments onto the information evaluators", () => {
    expect(resolveField("INFO Title", base)).toBe("年度报告");
    expect(resolveField("INFO NumPages", base)).toBe("12");
    expect(resolveField("INFO Filename", base)).toBe("报告.docx");
    expect(resolveField("INFO Unknown", base)).toBeNull();
  });

  it("reads DOCPROPERTY by name, case-insensitively", () => {
    expect(resolveField("DOCPROPERTY 项目代码", base)).toBe("PRJ-1");
    expect(resolveField("DOCPROPERTY wrong", base)).toBeNull();
  });

  it("returns null for unknown instructions", () => {
    expect(resolveField("MERGEFIELD name", base)).toBeNull();
    expect(resolveField("", base)).toBeNull();
  });

  it("keeps evaluateField as the registry entry point", () => {
    expect(evaluateField("TITLE", base)).toBe(resolveField("TITLE", base));
    expect(evaluateField("PAGE", { ...base, frame: undefined })).toBeNull();
  });
});
