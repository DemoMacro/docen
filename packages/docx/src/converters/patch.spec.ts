import { describe, expect, it } from "vitest";

import { generateDOCXSync, parseDOCXSync, patchDOCX, type JSONContent } from "../index";

const template = (text: string): Uint8Array =>
  generateDOCXSync({
    type: "doc",
    content: [{ type: "paragraph", content: [{ type: "text", text }] }],
  }) as Uint8Array;

const replacement = (...paragraphs: readonly string[]): JSONContent => ({
  type: "doc",
  content: paragraphs.map((text) => ({
    type: "paragraph",
    content: [{ type: "text", text }],
  })),
});

describe("patchDOCX", () => {
  it("replaces an inline placeholder without splitting the sentence", async () => {
    const patched = await patchDOCX({
      template: template("Hello {{name}}"),
      patches: { name: { content: replacement("Ada") } },
      outputType: "uint8array",
    });
    const document = parseDOCXSync(patched);

    expect(JSON.stringify(document)).toContain("Hello Ada");
    expect(JSON.stringify(document)).not.toContain("{{name}}");
  });

  it("replaces a whole placeholder paragraph with multiple block children", async () => {
    const patched = await patchDOCX({
      template: template("{{body}}"),
      patches: { body: { content: replacement("First", "Second") } },
      outputType: "uint8array",
    });
    const document = parseDOCXSync(patched);

    expect(JSON.stringify(document)).toContain("First");
    expect(JSON.stringify(document)).toContain("Second");
    expect(JSON.stringify(document)).not.toContain("{{body}}");
  });
});
