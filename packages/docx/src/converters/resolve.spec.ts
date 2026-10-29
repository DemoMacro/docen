import type { DocumentOptions } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { resolveDocument } from "./docx";

describe("resolveDocument run children", () => {
  it("reads 0.15 text payloads nested under runs", () => {
    const document: DocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    bold: true,
                    size: 28,
                    children: [{ text: "Nested", preserveSpace: true }],
                  },
                ],
              },
            },
          ],
        },
      ],
    };
    const json = resolveDocument(document);
    const text = json.content?.[0]?.content?.[0];

    expect(text).toMatchObject({ type: "text", text: "Nested" });
    expect(text?.marks?.map(({ type }) => type)).toContain("bold");
  });
});
