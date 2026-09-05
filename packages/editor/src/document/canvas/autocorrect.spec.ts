import { describe, expect, it } from "vitest";

import { autocorrectOf, correctionOf, smartQuoteOf } from "./autocorrect";

describe("smartQuoteOf", () => {
  it("opens after whitespace / start / opening punctuation", () => {
    expect(smartQuoteOf('"', "")).toBe("“");
    expect(smartQuoteOf('"', " ")).toBe("“");
    expect(smartQuoteOf("'", "\t")).toBe("‘");
    expect(smartQuoteOf('"', "(")).toBe("“");
    expect(smartQuoteOf('"', "—")).toBe("“");
  });

  it("closes after word characters and punctuation", () => {
    expect(smartQuoteOf('"', "o")).toBe("”");
    expect(smartQuoteOf("'", "n")).toBe("’");
    expect(smartQuoteOf('"', ",")).toBe("”");
    expect(smartQuoteOf('"', "”")).toBe("”");
  });

  it("leaves non-quotes alone", () => {
    expect(smartQuoteOf("a", "")).toBeNull();
    expect(smartQuoteOf("-", " ")).toBeNull();
  });
});

describe("correctionOf", () => {
  it("fixes known words and preserves case shape", () => {
    expect(correctionOf("teh")).toBe("the");
    expect(correctionOf("Teh")).toBe("The");
    expect(correctionOf("TEH")).toBe("THE");
    expect(correctionOf("recieve")).toBe("receive");
  });

  it("leaves unknown words alone", () => {
    expect(correctionOf("hello")).toBeNull();
  });
});

describe("autocorrectOf", () => {
  it("rewrites the word behind a boundary character, boundary trailing", () => {
    // "teh" + typed " " → "the " — the space stays as the word separator
    expect(autocorrectOf(" ", "teh")).toEqual({ text: "the ", back: 3 });
    expect(autocorrectOf(",", "Teh")).toEqual({ text: "The,", back: 3 });
  });

  it("replaces the quote in place with its curly side", () => {
    expect(autocorrectOf('"', "the ")).toEqual({ text: "“", back: 0 });
    expect(autocorrectOf("'", "don")).toEqual({ text: "’", back: 0 });
  });

  it("builds the em dash from the hyphen behind and ellipsis from two dots", () => {
    expect(autocorrectOf("-", "a-")).toEqual({ text: "—", back: 1 });
    expect(autocorrectOf(".", "a..")).toEqual({ text: "…", back: 2 });
  });

  it("returns null for ordinary characters", () => {
    expect(autocorrectOf("x", "teh ")).toBeNull();
    expect(autocorrectOf(" ", "hello")).toBeNull();
    expect(autocorrectOf("-", "a ")).toBeNull();
  });
});
