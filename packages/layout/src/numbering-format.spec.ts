import { describe, expect, it } from "vitest";

import { formatNumber, romanNumeral } from "./numbering-format";

describe("formatNumber", () => {
  it("renders decimal for the default and unknown tokens", () => {
    expect(formatNumber(undefined, 12)).toBe("12");
    expect(formatNumber("decimal", 12)).toBe("12");
    expect(formatNumber("notAToken", 12)).toBe("12");
  });

  it("renders the latin letter and roman runs", () => {
    expect(formatNumber("lowerLetter", 1)).toBe("a");
    expect(formatNumber("lowerLetter", 26)).toBe("z");
    expect(formatNumber("lowerLetter", 27)).toBe("aa");
    expect(formatNumber("upperLetter", 28)).toBe("AB");
    expect(formatNumber("lowerRoman", 4)).toBe("iv");
    expect(formatNumber("upperRoman", 1994)).toBe("MCMXCIV");
    expect(formatNumber("hex", 255)).toBe("FF");
  });

  it("renders the CJK numerals, legal forms keeping the leading 壹", () => {
    expect(formatNumber("chineseCounting", 1)).toBe("一");
    expect(formatNumber("chineseCounting", 10)).toBe("十");
    expect(formatNumber("chineseCounting", 11)).toBe("十一");
    expect(formatNumber("chineseCounting", 101)).toBe("一百零一");
    expect(formatNumber("chineseLegalSimplified", 10)).toBe("壹拾");
    expect(formatNumber("chineseLegalSimplified", 12)).toBe("壹拾贰");
    expect(formatNumber("chineseCountingThousand", 1234)).toBe("一千二百三十四");
  });

  it("falls back to decimal past a system's glyph table", () => {
    expect(formatNumber("decimalEnclosedCircle", 1)).toBe("①");
    expect(formatNumber("decimalEnclosedCircle", 20)).toBe("⑳");
    expect(formatNumber("decimalEnclosedCircle", 21)).toBe("21");
    expect(formatNumber("ganada", 14)).toBe("ㅎ");
    expect(formatNumber("ganada", 15)).toBe("15");
    expect(formatNumber("koreanCounting", 15)).toBe("15");
  });

  it("renders the kana, jamo, dash and english-word tokens", () => {
    expect(formatNumber("aiueo", 1)).toBe("あ");
    expect(formatNumber("iroha", 2)).toBe("ろ");
    expect(formatNumber("katakana", 2)).toBe("イ");
    expect(formatNumber("katakanaHalf", 3)).toBe("ｳ");
    expect(formatNumber("koreanCounting", 3)).toBe("다");
    expect(formatNumber("ganada", 2)).toBe("ㄴ");
    expect(formatNumber("numberInDash", 7)).toBe("- 7 -");
    expect(formatNumber("ordinal", 1)).toBe("1st");
    expect(formatNumber("ordinal", 11)).toBe("11th");
    expect(formatNumber("ordinal", 112)).toBe("112th");
    expect(formatNumber("cardinalText", 21)).toBe("twenty-one");
    expect(formatNumber("ordinalText", 1)).toBe("first");
    expect(formatNumber("ordinalText", 21)).toBe("twenty-first");
    expect(formatNumber("ordinalText", 100)).toBe("one hundredth");
  });

  it("keeps romanNumeral exported for the footnote ordinals", () => {
    expect(romanNumeral(9, false)).toBe("ix");
  });
});
