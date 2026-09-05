// @vitest-environment happy-dom
import { describe, expect, it } from "vitest";

import { wordWildcardToRegExp } from "./navigation";

describe("wordWildcardToRegExp", () => {
  it("translates the ? and * tokens", () => {
    expect(wordWildcardToRegExp("s?te")).toBe("s.te");
    expect(wordWildcardToRegExp("s*d")).toBe("s.*d");
  });

  it("maps word boundaries and negated sets", () => {
    expect(wordWildcardToRegExp("<pre>")).toBe("\\bpre\\b");
    expect(wordWildcardToRegExp("[!abc]")).toBe("[^abc]");
    expect(wordWildcardToRegExp("[a-m]x")).toBe("[a-m]x");
  });

  it("keeps counted repeats and quantifies the previous item for @", () => {
    expect(wordWildcardToRegExp("lo@t")).toBe("lo+t");
    expect(wordWildcardToRegExp("a{2,3}")).toBe("a{2,3}");
    expect(wordWildcardToRegExp("a{2,}b")).toBe("a{2,}b");
  });

  it("escapes metacharacters and malformed constructs to literals", () => {
    expect(wordWildcardToRegExp("a.b")).toBe("a\\.b");
    expect(wordWildcardToRegExp("(x)")).toBe("\\(x\\)");
    expect(wordWildcardToRegExp("a[bc")).toBe("a\\[bc");
    expect(wordWildcardToRegExp("a{b")).toBe("a\\{b");
    expect(wordWildcardToRegExp("@x")).toBe("@x");
  });

  it("produces a working RegExp for a Word-style pattern", () => {
    const re = new RegExp(wordWildcardToRegExp("<b?ll>"));
    expect(re.test("ball")).toBe(true);
    expect(re.test("bowl")).toBe(false);
  });
});
