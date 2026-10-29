import { describe, expect, it } from "vitest";

import { serializeJsonBinary, serializeRawParts } from "./json";

describe("JSON binary serialization", () => {
  it("rejects Blob values instead of silently serializing an empty object", () => {
    expect(() => serializeJsonBinary({ data: new Blob(["bytes"]) })).toThrow(TypeError);
  });

  it("rejects stream-backed raw parts instead of silently serializing an empty object", () => {
    const rawParts = [{ path: "media.bin", data: new ReadableStream() }];

    expect(() => serializeRawParts(rawParts)).toThrow(TypeError);
  });
});
