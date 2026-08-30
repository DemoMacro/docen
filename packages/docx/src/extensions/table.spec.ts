import { describe, it, expect } from "vitest";

import { parseDocx, renderDocx } from "./table";

// float round-trips byte-faithful through the table's renderDocx/parseDocx
// attrs passthrough (SKIP_KEYS drops only rows/columnWidthsRevision) — this is
// what lets a floating table keep its anchor across DOCX→JSON→DOCX even when
// the v1 renderer degrades (absolute anchor, center) and emits no CSS.
describe("table float round-trip", () => {
  it("preserves a full float anchor verbatim", () => {
    const float = {
      horizontalAnchor: "text",
      verticalAnchor: "text",
      absoluteHorizontalPosition: 720,
      absoluteVerticalPosition: 360,
      relativeHorizontalPosition: "right",
      relativeVerticalPosition: "top",
      leftFromText: 180,
      rightFromText: 180,
      topFromText: 200,
      bottomFromText: 200,
      overlap: "neverOverlap",
    };
    const back = parseDocx({ rows: [], ...renderDocx({ type: "table", attrs: { float } }) });
    expect(back.float).toEqual(float);
  });
});
