import type { ProjectedPresentation, ProjectedSlideMember } from "@docen/pptx";
import { describe, expect, it } from "vitest";

import { changedSlides, memberBatches } from "./slide-diff";

const pres = (text: string): ProjectedPresentation =>
  ({
    widthPx: 100,
    heightPx: 50,
    slides: [
      {
        members: [{ kind: "textBox", text: text } as never],
      },
    ],
  }) as unknown as ProjectedPresentation;

describe("changedSlides", () => {
  it("reports only structurally changed slides", () => {
    const previous = pres("before");
    const next = pres("after");
    expect(changedSlides(previous, next)).toEqual(new Set([0]));
  });

  it("reports none when every projection is unchanged", () => {
    const previous = pres("same");
    expect(changedSlides(previous, pres("same"))).toEqual(new Set());
  });

  it("forces a full paint when slide counts differ", () => {
    const previous = pres("same");
    (previous.slides as unknown[]).push({});
    expect(changedSlides(previous, pres("same"))).toBeUndefined();
  });
});

describe("memberBatches", () => {
  it("groups consecutive members from one slide child", () => {
    const members = [
      { kind: "path", sourceChildIndex: 1 },
      { kind: "path", sourceChildIndex: 1 },
      { kind: "shape", sourceChildIndex: 2 },
    ] as ProjectedSlideMember[];
    expect(memberBatches(members)).toEqual([[members[0], members[1]], [members[2]]]);
  });
});
