import type { ProjectedPresentation, ProjectedSlideMember } from "@docen/pptx";

/** A stable signature for projected paint data. Projection produces plain
 *  serializable values (image bytes have already become URLs), so JSON is the
 *  structural comparison — key insertion order is projection-owned and stable. */
export function paintSignature(value: unknown): string {
  return JSON.stringify(value, (_key, item) =>
    item instanceof Uint8Array ? `Uint8Array(${item.length})` : item,
  );
}

/** The slides whose projected paint differs. `undefined` means the first
 *  projection (or a shape change the old projection cannot pair with), so
 *  every slide is treated as changed. */
export function changedSlides(
  previous: ProjectedPresentation | null,
  next: ProjectedPresentation,
): Set<number> | undefined {
  if (!previous || previous.slides.length !== next.slides.length) return undefined;
  const changed = new Set<number>();
  for (const [index, slide] of next.slides.entries()) {
    const old = previous.slides[index];
    if (!old || paintSignature(old) !== paintSignature(slide)) changed.add(index);
  }
  return changed;
}

/** Consecutive members from one slide child form one reusable paint batch:
 *  a line's arrow members and a metafile's masked picture runs must repaint
 *  together to preserve their composite semantics. */
export function memberBatches(members: readonly ProjectedSlideMember[]): ProjectedSlideMember[][] {
  const batches: ProjectedSlideMember[][] = [];
  for (const member of members) {
    const batch = batches[batches.length - 1];
    if (batch && batch[0]?.sourceChildIndex === member.sourceChildIndex) batch.push(member);
    else batches.push([member]);
  }
  return batches;
}
