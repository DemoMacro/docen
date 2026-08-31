import type { MetafileMember, SourceCrop } from "../member";
import { carrierDrafts } from "./carrier";
import { finalize } from "./finalize";
import { embeddedEmfStream } from "./stream";

/**
 * Replay a dual-mode WMF's embedded EMF+ stream into drawing members sized to
 * the display box. Returns undefined when there is nothing to replay: no
 * usable stream, no drawable output, or degenerate geometry.
 *
 * The Office exporter emits an object definition immediately before each call
 * that uses it, so draws take paint state from the most recently defined
 * brush/pen while the record's ObjectId selects the path being drawn (slot ids
 * pair up exactly that way across the corpus).
 */
export function emfPlusMembers(
  bytes: Uint8Array,
  boxW: number,
  boxH: number,
  crop?: SourceCrop,
): MetafileMember[] | undefined {
  const emf = embeddedEmfStream(bytes);
  if (!emf) return undefined;
  const drafts = carrierDrafts(emf, undefined, 0);
  if (!drafts.length) return undefined;
  // The carrier's physical frame (EMR_HEADER rclFrame, [MS-EMF] §2.3.4.2)
  // anchors the display mapping, converted to EMF+ device px at 96 dpi — the
  // reference resolution these clipboard metafiles declare their extent at.
  // rclBounds is only an ink hint in a device basis that synthetic carriers
  // leave inconsistent with the frame; stretching bounds onto the extent
  // distorts the art Word renders at the frame's near-1:1 scale.
  let frame: { x: number; y: number; w: number; h: number } | undefined;
  if (emf.length >= 40) {
    const v = new DataView(emf.buffer, emf.byteOffset, emf.byteLength);
    if (v.getUint32(0, true) === 1 && v.getUint32(4, true) >= 88) {
      const l = (v.getInt32(24, true) * 96) / 2540,
        t = (v.getInt32(28, true) * 96) / 2540,
        r = (v.getInt32(32, true) * 96) / 2540,
        b = (v.getInt32(36, true) * 96) / 2540;
      if (r > l && b > t) frame = { x: l, y: t, w: r - l, h: b - t };
    }
  }
  return finalize(drafts, boxW, boxH, frame, crop);
}
