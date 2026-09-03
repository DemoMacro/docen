import type { MetafileMember, SourceCrop } from "../member";
import type { Draft } from "./draft";

const MEMBERS_CAP = 4000;

/** Whether a draft joins the union bounding box that drives the final
 *  normalization. After carrier world-transform projection all drafts share
 *  one device space, so text participates like any other content.
 *  Zero-area strokes join nothing: they render as hairlines at best yet can
 *  drag the box to a far corner. */
function participatesInBox(dr: Draft): boolean {
  if (dr.kind === "path") {
    const b = boxOf(dr);
    return b.x1 - b.x0 > 1 && b.y1 - b.y0 > 1;
  }
  return true;
}

/** Scale drafts from EMF device coordinates into the display box and shape
 *  the renderer-facing members. Word maps the EMF's declared physical frame
 *  (EMR_HEADER rclFrame at 96 dpi) onto the extent independently per axis —
 *  content past the frame edges stays inside the box instead of stretching
 *  everything above it. When the frame is unavailable or degenerate, the
 *  union bounding box stands in.
 *
 *  An a:srcRect crops the frame first (pixel-verified on the corpus stack
 *  where the visible frame is the top third of the carrier: Word stretches
 *  only that region onto the extent, near-aspect-preserved, and the records
 *  below the crop line never draw — members may now reach past the box, so
 *  the painter clips the replay to the extent like GDI does). */
export function finalize(
  drafts: Draft[],
  boxW: number,
  boxH: number,
  frame?: { x: number; y: number; w: number; h: number },
  crop?: SourceCrop,
): MetafileMember[] | undefined {
  let minX: number, minY: number, sX: number, sY: number;
  if (frame && frame.w > 0 && frame.h > 0) {
    minX = frame.x;
    minY = frame.y;
    let fw = frame.w,
      fh = frame.h;
    if (crop) {
      minX += crop.left * fw;
      minY += crop.top * fh;
      fw *= 1 - crop.left - crop.right;
      fh *= 1 - crop.top - crop.bottom;
    }
    if (!(fw > 0 && fh > 0)) return undefined;
    sX = boxW / fw;
    sY = boxH / fh;
  } else {
    let x0 = Infinity,
      y0 = Infinity,
      x1 = -Infinity,
      y1 = -Infinity;
    for (const dr of drafts) {
      if (!participatesInBox(dr)) continue;
      if (dr.kind === "path") {
        for (const [, nums] of dr.cmds) {
          for (let i = 0; i + 1 < nums.length; i += 2) {
            x0 = Math.min(x0, nums[i]);
            x1 = Math.max(x1, nums[i]);
            y0 = Math.min(y0, nums[i + 1]);
            y1 = Math.max(y1, nums[i + 1]);
          }
        }
        continue;
      }
      x0 = Math.min(x0, dr.x);
      y0 = Math.min(y0, dr.y);
      x1 = Math.max(x1, dr.x + dr.w);
      y1 = Math.max(y1, dr.y + dr.h);
    }
    if (crop) {
      const w = x1 - x0,
        h = y1 - y0;
      x0 += crop.left * w;
      y0 += crop.top * h;
      x1 -= crop.right * w;
      y1 -= crop.bottom * h;
    }
    minX = x0;
    minY = y0;
    if (!(x1 - x0 > 0 && y1 - y0 > 0)) return undefined;
    sX = boxW / (x1 - x0);
    sY = boxH / (y1 - y0);
  }
  const X = (x: number) => (x - minX) * sX;
  const Y = (y: number) => (y - minY) * sY;
  const members: MetafileMember[] = [];
  for (const dr of drafts) {
    if (dr.kind === "text") {
      members.push({
        kind: "textBox",
        x: X(dr.x),
        y: Y(dr.y),
        width: dr.w * sX,
        height: dr.h * sY,
        nowrap: true,
        ...(dr.cellInsetWorld ? { insets: { left: dr.cellInsetWorld * sX } } : {}),
        ...(dr.rotation ? { rotation: dr.rotation } : {}),
        runs: [
          {
            text: dr.text,
            family: dr.family,
            sizePx: dr.sizeWorld * sY,
            ...(dr.color ? { color: dr.color } : {}),
            ...(dr.bold ? { bold: true } : {}),
            ...(dr.letterSpacingWorld ? { letterSpacingPx: dr.letterSpacingWorld * sX } : {}),
          },
        ],
      });
      continue;
    }
    if (dr.kind === "pic") {
      members.push({
        kind: "picture",
        x: X(dr.x),
        y: Y(dr.y),
        width: dr.w * sX,
        height: dr.h * sY,
        src: dr.src,
        ...(dr.blend ? { blend: dr.blend } : {}),
        ...(dr.crop
          ? {
              crop: {
                left: Math.max(0, Math.min(1, dr.crop.l)),
                top: Math.max(0, Math.min(1, dr.crop.t)),
                right: Math.max(0, Math.min(1, dr.crop.r)),
                bottom: Math.max(0, Math.min(1, dr.crop.b)),
              },
            }
          : {}),
      });
      continue;
    }
    const parts: string[] = [];
    for (const [op, nums] of dr.cmds) {
      if (op === "Z") {
        parts.push("Z");
        continue;
      }
      const coords: string[] = [];
      for (let i = 0; i + 1 < nums.length; i += 2) {
        coords.push(round1(X(nums[i])), round1(Y(nums[i + 1])));
      }
      parts.push(op + coords.join(" "));
    }
    if (parts.length < 2) continue;
    const strokeWidth =
      dr.strokeWidth != null
        ? Math.max(1, Math.round(dr.strokeWidth * ((sX + sY) / 2)))
        : undefined;
    members.push({
      kind: "path",
      x: 0,
      y: 0,
      width: boxW,
      height: boxH,
      d: parts.join(""),
      ...(dr.fill ? { fill: dr.fill, fillRule: dr.fillRule ?? "evenodd" } : {}),
      ...(dr.strokeColor && strokeWidth
        ? {
            line: { px: strokeWidth, color: dr.strokeColor, ...(dr.dash ? { dash: dr.dash } : {}) },
          }
        : {}),
    });
    if (members.length > MEMBERS_CAP) return undefined;
  }
  return members.length ? members : undefined;
}

function round1(n: number): string {
  return String(Math.round(n * 10) / 10);
}

/** Boxes coincide within the tolerance that separates a dual pair's float32
 *  EMF+ geometry from its int16-quantized GDI twin. */
export function sameBox(
  a: { x0: number; y0: number; x1: number; y1: number },
  b: { x0: number; y0: number; x1: number; y1: number },
): boolean {
  return (
    Math.abs(a.x0 - b.x0) <= 2 &&
    Math.abs(a.y0 - b.y0) <= 2 &&
    Math.abs(a.x1 - b.x1) <= 2 &&
    Math.abs(a.y1 - b.y1) <= 2
  );
}

/** True when `blend` equals `solid` composited over white at one uniform
 *  alpha — the systematic relation between a dual pair's fills: argbHex
 *  pre-blends the EMF+ brush's partial alpha (25% FFE699 → FFF9E5) while the
 *  GDI fallback carries the flat COLORREF. Channel 255 carries no alpha
 *  information and is skipped; the remaining channels must agree on alpha
 *  within rounding. */
export function sameSourceBlend(blend: string | undefined, solid: string): boolean {
  if (!blend) return false;
  if (blend === solid) return true;
  let alpha = -1;
  for (let i = 0; i < 3; i++) {
    const b = parseInt(blend.slice(i * 2, i * 2 + 2), 16);
    const s = parseInt(solid.slice(i * 2, i * 2 + 2), 16);
    if (s === 255) {
      if (b !== 255) return false;
      continue;
    }
    const a = (b - 255) / (s - 255);
    if (a <= 0 || a >= 1) return false;
    if (alpha < 0) alpha = a;
    else if (Math.abs(a - alpha) > 0.03) return false;
  }
  return alpha > 0;
}

/** World-space bounding box of one draft (diagnostics hook shape). */
export function boxOf(dr: Draft): { x0: number; y0: number; x1: number; y1: number } {
  if (dr.kind !== "path") {
    return { x0: dr.x, y0: dr.y, x1: dr.x + dr.w, y1: dr.y + dr.h };
  }
  let x0 = Infinity,
    y0 = Infinity,
    x1 = -Infinity,
    y1 = -Infinity;
  for (const [, nums] of dr.cmds) {
    for (let i = 0; i + 1 < nums.length; i += 2) {
      x0 = Math.min(x0, nums[i]);
      x1 = Math.max(x1, nums[i]);
      y0 = Math.min(y0, nums[i + 1]);
      y1 = Math.max(y1, nums[i + 1]);
    }
  }
  return { x0, y0, x1, y1 };
}
