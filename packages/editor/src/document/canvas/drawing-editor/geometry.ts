/**
 * Drawing-editor geometry — the format-independent math behind the selection
 * frame: which resize handle a pointer grabbed, and what the box becomes when
 * that handle drags. Pure functions on plain rects; the host adapter turns the
 * resulting box back into document attrs (docx today, pptx/xlsx tomorrow).
 */

export type HandleId = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";

/** The edited drawing's box — page-local px at scale 1 (the hit box's space). */
export interface Box {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** The eight handles, clockwise from the top-left corner. */
export const HANDLES: readonly HandleId[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];

/** Half the grab zone beyond a handle's visual size — Word's handles are small
 *  squares but the pointer catches a few px around them. */
const GRAB = 6;

/** The handle under a page-local point, or null when the point is on the body
 *  (a move) or outside the frame entirely. Corner/edge handles win over the
 *  body — a drag starting on the frame edge resizes, never moves. */
export function handleAt(box: Box, px: number, py: number): HandleId | null {
  const within = px >= box.x - GRAB && px <= box.x + box.width + GRAB;
  const vertical = py >= box.y - GRAB && py <= box.y + box.height + GRAB;
  if (!within || !vertical) return null;
  const nearW = Math.abs(px - box.x) <= GRAB;
  const nearE = Math.abs(px - (box.x + box.width)) <= GRAB;
  const nearN = Math.abs(py - box.y) <= GRAB;
  const nearS = Math.abs(py - (box.y + box.height)) <= GRAB;
  if (nearN && nearW) return "nw";
  if (nearN && nearE) return "ne";
  if (nearS && nearW) return "sw";
  if (nearS && nearE) return "se";
  if (nearN) return "n";
  if (nearS) return "s";
  if (nearW) return "w";
  if (nearE) return "e";
  return null;
}

/** Floor to a whole px, never below 1 — a drag through zero must not flip the
 *  box or hand the painter a 0-size node. */
const positive = (n: number): number => Math.max(1, Math.round(n));

/** The box after dragging `handle` by (dx, dy). The opposite edge/corner stays
 *  anchored (Word: dragging a handle keeps the rest of the frame put). Corner
 *  drags keep the aspect ratio (Word's default for pictures); edges resize one
 *  axis freely. */
export function resizeBox(box: Box, handle: HandleId, dx: number, dy: number, min = 24): Box {
  const west = handle.includes("w");
  const east = handle.includes("e");
  const north = handle.includes("n");
  const south = handle.includes("s");
  const corner = west !== east && north !== south;

  let { x, y, width, height } = box;
  if (west) {
    width = box.width - dx;
    x = box.x + dx;
  } else if (east) {
    width = box.width + dx;
  }
  if (north) {
    height = box.height - dy;
    y = box.y + dy;
  } else if (south) {
    height = box.height + dy;
  }

  if (corner) {
    // Aspect-locked: the larger axis drag drives, the other follows. The
    // anchor (opposite corner) stays where it started.
    const ratio = box.height / box.width;
    if (width / box.width >= height / box.height) {
      height = width * ratio;
    } else {
      width = height / ratio;
    }
    if (west) x = box.x + box.width - width;
    if (north) y = box.y + box.height - height;
  }

  // Below the minimum, clamp the moving edge to the minimum size instead of
  // letting the drag invert the box past its anchor.
  if (width < min) {
    if (west) x = box.x + box.width - min;
    width = min;
  }
  if (height < min) {
    if (north) y = box.y + box.height - min;
    height = min;
  }
  return { x: positive(x), y: positive(y), width: positive(width), height: positive(height) };
}

/** The clockwise angle (degrees) the pointer swept around the box center
 *  from point a to point b, normalized to (-180, 180] — one gesture step's
 *  delta (the caller accumulates per pointer-move, so the normalization
 *  keeps the spin continuous across the ±180° wrap). Screen px; the angle
 *  is scale-free. */
export function rotateDelta(
  cx: number,
  cy: number,
  ax: number,
  ay: number,
  bx: number,
  by: number,
): number {
  const a = (Math.atan2(ay - cy, ax - cx) * 180) / Math.PI;
  const b = (Math.atan2(by - cy, bx - cx) * 180) / Math.PI;
  const delta = b - a;
  if (delta > 180) return delta - 360;
  if (delta <= -180) return delta + 360;
  return delta;
}

/** The a:srcRect crop as fractions of the source (0.1 = 10% off that edge). */
export interface CropFractions {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** The full (uncropped) source rectangle behind a cropped box: the visible
 *  width covers 1 − left − right of the source, so the source spans
 *  width / (1 − left − right), shifted left by the left crop's share of the
 *  visible width. Degenerate crops (fully cropped axis) fall back to the
 *  visible box. Page-local px. */
export function cropFullBox(box: Box, crop: CropFractions): Box {
  const w = 1 - crop.left - crop.right;
  const h = 1 - crop.top - crop.bottom;
  return {
    x: box.x - box.width * (w > 0 ? crop.left / w : 0),
    y: box.y - box.height * (h > 0 ? crop.top / h : 0),
    width: w > 0 ? box.width / w : box.width,
    height: h > 0 ? box.height / h : box.height,
  };
}

/** The crop after dragging crop handle `handle` by fraction deltas (the
 *  caller divides a pointer delta by the full box's size; dl/dr are rightward
 *  and dt/db downward). Insets grow as their edge drags inward — a west
 *  handle moving right grows `left`, an east handle moving right (dr > 0)
 *  shrinks `right` — and each inset clamps so the opposite inset keeps a
 *  visible remainder of at least MIN. */
export function resizeCrop(
  crop: CropFractions,
  handle: HandleId,
  dl: number,
  dt: number,
  dr: number,
  db: number,
): CropFractions {
  const MAX = 0.9;
  const MIN = 0.1;
  const next = { ...crop };
  if (handle.includes("w")) next.left = clampCrop(crop.left + dl, crop.right, MIN, MAX);
  if (handle.includes("e")) next.right = clampCrop(crop.right - dr, crop.left, MIN, MAX);
  if (handle.includes("n")) next.top = clampCrop(crop.top + dt, crop.bottom, MIN, MAX);
  if (handle.includes("s")) next.bottom = clampCrop(crop.bottom - db, crop.top, MIN, MAX);
  return next;
}

/** One inset dragged by `value`, clamped so the opposite inset keeps at most
 *  MAX − MIN of the axis visible. */
function clampCrop(value: number, opposite: number, min: number, max: number): number {
  return Math.min(Math.max(value, 0), Math.max(0, max - opposite - min));
}
