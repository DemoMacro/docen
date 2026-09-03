export interface Xform {
  m11: number;
  m12: number;
  m21: number;
  m22: number;
  dx: number;
  dy: number;
}

export const IDENTITY: Xform = { m11: 1, m12: 0, m21: 0, m22: 1, dx: 0, dy: 0 };

export function xformPoint(t: Xform, x: number, y: number): [number, number] {
  return [x * t.m11 + y * t.m21 + t.dx, y * t.m22 + x * t.m12 + t.dy];
}

export function combine(a: Xform, b: Xform): Xform {
  return {
    m11: a.m11 * b.m11 + a.m21 * b.m12,
    m12: a.m12 * b.m11 + a.m22 * b.m12,
    m21: a.m11 * b.m21 + a.m21 * b.m22,
    m22: a.m12 * b.m21 + a.m22 * b.m22,
    dx: a.dx * b.m11 + a.dy * b.m12 + b.dx,
    dy: a.dx * b.m21 + a.dy * b.m22 + b.dy,
  };
}

/** The uniform length scale the transform applies — the larger column norm
 *  (rotation-safe: pen widths and glyph ems scale with their columns, not
 *  their rows). */
export function scaleOf(t: Xform): number {
  return Math.max(Math.hypot(t.m11, t.m21), Math.hypot(t.m12, t.m22));
}

/** XFORM payload: some exporters prefix a byte-length word (0x18); accept
 *  both shapes so a bare matrix still parses. */
export function readXform(view: DataView, at: number): Xform {
  const base = view.getUint32(at, true) === 24 ? at + 4 : at;
  return {
    m11: view.getFloat32(base, true),
    m12: view.getFloat32(base + 4, true),
    m21: view.getFloat32(base + 8, true),
    m22: view.getFloat32(base + 12, true),
    dx: view.getFloat32(base + 16, true),
    dy: view.getFloat32(base + 20, true),
  };
}

/** Raw EMR-world XFORM: six floats directly behind the record header (the
 *  byte-length prefix seen on EMF+ payloads does not appear here). */
export function readCarrierXform(view: DataView, at: number): Xform {
  return {
    m11: view.getFloat32(at, true),
    m12: view.getFloat32(at + 4, true),
    m21: view.getFloat32(at + 8, true),
    m22: view.getFloat32(at + 12, true),
    dx: view.getFloat32(at + 16, true),
    dy: view.getFloat32(at + 20, true),
  };
}
