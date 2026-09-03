import { xformPoint, type Xform } from "./xform";

// A sub-path as raw command tuples; "M" starts, "L"/"C" continue, "Z" closes.
export type PathCmds = Array<["M" | "L" | "C" | "Z", number[]]>;

export interface PathDraft {
  kind: "path";
  cmds: PathCmds;
  fill?: string;
  /** Fill rule the source declared (GDI SetPolyFillMode); the finalize pass
   *  defaults paths without one to evenodd (GDI's ALTERNATE device default). */
  fillRule?: "evenodd" | "nonzero";
  strokeWidth?: number;
  strokeColor?: string;
  /** Preset dash token — threaded to the member's line.dash verbatim. */
  dash?: string;
}

export interface PicDraft {
  kind: "pic";
  x: number;
  y: number;
  w: number;
  h: number;
  src: string;
  /** Source-rectangle crop, fractions of each image edge (DrawImagePoints). */
  crop?: { l: number; t: number; r: number; b: number };
  /** Ternary raster-op emulation blend (SRCPAINT → screen, SRCAND → multiply). */
  blend?: "screen" | "multiply";
}

/** A GDI-side ExtTextOutW run, kept in the same world coordinate space as
 *  the EMF+ drafts so finalize maps everything through one normalization.
 *  Dual-mode files carry real text as GDI records — their EmfPlusDrawString
 *  counterparts are zero-data stubs. */
export interface TextDraft {
  kind: "text";
  x: number;
  y: number;
  w: number;
  h: number;
  text: string;
  family: string;
  sizeWorld: number;
  color?: string;
  bold?: boolean;
  /** Extra per-character advance from the trail Dx run, device px — GDI
   *  tracked text (letter-spaced labels) reports wider advances than the
   *  glyphs are wide; the member threads the difference into the run style. */
  letterSpacingWorld?: number;
  /** Cross-axis inset of the ink inside its cell, device px — a vertical
   *  (@font) run centers each glyph in its cell, so the ink starts
   *  (cellW − em)/2 into the box the fold produced. */
  cellInsetWorld?: number;
  /** Clockwise screen angle of the text-direction column, degrees — a
   *  rotated world transform means vertical text (plan-box columns); the
   *  member carries it so the paint rotates instead of laying the run
   *  horizontally across neighboring columns. */
  rotation?: number;
}

export type Draft = PathDraft | PicDraft | TextDraft;

export function pushPath(
  drafts: Draft[],
  rawCmds: PathCmds,
  xf: Xform,
  paint: {
    fill?: string;
    fillRule?: "evenodd" | "nonzero";
    strokeColor?: string;
    strokeWidth?: number;
    dash?: string;
  },
): void {
  const transformed: PathCmds = rawCmds.map(([op, nums]) => {
    if (op === "Z") return [op, nums];
    const out: number[] = [];
    for (let i = 0; i + 1 < nums.length; i += 2) {
      const [x, y] = xformPoint(xf, nums[i], nums[i + 1]);
      out.push(x, y);
    }
    return [op, out];
  });
  drafts.push({
    kind: "path",
    cmds: transformed,
    ...(paint.fill
      ? { fill: paint.fill, ...(paint.fillRule ? { fillRule: paint.fillRule } : {}) }
      : {}),
    ...(paint.strokeColor && paint.strokeWidth != null
      ? {
          strokeColor: paint.strokeColor,
          strokeWidth: paint.strokeWidth,
          ...(paint.dash ? { dash: paint.dash } : {}),
        }
      : {}),
  });
}

export function pushRect(
  drafts: Draft[],
  x: number,
  y: number,
  w: number,
  h: number,
  fill?: string,
): void {
  if (!fill || !(w >= 0 && h >= 0)) return;
  drafts.push({
    kind: "path",
    cmds: [
      ["M", [x, y]],
      ["L", [x + w, y]],
      ["L", [x + w, y + h]],
      ["L", [x, y + h]],
      ["Z", []],
    ],
    fill,
  });
}
