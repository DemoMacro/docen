/** Neutral drawing members a metafile replay emits — box-local px geometry,
 *  no Leafer and no docen types, so any renderer can consume them. */

/** One GDI ExtTextOut / EMF+ DrawString text emission, box-local. */
export interface MetafileTextRun {
  text: string;
  family?: string;
  sizePx: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  letterSpacingPx?: number;
}

/** Source-rect crop as 0..1 fractions per edge (a:srcRect /1000). */
export interface MetafileCrop {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** a:srcRect crop shared by the carrier and WMF-body replays. */
export type SourceCrop = MetafileCrop;

export type MetafileMember =
  | {
      kind: "picture";
      x: number;
      y: number;
      width: number;
      height: number;
      src?: string;
      blend?: "screen" | "multiply";
      crop?: MetafileCrop;
    }
  | {
      kind: "shape";
      x: number;
      y: number;
      width: number;
      height: number;
      preset?: string;
      fill?: string;
      line?: { px: number; color?: string };
    }
  | {
      kind: "path";
      x: number;
      y: number;
      width: number;
      height: number;
      d: string;
      fillRule?: "evenodd" | "nonzero";
      fill?: string;
      line?: { px: number; color?: string; dash?: string };
    }
  | {
      kind: "textBox";
      x: number;
      y: number;
      width: number;
      height: number;
      nowrap?: boolean;
      insets?: { left?: number; top?: number; right?: number; bottom?: number };
      rotation?: number;
      runs: MetafileTextRun[];
    };
