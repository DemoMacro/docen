/** A drawing's painted box plus its identity — how a click hit it and how the
 *  box re-resolves after a re-render (host laid paragraph + drawing index).
 *  `childPath` marks a group member's box. */
export interface DrawingHit {
  page: number;
  para: unknown;
  index: number;
  kind: "drawing" | "inline";
  x: number;
  y: number;
  width: number;
  height: number;
  /** The hit group member's index path (absent = the drawing's own box). */
  childPath?: readonly number[];
  /** Clockwise degrees — the selection frame tilts with the drawing. */
  rotation?: number;
}
