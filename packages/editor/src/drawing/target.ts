import type { ChartPartHit } from "@docen/core";

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
  /** The hit chart sub-element (Word's second-stage click inside the framed
   *  chart — the selection itself stays the chart's NodeSelection). */
  chartPart?: ChartPartHit;
  /** Clockwise degrees — the selection frame tilts with the drawing. */
  rotation?: number;
}
