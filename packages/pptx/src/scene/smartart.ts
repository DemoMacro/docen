// Graphic-frame SmartArt (p:graphicFrame dgm:relIds) projection: the frame
// plus the parsed node tree and layout family. The full DGM layout engine is
// a later batch; the painter consumes this normalized fallback payload.

import type { LayoutDrawingMember } from "@docen/layout";
import type { SmartArtOptions } from "@office-open/pptx";

import { emuOf, type Xform } from "./geometry";

export function smartArtMember(
  smartArt: SmartArtOptions,
  t: Xform,
  childPath: readonly number[] | undefined,
): LayoutDrawingMember {
  return {
    kind: "smartArt",
    x: t.sx * emuOf(smartArt.x) + t.dx,
    y: t.sy * emuOf(smartArt.y) + t.dy,
    width: t.sx * emuOf(smartArt.width),
    height: t.sy * emuOf(smartArt.height),
    ...(typeof smartArt.layout === "string" ? { layout: smartArt.layout } : {}),
    ...(childPath ? { childPath } : {}),
    nodes: smartArt.nodes ?? [],
  };
}
