// Bridge from the leafer-x-metafile plugin's neutral member types to the
// layout engine's drawing members: field-by-field passthrough for pictures,
// shapes, and paths; plugin text runs become one single-run paragraph block
// each (the layout engine lays blocks, the metafile replay only knows runs).

import type {
  LayoutBlock,
  LayoutDrawingMember,
  LayoutPictureCrop,
  LayoutTextStyle,
} from "@docen/layout";
import type { MetafileCrop, MetafileMember, MetafileTextRun } from "leafer-x-metafile";

function cropOf(crop: MetafileCrop): LayoutPictureCrop {
  return { left: crop.left, top: crop.top, right: crop.right, bottom: crop.bottom };
}

function styleOf(run: MetafileTextRun): LayoutTextStyle {
  return {
    family: run.family ?? "",
    sizePx: run.sizePx,
    ...(run.color ? { color: run.color } : {}),
    ...(run.bold ? { bold: true } : {}),
    ...(run.italic ? { italic: true } : {}),
    ...(run.letterSpacingPx ? { letterSpacingPx: run.letterSpacingPx } : {}),
  };
}

function runBlock(run: MetafileTextRun): LayoutBlock {
  return { kind: "paragraph", inline: [{ kind: "text", text: run.text, style: styleOf(run) }] };
}

/** Replay members → layout drawing members (identity for geometry; text runs
 *  wrap into paragraph blocks). */
export function toLayoutMembers(members: MetafileMember[]): LayoutDrawingMember[] {
  return members.map((member): LayoutDrawingMember => {
    switch (member.kind) {
      case "picture":
        return {
          kind: "picture",
          x: member.x,
          y: member.y,
          width: member.width,
          height: member.height,
          ...(member.src ? { src: member.src } : {}),
          ...(member.blend ? { blend: member.blend } : {}),
          ...(member.crop ? { crop: cropOf(member.crop) } : {}),
        };
      case "shape":
        return {
          kind: "shape",
          x: member.x,
          y: member.y,
          width: member.width,
          height: member.height,
          ...(member.preset ? { preset: member.preset } : {}),
          ...(member.fill ? { fill: member.fill } : {}),
          ...(member.line ? { line: member.line } : {}),
        };
      case "path":
        return {
          kind: "path",
          x: member.x,
          y: member.y,
          width: member.width,
          height: member.height,
          d: member.d,
          ...(member.fillRule ? { fillRule: member.fillRule } : {}),
          ...(member.fill ? { fill: member.fill } : {}),
          ...(member.line ? { line: member.line } : {}),
        };
      case "textBox":
        return {
          kind: "textBox",
          x: member.x,
          y: member.y,
          width: member.width,
          height: member.height,
          ...(member.nowrap ? { nowrap: true } : {}),
          ...(member.insets ? { insets: member.insets } : {}),
          ...(member.rotation ? { rotation: member.rotation } : {}),
          blocks: member.runs.map(runBlock),
        };
    }
  });
}
