// Bridge from the leafer-x-metafile plugin's neutral member types to the
// layout engine's drawing members. picture/shape/path members are structural
// subsets of their layout counterparts (the layout types carry the extra
// editor-side fields: flips, opacity, cap/join) — passed through as-is. Plugin
// text runs become one single-run paragraph block each (the layout engine lays
// blocks, the metafile replay only knows runs).

import type { LayoutBlock, LayoutDrawingMember, LayoutTextStyle } from "@docen/layout";
import type { MetafileMember, MetafileTextRun } from "leafer-x-metafile";

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
 *  wrap into paragraph blocks). Members come from the replay cache and are
 *  shared by reference — the projection treats them as read-only. */
export function toLayoutMembers(members: MetafileMember[]): LayoutDrawingMember[] {
  return members.map(
    (member): LayoutDrawingMember =>
      member.kind === "textBox"
        ? {
            kind: "textBox",
            x: member.x,
            y: member.y,
            width: member.width,
            height: member.height,
            ...(member.nowrap ? { nowrap: true } : {}),
            ...(member.insets ? { insets: member.insets } : {}),
            ...(member.rotation ? { rotation: member.rotation } : {}),
            blocks: member.runs.map(runBlock),
          }
        : member,
  );
}
