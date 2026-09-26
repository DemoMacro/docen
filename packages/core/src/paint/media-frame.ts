import type { LayoutDrawingMember } from "@docen/layout";
import { Path as LeaferPath, Rect, Text, type IGroup } from "leafer-ui";

import type { PaintContext } from "./context";
import { addDecodedImage } from "./image";

// Stable media-frame rendering: poster frames preserve the deck's visual
// composition; missing posters get a readable player. Native playback can
// later replace the badge without changing the projection contract.

type MediaMember = Extract<LayoutDrawingMember, { kind: "mediaFrame" }>;

function playbackBadge(tree: IGroup, member: MediaMember): void {
  const size = Math.min(64, Math.max(28, Math.min(member.width, member.height) * 0.28));
  const x = member.x + (member.width - size) / 2;
  const y = member.y + (member.height - size) / 2;
  tree.add(
    new Rect({
      x,
      y,
      width: size,
      height: size,
      fill: "#202124",
      opacity: 0.72,
      cornerRadius: size / 2,
    }),
  );
  const icon =
    member.media === "video"
      ? `M ${size / 2 - 7} ${size / 2 - 10} L ${size / 2 + 10} ${size / 2} L ${size / 2 - 7} ${size / 2 + 10} Z`
      : `M ${size / 2 - 10} ${size / 2 - 5} h 6 l 7 -7 v 24 l -7 -7 h -6 Z`;
  tree.add(
    new LeaferPath({
      x,
      y,
      path: icon,
      fill: "#FFFFFF",
    }),
  );
}

function controlBar(tree: IGroup, member: MediaMember): void {
  if (member.height < 72 || member.width < 96) return;
  const height = 24;
  const y = member.y + member.height - height - 10;
  const width = Math.min(member.width - 24, 260);
  const x = member.x + (member.width - width) / 2;
  tree.add(
    new Rect({
      x,
      y,
      width,
      height,
      fill: "#202124",
      opacity: 0.72,
      cornerRadius: height / 2,
    }),
  );
  tree.add(
    new LeaferPath({
      x,
      y,
      path: `M 12 ${height / 2 - 5} L 26 ${height / 2} L 12 ${height / 2 + 5} Z`,
      fill: "#FFFFFF",
    }),
  );
  tree.add(
    new Rect({
      x: x + 36,
      y: y + height / 2 - 1,
      width: Math.max(12, width - 52),
      height: 2,
      fill: "#DADCE0",
      cornerRadius: 1,
    }),
  );
  tree.add(
    new Rect({
      x: x + 36,
      y: y + height / 2 - 1,
      width: Math.max(4, (width - 52) * 0.1),
      height: 2,
      fill: "#8AB4F8",
      cornerRadius: 1,
    }),
  );
}

function fileNameLabel(tree: IGroup, member: MediaMember): void {
  if (!member.fileName || member.width < 180 || member.height < 100) return;
  const name = member.fileName.length > 48 ? `…${member.fileName.slice(-47)}` : member.fileName;
  tree.add(
    new Text({
      x: member.x + 10,
      y: member.y + 8,
      width: member.width - 20,
      text: name,
      fontSize: 11,
      fill: "#FFFFFF",
      overflow: "ellipsis",
    }),
  );
}

export function paintMediaFrame(tree: IGroup, member: MediaMember, ctx: PaintContext): void {
  if (member.src) {
    addDecodedImage(tree, member.src, member.x, member.y, member.width, member.height, ctx);
  } else {
    tree.add(
      new Rect({
        x: member.x,
        y: member.y,
        width: member.width,
        height: member.height,
        fill: member.media === "audio" ? "#1F2937" : "#111827",
      }),
    );
    fileNameLabel(tree, member);
  }
  playbackBadge(tree, member);
  controlBar(tree, member);
  tree.add(
    new Rect({
      x: member.x,
      y: member.y,
      width: member.width,
      height: member.height,
      fill: "transparent",
      stroke: "#5F6368",
      strokeWidth: 1,
    }),
  );
}
