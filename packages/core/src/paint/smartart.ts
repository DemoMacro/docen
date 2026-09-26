import type { LayoutDrawingMember, LayoutSmartArtNode } from "@docen/layout";
import { Ellipse, Path as LeaferPath, Rect, Text, type IGroup } from "leafer-ui";

// Stable SmartArt fallback: the common layout families get real shapes and
// connections; unknown families use an evenly distributed grid. This is not
// PowerPoint's DGM constraint solver, but it keeps parsed decks readable and
// the frame selectable rather than dropping SmartArt on the floor.

type SmartArtMember = Extract<LayoutDrawingMember, { kind: "smartArt" }>;

interface NodeBox {
  x: number;
  y: number;
  width: number;
  height: number;
  text: string;
  fill: string;
}

const ACCENTS = ["4472C4", "ED7D31", "70AD47", "FFC000", "5B9BD5", "A5A5A5"];
const CHILD_FILL = "8FAADC";
const LINK = "#8496B0";

const nodesOf = (nodes: readonly LayoutSmartArtNode[] | undefined): LayoutSmartArtNode[] =>
  Array.isArray(nodes) ? nodes : [];

const textOf = (node: LayoutSmartArtNode, index: number): string =>
  node.text || `Item ${index + 1}`;

function familyOf(layout: string | undefined): "process" | "cycle" | "hierarchy" | "list" | "grid" {
  const family = layout?.toLowerCase() ?? "";
  if (family.includes("process") || family.includes("chevron") || family.includes("arrow"))
    return "process";
  if (family.includes("cycle") || family.includes("radial") || family.includes("venn"))
    return "cycle";
  if (family.includes("hierarchy") || family.includes("orgchart")) return "hierarchy";
  if (family.includes("list")) return "list";
  return "grid";
}

const isHierarchy = (family: string): boolean => family === "hierarchy";

function leavesOf(nodes: readonly LayoutSmartArtNode[]): number {
  if (nodes.length === 0) return 1;
  return nodes.reduce(
    (sum, node) => sum + (node.children?.length ? leavesOf(node.children) : 1),
    0,
  );
}

function depthOf(nodes: readonly LayoutSmartArtNode[]): number {
  if (nodes.length === 0) return 0;
  return 1 + Math.max(0, ...nodes.map((node) => depthOf(node.children ?? [])));
}

function boxesFor(member: SmartArtMember): {
  boxes: NodeBox[];
  links: { x1: number; y1: number; x2: number; y2: number }[];
} {
  const nodes = nodesOf(member.nodes);
  if (nodes.length === 0) return { boxes: [], links: [] };
  const family = familyOf(member.layout);
  const gap = Math.min(18, Math.max(8, member.width * 0.03));

  if (family === "process" || (family === "list" && !member.layout?.startsWith("v"))) {
    const width = Math.max(24, (member.width - gap * (nodes.length - 1)) / nodes.length);
    const height = Math.min(84, member.height * 0.48);
    const y = (member.height - height) / 2;
    return {
      boxes: nodes.map((node, index) => ({
        x: index * (width + gap),
        y,
        width,
        height,
        text: textOf(node, index),
        fill: ACCENTS[index % ACCENTS.length]!,
      })),
      links: [],
    };
  }

  if (family === "list") {
    const height = Math.max(24, (member.height - gap * (nodes.length - 1)) / nodes.length);
    const width = Math.min(520, member.width);
    const x = (member.width - width) / 2;
    return {
      boxes: nodes.map((node, index) => ({
        x,
        y: index * (height + gap),
        width,
        height,
        text: textOf(node, index),
        fill: ACCENTS[index % ACCENTS.length]!,
      })),
      links: [],
    };
  }

  if (family === "cycle") {
    const width = Math.min(150, Math.max(48, member.width / 3 - gap));
    const height = Math.min(84, Math.max(36, member.height / 3 - gap));
    const cx = member.width / 2;
    const cy = member.height / 2;
    const rx = Math.max(0, (member.width - width) / 2);
    const ry = Math.max(0, (member.height - height) / 2);
    return {
      boxes: nodes.map((node, index) => {
        const angle = -Math.PI / 2 + (index / nodes.length) * Math.PI * 2;
        return {
          x: cx + Math.cos(angle) * rx - width / 2,
          y: cy + Math.sin(angle) * ry - height / 2,
          width,
          height,
          text: textOf(node, index),
          fill: ACCENTS[index % ACCENTS.length]!,
        };
      }),
      links: [],
    };
  }

  if (isHierarchy(family)) {
    const boxes: NodeBox[] = [];
    const links: { x1: number; y1: number; x2: number; y2: number }[] = [];
    const levels = Math.max(1, depthOf(nodes));
    const slots = leavesOf(nodes);
    const slotWidth = member.width / slots;
    const width = Math.min(170, slotWidth - gap);
    const height = Math.max(28, Math.min(68, (member.height - gap * (levels - 1)) / levels));

    const place = (node: LayoutSmartArtNode, level: number, from: number, to: number): number => {
      const children = node.children ?? [];
      const childFrom = from;
      let cursor = childFrom;
      const childCenters: number[] = [];
      for (const child of children) {
        const span = leavesOf([child]);
        const center = place(child, level + 1, cursor, cursor + span - 1);
        childCenters.push(center);
        cursor += span;
      }
      const slot = (from + to) / 2;
      const x = slot * slotWidth + (slotWidth - width) / 2;
      const y = level * (height + gap);
      const index = boxes.length;
      boxes.push({
        x,
        y,
        width,
        height,
        text: textOf(node, index),
        fill: children.length > 0 ? ACCENTS[0]! : CHILD_FILL,
      });
      for (const childCenter of childCenters) {
        links.push({
          x1: x + width / 2,
          y1: y + height,
          x2: childCenter,
          y2: y + height + gap,
        });
      }
      return x + width / 2;
    };

    let cursor = 0;
    for (const node of nodes) {
      const span = leavesOf([node]);
      place(node, 0, cursor, cursor + span - 1);
      cursor += span;
    }
    return { boxes, links };
  }

  const columns = Math.min(nodes.length, Math.max(1, Math.ceil(Math.sqrt(nodes.length))));
  const rows = Math.ceil(nodes.length / columns);
  const width = Math.max(32, (member.width - gap * (columns - 1)) / columns);
  const height = Math.max(28, (member.height - gap * (rows - 1)) / rows);
  return {
    boxes: nodes.map((node, index) => {
      const column = index % columns;
      const row = Math.floor(index / columns);
      const rowItems = Math.min(columns, nodes.length - row * columns);
      const rowWidth = rowItems * width + (rowItems - 1) * gap;
      return {
        x: (member.width - rowWidth) / 2 + column * (width + gap),
        y: row * (height + gap),
        width,
        height,
        text: textOf(node, index),
        fill: ACCENTS[index % ACCENTS.length]!,
      };
    }),
    links: [],
  };
}

export function paintSmartArt(tree: IGroup, member: SmartArtMember): void {
  const nodes = nodesOf(member.nodes);
  const family = familyOf(member.layout);
  const { boxes, links } = boxesFor(member);

  if (nodes.length === 0) {
    tree.add(
      new Rect({
        x: member.x,
        y: member.y,
        width: member.width,
        height: member.height,
        fill: "#F3F4F6",
        stroke: LINK,
        strokeWidth: 1,
        cornerRadius: 6,
      }),
    );
    return;
  }

  if (family === "cycle") {
    tree.add(
      new Ellipse({
        x: member.x + member.width * 0.12,
        y: member.y + member.height * 0.18,
        width: member.width * 0.76,
        height: member.height * 0.64,
        fill: "transparent",
        stroke: LINK,
        strokeWidth: 1.5,
      }),
    );
  }

  for (const link of links) {
    tree.add(
      new LeaferPath({
        x: member.x,
        y: member.y,
        path: `M ${link.x1} ${link.y1} L ${link.x1} ${(link.y1 + link.y2) / 2} L ${link.x2} ${(link.y1 + link.y2) / 2} L ${link.x2} ${link.y2}`,
        stroke: LINK,
        strokeWidth: 1.5,
      }),
    );
  }

  for (const [index, box] of boxes.entries()) {
    tree.add(
      new Rect({
        x: member.x + box.x,
        y: member.y + box.y,
        width: box.width,
        height: box.height,
        fill: `#${box.fill}`,
        stroke: "#FFFFFF",
        strokeWidth: 1,
        cornerRadius: 6,
      }),
    );
    tree.add(
      new Text({
        x: member.x + box.x + 6,
        y: member.y + box.y + 4,
        width: Math.max(12, box.width - 12),
        height: Math.max(14, box.height - 8),
        text: box.text,
        fontSize: Math.max(10, Math.min(15, box.width / 8, box.height / 3)),
        fill: "#FFFFFF",
        textAlign: "center",
        verticalAlign: "middle",
        overflow: "ellipsis",
      }),
    );
    if (family === "process" && index < boxes.length - 1) {
      const current = boxes[index]!;
      const next = boxes[index + 1]!;
      const y = member.y + current.y + current.height / 2;
      const x0 = member.x + current.x + current.width;
      const x1 = member.x + next.x;
      tree.add(
        new LeaferPath({
          path: `M ${x0 + 2} ${y} L ${x1 - 8} ${y} M ${x1 - 8} ${y} L ${x1 - 14} ${y - 5} M ${x1 - 8} ${y} L ${x1 - 14} ${y + 5}`,
          stroke: LINK,
          strokeWidth: 1.5,
        }),
      );
    }
  }
}
