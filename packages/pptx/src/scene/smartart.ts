// Graphic-frame SmartArt (p:graphicFrame dgm:relIds) projection: the frame
// plus the parsed node tree and layout family. The full DGM layout engine is
// a later batch; the painter consumes this normalized fallback payload.

import type { LayoutDrawingMember } from "@docen/layout";
import type {
  DataModelConnectionOptions,
  DataModelPointMemberOptions,
  DataModelPointOptions,
  TreeNode,
} from "@office-open/core/smartart";
import type { SmartArtOptions } from "@office-open/pptx";

import { emuOf, type Xform } from "./geometry";

function isDataModelPoint(value: DataModelPointMemberOptions): value is DataModelPointOptions {
  return !("choices" in value);
}

function dataModelTextOf(point: DataModelPointOptions): string {
  if (point.text) return point.text;
  const body = point.textBody;
  if (body?.text) return body.text;
  return (body?.paragraphs ?? [])
    .map((paragraph) => (typeof paragraph === "string" ? paragraph : (paragraph.text ?? "")))
    .join("");
}

function dataModelPointsOf(
  members: readonly DataModelPointMemberOptions[],
): DataModelPointOptions[] {
  return members
    .flatMap((member) =>
      isDataModelPoint(member)
        ? [member]
        : [member.choices[0]?.point ?? member.fallback ?? undefined],
    )
    .filter((point): point is DataModelPointOptions => point != null);
}

/** Derive the painter's editable tree when only the complete 0.15 data model
 *  is present (the common parsed form; authoring may supply nodes directly). */
function dataModelNodesOf(
  points: readonly DataModelPointOptions[],
  connections: readonly DataModelConnectionOptions[],
): TreeNode[] {
  const nodes = new Map<string, { text: string; children: TreeNode[] }>(
    points
      .filter((point) => point.type === "node" || point.type === undefined)
      .map((point) => [point.modelId, { text: dataModelTextOf(point), children: [] }]),
  );
  const parents = new Map<string, string>();
  for (const connection of connections) {
    if (connection.type === "parOf" || connection.type === undefined)
      parents.set(connection.destinationId, connection.sourceId);
  }
  for (const [id, node] of nodes) {
    const parent = nodes.get(parents.get(id) ?? "");
    if (parent?.children) parent.children.push(node);
  }
  return [...nodes].filter(([id]) => (parents.get(id) ?? "0") === "0").map(([, node]) => node);
}

export function smartArtMember(
  smartArt: SmartArtOptions,
  t: Xform,
  childPath: readonly number[] | undefined,
): LayoutDrawingMember {
  const layout =
    typeof smartArt.layout === "string"
      ? smartArt.layout
      : typeof smartArt.layout?.uniqueId === "string"
        ? (smartArt.layout.uniqueId.split("/").pop() ?? undefined)
        : undefined;
  return {
    kind: "smartArt",
    x: t.sx * emuOf(smartArt.x) + t.dx,
    y: t.sy * emuOf(smartArt.y) + t.dy,
    width: t.sx * emuOf(smartArt.width),
    height: t.sy * emuOf(smartArt.height),
    ...(layout ? { layout } : {}),
    ...(childPath ? { childPath } : {}),
    nodes:
      smartArt.nodes.length > 0
        ? smartArt.nodes
        : dataModelNodesOf(
            dataModelPointsOf(smartArt.dataModel?.points ?? []),
            smartArt.dataModel?.connections ?? [],
          ),
  };
}
