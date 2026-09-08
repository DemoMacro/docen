import type { Node as PMNode } from "@tiptap/pm/model";
import { NodeSelection, type EditorState, type Selection } from "@tiptap/pm/state";

/** The drawing under a NodeSelection — which context tab its selection calls
 *  for: "picture" (Picture Tools), "shape"/"group" (Drawing Tools). Null on
 *  any other selection (text, table, math, …). */
export function drawingSelectionKind(state: EditorState): "picture" | "shape" | "group" | null {
  const sel = state.selection;
  if (!(sel instanceof NodeSelection)) return null;
  const name = sel.node.type.name;
  if (name === "image") return "picture";
  if (name === "wpsShape") return "shape";
  if (name === "wpgGroup") return "group";
  return null;
}

/** What hit→position resolution needs from the host editor: the document, the
 *  live selection (the entered-group gate reads it), and the caret map's
 *  paragraph pairing. */
export interface DrawingDocxHost {
  doc: PMNode;
  selection: Selection;
  posOfPara(para: unknown): number | null;
}

/** The PM node position of a drawing hit's target — the host paragraph's
 *  inner position via the caret map, then the index-th node of the hit's
 *  kind: "drawing" counts floating pictures + wps shapes + wpg groups
 *  (projectDrawings' run order = the paragraph's content order), "inline"
 *  counts the paragraph's non-floating images (the line items' picture
 *  order). A hit with a childPath targets a group member: it resolves to
 *  the member only while the group is entered (the selection sits inside
 *  the group's span — Word: a click selects the group until then). */
export function drawingNodePos(
  host: DrawingDocxHost,
  para: unknown,
  index: number,
  kind: "drawing" | "inline",
  childPath?: readonly number[],
): number | null {
  const innerPos = host.posOfPara(para);
  if (innerPos == null) return null;
  const parentNode = host.doc.nodeAt(innerPos - 1);
  if (!parentNode) return null;
  let seen = 0;
  let hit = -1;
  parentNode.forEach((child, offset) => {
    const target =
      kind === "drawing"
        ? child.type.name === "wpsShape" ||
          child.type.name === "wpgGroup" ||
          (child.type.name === "image" && child.attrs.floating != null)
        : child.type.name === "image" && child.attrs.floating == null;
    if (target && hit < 0 && seen++ === index) hit = innerPos + offset;
  });
  if (hit < 0) return null;
  if (!childPath?.length) return hit;
  const group = host.doc.nodeAt(hit);
  if (!group) return null;
  const sel = host.selection;
  const inside = sel instanceof NodeSelection && sel.from >= hit && sel.from < hit + group.nodeSize;
  if (!inside) return hit;
  return descendGroupChild(host.doc, hit, childPath) ?? hit;
}

/** The position of the member a childPath addresses: descend from the group
 *  node at `pos`, one content level per path segment (pos + 1 + the
 *  preceding siblings' node sizes). Null when the path outruns the content
 *  (a stale hit against a re-edited group). */
export function descendGroupChild(
  doc: PMNode,
  pos: number,
  childPath: readonly number[],
): number | null {
  let node = doc.nodeAt(pos);
  if (!node) return null;
  let at = pos;
  for (const i of childPath) {
    if (i < 0 || i >= node.childCount) return null;
    let offset = 0;
    for (let k = 0; k < i; k++) offset += node.child(k).nodeSize;
    at += 1 + offset;
    node = node.child(i);
  }
  return at;
}
