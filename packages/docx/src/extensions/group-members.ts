import { encodeBase64 } from "@office-open/core";
import { convertEmuToPixels } from "@office-open/core/util";
import type {
  GroupChildMediaData,
  GroupMediaData,
  GroupOptions,
  MediaDataTransformation,
  MediaTransformation,
  ParagraphOptions,
  ShapeOptions,
} from "@office-open/docx";

import { cleanAttrs } from "../converters/styles";
import type { JSONContent } from "../core";
import { mediaOfSrc } from "./image";
import { MEDIA_INLINE_LIMIT, registerMediaBlob } from "./media-registry";
import type { ResolveContext } from "./types";
import { resolveWpsShape } from "./wps-shape";

/**
 * Group member mapping — the two legs between office-open's
 * `GroupChildMediaData[]` (the wpg group's children, laid out in the group's
 * child coordinate space) and the wpgGroup node's PM content (wpsShape /
 * image / wpgGroup member nodes).
 *
 * A member that has no PM node of its own (a chart, a content part) rides in
 * an inlinePassthrough atom instead — the same shape the top-level paragraph
 * run stream uses for those children — so the group stays editable around it.
 * Only a member whose bytes/payload are genuinely missing falls the whole
 * group back to the generic passthrough (a parse artifact, not authored
 * data).
 */

const isRecord = (v: unknown): v is Record<string, unknown> => typeof v === "object" && v !== null;

/** A group child's embedded bytes (BaseMediaEntry.data) — the only form the
 *  data-URL rebuild accepts. */
function childBytesOf(child: object): Uint8Array | null {
  const data = (child as { data?: unknown }).data;
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  return null;
}

/** MediaDataTransformation (the group child space: offset.emus + emus/pixels
 *  dual track) → MediaTransformation (the attrs form the wpsShape carrier
 *  reads: offset.{left,top} + width/height, native EMU). */
function runTransformOf(t: MediaDataTransformation): MediaTransformation {
  const out: MediaTransformation = {
    width: t.emus.x,
    height: t.emus.y,
  };
  if (t.offset?.emus) out.offset = { left: t.offset.emus.x, top: t.offset.emus.y };
  if (t.flipHorizontal === true) out.flipHorizontal = true;
  if (t.flipVertical === true) out.flipVertical = true;
  if (t.rotation !== undefined) out.rotation = t.rotation;
  if (t.effectExtent) out.effectExtent = t.effectExtent;
  return out;
}

/** MediaTransformation (attrs form) → MediaDataTransformation (group child
 *  space). The pixels track is derived (round of the EMU → px conversion,
 *  office-open's createTransformation formula) — neither stringify nor the
 *  layout projection reads it, so the EMU pair is the single source of truth
 *  the round-trip keeps. */
function dataTransformOf(t: MediaTransformation): MediaDataTransformation {
  const left = typeof t.offset?.left === "number" ? t.offset.left : 0;
  const top = typeof t.offset?.top === "number" ? t.offset.top : 0;
  const width = t.width as number;
  const height = t.height as number;
  return {
    offset: {
      emus: { x: left, y: top },
      pixels: {
        x: Math.round(convertEmuToPixels(left)),
        y: Math.round(convertEmuToPixels(top)),
      },
    },
    emus: { x: width, y: height },
    pixels: {
      x: Math.round(convertEmuToPixels(width)),
      y: Math.round(convertEmuToPixels(height)),
    },
    ...(t.flipHorizontal === true ? { flipHorizontal: true } : {}),
    ...(t.flipVertical === true ? { flipVertical: true } : {}),
    ...(t.rotation !== undefined ? { rotation: t.rotation } : {}),
    ...(t.effectExtent ? { effectExtent: t.effectExtent } : {}),
  };
}

/** One group child → a PM member node, or null when the child carries no
 *  modelable payload (missing bytes, non-record shape data) — the caller
 *  falls the whole group back to passthrough. Chart/content-part children
 *  become inlinePassthrough atoms (no PM node of their own). */
export function groupChildToMemberNode(
  child: GroupChildMediaData,
  ctx: ResolveContext,
): JSONContent | null {
  const t = child.transformation;
  if (!isRecord(t)) return null;

  if (child.type === "wps") {
    const data = child.data;
    if (!isRecord(data)) return null;
    // The child-level fill/outline (GroupCommonMediaData) win over the shape's
    // own — office-open's stringify reads `child.fill ?? data.fill`, so the
    // merge here keeps one authority and round-trips to the same XML.
    return resolveWpsShape(
      {
        ...data,
        fill: child.fill ?? data.fill,
        outline: child.outline ?? data.outline,
        transformation: runTransformOf(t),
      } as ShapeOptions,
      ctx,
    );
  }

  if (child.type === "wpg") {
    const content: JSONContent[] = [];
    for (const inner of child.children ?? []) {
      const node = groupChildToMemberNode(inner, ctx);
      if (!node) return null;
      content.push(node);
    }
    const { children: _omit, ...group } = child;
    return {
      type: "wpgGroup",
      content,
      attrs: { wpgGroup: cleanAttrs(group as Record<string, unknown>) },
    };
  }

  if (child.type === "chart" || child.type === "contentPart") {
    return { type: "inlinePassthrough", attrs: { data: JSON.stringify(child) } };
  }

  // Raster/svg picture member — same payload contract as the top-level image
  // node (embedded bytes → data URL; a linked-only picture has none). The
  // compression field has no image attr (mirroring the top-level picture
  // round-trip) and is not carried.
  const bytes = childBytesOf(child);
  if (!bytes) return null;
  const attrs: Record<string, unknown> = {
    src:
      bytes.byteLength > MEDIA_INLINE_LIMIT
        ? registerMediaBlob(bytes, child.type)
        : `data:image/${child.type};base64,${encodeBase64(bytes)}`,
    // px display size (UI); the EMU child-space box rides groupXfrm.
    width: Math.round(convertEmuToPixels(t.emus.x)),
    height: Math.round(convertEmuToPixels(t.emus.y)),
    groupXfrm: {
      x: t.offset?.emus?.x ?? 0,
      y: t.offset?.emus?.y ?? 0,
      cx: t.emus.x,
      cy: t.emus.y,
    },
  };
  if (t.flipHorizontal === true) attrs.flipH = true;
  if (t.flipVertical === true) attrs.flipV = true;
  if (t.rotation !== undefined) attrs.rotation = t.rotation;
  if (t.effectExtent) attrs.effectExtent = t.effectExtent;
  if (child.sourceRectangle) attrs.crop = child.sourceRectangle;
  if (child.nonVisualProperties) attrs.nonVisualProperties = child.nonVisualProperties;
  if (child.useLocalDpi !== undefined) attrs.useLocalDpi = child.useLocalDpi;
  // A grouped picture's spPr fill/outline ride the group-child extension.
  if (child.fill) attrs.fill = child.fill;
  if (child.outline) attrs.outline = child.outline;
  return { type: "image", attrs };
}

/** A PM member node → one group child. `compileBody` compiles a wpsShape
 *  member's editable text body back to ParagraphOptions (the converter's
 *  compileShapeBody). Returns null for a member with no child representation
 *  — dropped, mirroring the resolve-side fallback. */
export function memberNodeToGroupChild(
  node: JSONContent,
  compileBody: (node: JSONContent) => (ParagraphOptions | string)[],
): GroupChildMediaData | null {
  const attrs = (node.attrs ?? {}) as Record<string, unknown>;

  if (node.type === "wpsShape") {
    const ws = attrs.wpsShape;
    if (!isRecord(ws)) return null;
    const { transformation, children: _body, ...data } = ws;
    if (!isRecord(transformation)) return null;
    // fill/outline compile into the shape data (the child-level extension is
    // an alternative carrier of the same field — one authority is enough).
    return cleanAttrs({
      type: "wps",
      transformation: dataTransformOf(transformation as unknown as MediaTransformation),
      data: { ...data, children: compileBody(node) },
    }) as unknown as GroupChildMediaData;
  }

  if (node.type === "wpgGroup") {
    const g = attrs.wpgGroup;
    if (!isRecord(g)) return null;
    const children: GroupChildMediaData[] = [];
    for (const inner of node.content ?? []) {
      const child = memberNodeToGroupChild(inner, compileBody);
      if (child) children.push(child);
    }
    // Nested-group attrs stay in their child-space form (GroupMediaData minus
    // children) — the transformation round-trips verbatim, no unit loss. The
    // type tag is written, not spread: a hand-authored attrs object may omit it.
    return { ...(g as object), type: "wpg", children } as GroupMediaData;
  }

  if (node.type === "inlinePassthrough") {
    try {
      return JSON.parse((attrs.data as string) ?? "null") as GroupChildMediaData;
    } catch {
      return null;
    }
  }

  if (node.type === "image") {
    const src = attrs.src as string | undefined;
    if (!src) return null;
    // The EMU child-space box is the single source of truth; px attrs are
    // display-only.
    const xfrm = attrs.groupXfrm as { x: number; y: number; cx: number; cy: number } | undefined;
    if (!xfrm || !Number.isFinite(xfrm.cx) || !Number.isFinite(xfrm.cy)) return null;
    const media = mediaOfSrc(attrs, src);
    if (!media) return null;
    return cleanAttrs({
      type: media.type,
      data: media.bytes,
      transformation: dataTransformOf({
        offset: { left: xfrm.x, top: xfrm.y },
        width: xfrm.cx,
        height: xfrm.cy,
        ...(attrs.flipH === true ? { flipHorizontal: true } : {}),
        ...(attrs.flipV === true ? { flipVertical: true } : {}),
        ...(typeof attrs.rotation === "number" ? { rotation: attrs.rotation } : {}),
        ...(attrs.effectExtent ? { effectExtent: attrs.effectExtent } : {}),
      } as MediaTransformation),
      ...(attrs.crop ? { sourceRectangle: attrs.crop } : {}),
      ...(attrs.nonVisualProperties ? { nonVisualProperties: attrs.nonVisualProperties } : {}),
      ...(attrs.useLocalDpi !== null && attrs.useLocalDpi !== undefined
        ? { useLocalDpi: attrs.useLocalDpi }
        : {}),
      ...(attrs.fill ? { fill: attrs.fill } : {}),
      ...(attrs.outline ? { outline: attrs.outline } : {}),
    }) as unknown as GroupChildMediaData;
  }

  return null;
}

/** The top-level wpg group resolve: GroupOptions → a wpgGroup node whose
 *  content is the member sequence. Null (→ the generic passthrough fallback)
 *  when the group has no modelable children — parse artifacts only; authored
 *  groups carry wps/picture/wpg members. */
export function resolveGroupOptions(g: GroupOptions, ctx: ResolveContext): JSONContent | null {
  const children = g?.children;
  if (!Array.isArray(children) || children.length === 0) return null;
  const content: JSONContent[] = [];
  for (const child of children) {
    const node = groupChildToMemberNode(child, ctx);
    if (!node) return null;
    content.push(node);
  }
  const { children: _omit, ...attrs } = g;
  return { type: "wpgGroup", content, attrs: { wpgGroup: cleanAttrs(attrs) } };
}
