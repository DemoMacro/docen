// The PPTX projection: PresentationOptions → flat slide-absolute drawing
// members the core painter paints. Batch 1 covers the vector surface —
// shapes, pictures, lines/connectors, groups, shape text and chart frames —
// through the same format-neutral extraction helpers the docx projection
// shares (@docen/core/geometry). Batch gaps: table/smartart/media-frame
// children (no member kind / painter wiring yet), a shape's text stays
// unspun when the shape itself rotates, group-level rotation, connector
// presets whose endpoints run reversed, live field evaluation — each lands
// with its follow-up batch.

import {
  colorOf,
  fillOpacityOf,
  lineEndMembersOf,
  linePathData,
  measureEmu,
  outlineOf,
  outerShadowOf,
  presetShapePaths,
  presetShapeTextRect,
  solidFillOf,
} from "@docen/core/geometry";
import {
  emuToPx,
  ptToPx,
  type LayoutDrawingLine,
  type LayoutDrawingMember,
  type LayoutDrawingShadow,
  type LayoutInline,
  type LayoutPictureCrop,
  type LayoutParagraph,
  type LayoutTextStyle,
} from "@docen/layout";
import type { GeometryGuide } from "@office-open/core";
import type {
  PresetGeometryOptions,
  RunFont,
  ShapeType,
  SourceRectangleOptions,
  TextBodyOptions,
  TextCharacterPropertiesOptions,
  TextFont,
} from "@office-open/core/drawing";
import type {
  ConnectorOptions,
  GroupOptions,
  LineShapeOptions,
  PictureOptions,
  PresentationOptions,
  ShapeOptions,
  SlideChild,
  SlideOptions,
} from "@office-open/pptx";

// ── projected shape ──

/** One projected presentation: the slide size plus each slide's members. */
export interface ProjectedPresentation {
  /** Slide width in px (96 dpi; EMU ÷ 9525). */
  widthPx: number;
  /** Slide height in px. */
  heightPx: number;
  slides: ProjectedSlide[];
}

/** One projected slide. */
export interface ProjectedSlide {
  /** The slide's solid background fill, hex RRGGBB; absent → the painter's
   *  page default. Gradient/picture backgrounds stay unprojected (batch gap). */
  background?: string;
  /** Slide-absolute drawing members, paint order = document order. */
  members: LayoutDrawingMember[];
}

/** Project a parsed presentation into the paintable shape: sizes are resolved
 *  to px and every slide's children flatten into members. Pure — the input is
 *  not mutated, so callers may re-project on demand. */
export function projectPresentation(pres: PresentationOptions): ProjectedPresentation {
  const { widthPx, heightPx } = slideSizePx(pres.size);
  return { widthPx, heightPx, slides: (pres.slides ?? []).map(projectSlide) };
}

const NAMED_SLIDE_PX = {
  "16:9": { width: 1280, height: 720 },
  "4:3": { width: 960, height: 720 },
} as const;

/** The named slide classes at 96 dpi; an explicit size resolves its EMU.
 *  Absent → 16:9 (PowerPoint's default). */
function slideSizePx(size: PresentationOptions["size"]): { widthPx: number; heightPx: number } {
  if (size === "4:3")
    return { widthPx: NAMED_SLIDE_PX["4:3"].width, heightPx: NAMED_SLIDE_PX["4:3"].height };
  if (size === "16:9" || size === undefined)
    return { widthPx: NAMED_SLIDE_PX["16:9"].width, heightPx: NAMED_SLIDE_PX["16:9"].height };
  return {
    widthPx: emuToPx(measureEmu(size.width) ?? 0),
    heightPx: emuToPx(measureEmu(size.height) ?? 0),
  };
}

function projectSlide(slide: SlideOptions): ProjectedSlide {
  return {
    ...(slide.background ? { background: solidFillOf(slide.background.fill) } : {}),
    members: childMembers(slide.children ?? [], IDENTITY, []),
  };
}

// ── children walk ──

/** An ancestor group chain's affine map: child boxes are computed in their
 *  own EMU space, then mapped through the group's chOff/chExt scaling. */
interface Xform {
  sx: number;
  sy: number;
  dx: number;
  dy: number;
}
const IDENTITY: Xform = { sx: 1, sy: 1, dx: 0, dy: 0 };

/** A measure field to px: EMU passes through (÷9525), a UM string resolves.
 *  Missing extents land at 0 — the painter clips degenerate boxes. */
function emuOf(v: unknown): number {
  return emuToPx(measureEmu(v) ?? 0);
}

function childMembers(
  children: readonly SlideChild[],
  t: Xform,
  prefix: readonly number[],
): LayoutDrawingMember[] {
  const out: LayoutDrawingMember[] = [];
  children.forEach((child, i) => {
    // The group-address path only exists below a group — a slide's own
    // members have nothing above them to address into.
    const cp = prefix.length > 0 ? [...prefix, i] : undefined;
    if ("shape" in child) {
      out.push(...shapeMembers(child.shape, t, cp));
    } else if ("picture" in child) {
      const m = pictureMember(child.picture, t, cp);
      if (m) out.push(m);
    } else if ("line" in child) {
      out.push(...endpointMembers(child.line, t, cp, {}));
    } else if ("connector" in child) {
      out.push(
        ...endpointMembers(
          child.connector,
          t,
          cp,
          geometryOf(child.connector.properties?.geometry),
        ),
      );
    } else if ("group" in child) {
      out.push(...groupMembers(child.group, t, [...prefix, i]));
    } else if ("chart" in child) {
      const c = child.chart;
      out.push({
        kind: "chart",
        x: t.sx * emuOf(c.x) + t.dx,
        y: t.sy * emuOf(c.y) + t.dy,
        width: t.sx * emuOf(c.width),
        height: t.sy * emuOf(c.height),
        ...(cp ? { childPath: cp } : {}),
        chart: c,
      });
    }
  });
  return out;
}

function groupMembers(
  group: GroupOptions,
  t: Xform,
  path: readonly number[],
): LayoutDrawingMember[] {
  const x = t.sx * emuOf(group.x) + t.dx;
  const y = t.sy * emuOf(group.y) + t.dy;
  const w = t.sx * emuOf(group.width);
  const h = t.sy * emuOf(group.height);
  // An unset chOff/chExt stringifies to the group box itself (identity).
  const chW = emuOf(group.childExtentWidth);
  const chH = emuOf(group.childExtentHeight);
  const sx = chW > 0 ? w / chW : 1;
  const sy = chH > 0 ? h / chH : 1;
  const inner: Xform = {
    sx,
    sy,
    dx: x - emuOf(group.childOffsetX) * sx,
    dy: y - emuOf(group.childOffsetY) * sy,
  };
  return childMembers(group.children ?? [], inner, path);
}

/** The geometry field's two shapes (bare token / object with adjustment
 *  guides) → the evaluator's (preset, adjustments) pair. */
function geometryOf(g: ShapeType | PresetGeometryOptions | undefined): {
  preset?: string;
  adjustments?: readonly GeometryGuide[];
} {
  if (g == null) return {};
  if (typeof g === "string") return { preset: g };
  return { preset: g.preset, ...(g.adjustmentValues ? { adjustments: g.adjustmentValues } : {}) };
}

// ── shape (p:sp) ──

const BOX_PRESETS = new Set(["rect", "roundRect", "ellipse"]);
const STRAIGHT_PRESETS = new Set(["line", "straightConnector1"]);

// DrawingML bodyPr default insets (lIns/rIns 0.1", tIns/bIns 0.05").
const BODY_INSET_EMU = { left: 91440, top: 45720, right: 91440, bottom: 45720 };

interface ShapePaint {
  x: number;
  y: number;
  w: number;
  h: number;
  preset?: string;
  adjustments?: readonly GeometryGuide[];
  fill?: string;
  opacity?: number;
  line?: LayoutDrawingLine;
  shadow?: LayoutDrawingShadow;
  rotation?: number;
  childPath?: readonly number[];
}

function shapeMembers(
  shape: ShapeOptions,
  t: Xform,
  childPath: readonly number[] | undefined,
): LayoutDrawingMember[] {
  const paint: ShapePaint = {
    x: t.sx * emuOf(shape.x) + t.dx,
    y: t.sy * emuOf(shape.y) + t.dy,
    w: t.sx * emuOf(shape.width),
    h: t.sy * emuOf(shape.height),
    ...geometryOf(shape.properties?.geometry),
    fill: solidFillOf(shape.properties?.fill),
    opacity: fillOpacityOf(shape.properties?.fill),
    line: outlineOf(shape.properties?.outline),
    shadow: outerShadowOf(shape.properties?.effects),
    ...(shape.rotation ? { rotation: shape.rotation } : {}),
    ...(childPath ? { childPath } : {}),
  };
  // A text-carrying shape projects its body (PowerPoint writes an empty a:p
  // into every bare shape — only non-empty content makes a text box).
  if (hasText(shape.textBody)) {
    return [textBoxMember(shape.textBody, paint)];
  }
  const {
    x,
    y,
    w,
    h,
    preset,
    adjustments,
    fill,
    opacity,
    line,
    shadow,
    rotation,
    childPath: cp,
  } = paint;

  // Box presets (or none) paint as the plain shape member the renderer knows.
  if (!preset || BOX_PRESETS.has(preset)) {
    return [
      {
        kind: "shape",
        x,
        y,
        width: w,
        height: h,
        ...(preset && preset !== "rect" ? { preset } : {}),
        ...(fill ? { fill } : {}),
        ...(opacity != null ? { opacity } : {}),
        ...(line ? { line } : {}),
        ...(shadow ? { shadow } : {}),
        ...(rotation ? { rotation } : {}),
        ...(cp ? { childPath: cp } : {}),
      },
    ];
  }
  // A straight line paints the box diagonal; the stroke's end arrows expand
  // into their own fill members beside it.
  if (STRAIGHT_PRESETS.has(preset)) {
    return [
      {
        kind: "path",
        x,
        y,
        width: w,
        height: h,
        d: linePathData(w, h),
        ...(line ? { line } : {}),
        ...(shadow ? { shadow } : {}),
        ...(rotation ? { rotation } : {}),
        ...(cp ? { childPath: cp } : {}),
      },
      ...lineEndMembersOf(line, 0, 0, w, h).map((m) => ({
        ...m,
        x: m.x + x,
        y: m.y + y,
        ...(rotation ? { rotation } : {}),
        ...(cp ? { childPath: cp } : {}),
      })),
    ];
  }
  // Every other preset expands through the ECMA-376 evaluator — one path
  // member per spec path (fill-only/stroke-only layers paint separately).
  // An unknown token falls back to the plain rectangle.
  const outlines = presetShapePaths(preset, w, h, adjustments);
  if (!outlines) {
    return [
      {
        kind: "shape",
        x,
        y,
        width: w,
        height: h,
        ...(fill ? { fill } : {}),
        ...(opacity != null ? { opacity } : {}),
        ...(line ? { line } : {}),
        ...(shadow ? { shadow } : {}),
        ...(rotation ? { rotation } : {}),
        ...(cp ? { childPath: cp } : {}),
      },
    ];
  }
  return outlines.map((o) => ({
    kind: "path",
    x,
    y,
    width: w,
    height: h,
    d: o.d,
    ...(o.fill && fill ? { fill } : {}),
    ...(o.fill && opacity != null ? { opacity } : {}),
    ...(o.stroke && line ? { line } : {}),
    ...(shadow ? { shadow } : {}),
    ...(rotation ? { rotation } : {}),
    ...(cp ? { childPath: cp } : {}),
  }));
}

/** A shape's body counts as text only with non-empty paragraph content —
 *  PowerPoint writes an empty a:p into every bare shape. */
function hasText(body: TextBodyOptions | undefined): body is TextBodyOptions {
  if (!body) return false;
  if (body.text) return true;
  return (body.paragraphs ?? []).some((p) =>
    typeof p === "string" ? p.length > 0 : p.text != null || (p.children?.length ?? 0) > 0,
  );
}

function textBoxMember(body: TextBodyOptions, paint: ShapePaint): LayoutDrawingMember {
  const { w, h, preset, adjustments } = paint;
  const straight = preset != null && STRAIGHT_PRESETS.has(preset);
  // A non-box preset paints its evaluated silhouette under the text (a text
  // ellipse stays an ellipse); the fill layers merge into one d. A straight
  // line skips the evaluator — its d is the box diagonal.
  const outlines = straight
    ? [{ d: linePathData(w, h), fill: false, stroke: true }]
    : preset && !BOX_PRESETS.has(preset)
      ? presetShapePaths(preset, w, h, adjustments)
      : undefined;
  const silhouette =
    outlines
      ?.filter((o) => o.fill)
      .map((o) => o.d)
      .join(" ") || undefined;
  const bp = body.bodyProperties;
  const ins = (v: unknown, def: number) => emuToPx(measureEmu(v) ?? def);
  // Text stacks inside the preset's text rectangle (a circle keeps its words
  // off the rim); the bodyPr insets apply on top of that shrink.
  const tr = preset && !straight ? presetShapeTextRect(preset, w, h, adjustments) : undefined;
  return {
    kind: "textBox",
    x: paint.x,
    y: paint.y,
    width: w,
    height: h,
    ...(preset ? { preset } : {}),
    ...(silhouette ? { d: silhouette } : {}),
    ...(paint.fill ? { fill: paint.fill } : {}),
    ...(paint.opacity != null ? { opacity: paint.opacity } : {}),
    ...(paint.line ? { line: paint.line } : {}),
    ...(paint.shadow ? { shadow: paint.shadow } : {}),
    ...(paint.childPath ? { childPath: paint.childPath } : {}),
    insets: {
      left: ins(bp?.lIns, BODY_INSET_EMU.left) + (tr ? Math.max(0, tr.l) : 0),
      top: ins(bp?.tIns, BODY_INSET_EMU.top) + (tr ? Math.max(0, tr.t) : 0),
      right: ins(bp?.rIns, BODY_INSET_EMU.right) + (tr ? Math.max(0, w - tr.r) : 0),
      bottom: ins(bp?.bIns, BODY_INSET_EMU.bottom) + (tr ? Math.max(0, h - tr.b) : 0),
    },
    anchor: bp?.anchor === "center" || bp?.anchor === "bottom" ? bp.anchor : "top",
    ...(bp?.vertical === "vertical" || bp?.vertical === "vertical270"
      ? { textVertical: bp.vertical }
      : {}),
    ...(bp?.spAutoFit === true ? { autoFit: true } : {}),
    blocks: textBlocks(body),
  };
}

// ── picture (p:pic) ──

const PIC_MIME: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
};

/** Bytes → base64 (btoa is universal: browsers and Node ≥ 16). */
function base64Of(bytes: Uint8Array): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return btoa(bin);
}

/** a:srcRect integer-percent insets → the crop fractions the member carries;
 *  an all-zero rect is no crop. */
function cropOf(sr: SourceRectangleOptions | undefined): LayoutPictureCrop | undefined {
  const l = sr?.left ?? 0;
  const t = sr?.top ?? 0;
  const r = sr?.right ?? 0;
  const b = sr?.bottom ?? 0;
  if (!l && !t && !r && !b) return undefined;
  return { left: l / 100, top: t / 100, right: r / 100, bottom: b / 100 };
}

function pictureMember(
  pic: PictureOptions,
  t: Xform,
  childPath: readonly number[] | undefined,
): LayoutDrawingMember | undefined {
  // emf/wmf have no browser decoder — an empty frame (registered gap).
  const mime = PIC_MIME[pic.type];
  const crop = cropOf(pic.sourceRectangle);
  const line = outlineOf(pic.outline);
  const shadow = outerShadowOf(pic.effects);
  // The base contract allows three data shapes: a data URL passes through
  // verbatim, bare base64 and bytes get the mime wrapper (none exists for
  // emf/wmf — nothing the browser could decode anyway).
  const src =
    typeof pic.data === "string"
      ? pic.data.startsWith("data:")
        ? pic.data
        : mime
          ? `data:${mime};base64,${pic.data}`
          : undefined
      : mime && pic.data instanceof Uint8Array
        ? `data:${mime};base64,${base64Of(pic.data)}`
        : undefined;
  return {
    kind: "picture",
    x: t.sx * emuOf(pic.x) + t.dx,
    y: t.sy * emuOf(pic.y) + t.dy,
    width: t.sx * emuOf(pic.width),
    height: t.sy * emuOf(pic.height),
    ...(src ? { src } : {}),
    ...(pic.flipHorizontal ? { flipH: true } : {}),
    ...(pic.flipVertical ? { flipV: true } : {}),
    ...(pic.rotation ? { rotation: pic.rotation } : {}),
    ...(crop ? { crop } : {}),
    ...(line ? { line } : {}),
    ...(shadow ? { shadow } : {}),
    ...(childPath ? { childPath } : {}),
  };
}

// ── line (p:sp) / connector (p:cxnSp) ──

function endpointMembers(
  o: LineShapeOptions | ConnectorOptions,
  t: Xform,
  childPath: readonly number[] | undefined,
  g: { preset?: string; adjustments?: readonly GeometryGuide[] },
): LayoutDrawingMember[] {
  const x1 = emuOf(o.x1);
  const y1 = emuOf(o.y1);
  const x2 = emuOf(o.x2);
  const y2 = emuOf(o.y2);
  const x = t.sx * Math.min(x1, x2) + t.dx;
  const y = t.sy * Math.min(y1, y2) + t.dy;
  const w = t.sx * Math.abs(x2 - x1);
  const h = t.sy * Math.abs(y2 - y1);
  const line = outlineOf(o.properties?.outline);
  const fill = solidFillOf(o.properties?.fill);
  const shadow = outerShadowOf(o.properties?.effects);
  const p = (v: number): string => String(Math.round(v * 100) / 100);
  // Endpoints carry the direction (parse resolved the xfrm flips into them):
  // the segment runs between the corners the endpoints sit on.
  const lx1 = x1 <= x2 ? 0 : w;
  const ly1 = y1 <= y2 ? 0 : h;
  const d = `M ${p(lx1)} ${p(ly1)} L ${p(w - lx1)} ${p(h - ly1)}`;
  const cp = () => (childPath ? { childPath } : {});

  // A bent/elbow connector expands through the evaluator — but an evaluated
  // path always runs (0,0)→(w,h), and a reversed endpoint pair would need a
  // mirror the member model has no slot for, so those stay the diagonal
  // (registered gap).
  const outlines =
    g.preset && !STRAIGHT_PRESETS.has(g.preset) && x1 <= x2 && y1 <= y2
      ? presetShapePaths(g.preset, w, h, g.adjustments)
      : undefined;
  if (outlines) {
    return outlines.map((part) => ({
      kind: "path",
      x,
      y,
      width: w,
      height: h,
      d: part.d,
      ...(part.fill && fill ? { fill } : {}),
      ...(part.stroke && line ? { line } : {}),
      ...(shadow ? { shadow } : {}),
      ...cp(),
    }));
  }
  return [
    {
      kind: "path",
      x,
      y,
      width: w,
      height: h,
      d,
      ...(line ? { line } : {}),
      ...(shadow ? { shadow } : {}),
      ...cp(),
    },
    ...lineEndMembersOf(line, lx1, ly1, w - lx1, h - ly1).map((m) => ({
      ...m,
      x: m.x + x,
      y: m.y + y,
      ...cp(),
    })),
  ];
}

// ── shape text ──

// TextAlignment → the layout's align tokens: justify's spellings land on
// "both", the distributed pair on "distribute".
const ALIGN_MAP = {
  left: "left",
  center: "center",
  right: "right",
  justify: "both",
  lowJustification: "both",
  distribute: "distribute",
  thaiDistributed: "distribute",
} as const;

// PowerPoint's default run size (a:sz 1800) and face (the theme minor font —
// master/lstStyle defaults are a batch gap; every run carries explicit props
// in common decks).
const DEFAULT_FAMILY = "Calibri";
const DEFAULT_SIZE_PX = ptToPx(18);

/** A body's paragraphs as projected blocks; the paragraph shape (not the
 *  LayoutBlock union) so consumers read align/spacing directly. */
function textBlocks(body: TextBodyOptions): LayoutParagraph[] {
  const paragraphs = body.paragraphs ?? (body.text != null ? [body.text] : []);
  return paragraphs.map((p) => {
    const para = typeof p === "string" ? { text: p } : p;
    const props = para.properties;
    const inline: LayoutInline[] = [];
    const kids = para.children ?? (para.text != null ? [para.text] : []);
    for (const kid of kids) {
      if (typeof kid === "string") {
        inline.push({ kind: "text", text: kid, style: runStyle(undefined) });
      } else if ("break" in kid) {
        inline.push({ kind: "break" });
      } else if ("type" in kid) {
        // a:fld paints its cached display text; live evaluation (slidenum,
        // datetime) lands with a field engine — batch gap.
        inline.push({ kind: "text", text: kid.text ?? "", style: runStyle(kid.properties) });
      } else {
        inline.push({ kind: "text", text: kid.text ?? "", style: runStyle(kid) });
      }
    }
    return {
      kind: "paragraph",
      inline,
      ...(props?.alignment ? { align: ALIGN_MAP[props.alignment] } : {}),
      ...(props?.spaceBefore != null || props?.spaceAfter != null
        ? {
            spacing: {
              beforePx: ptToPx(props.spaceBefore ?? 0),
              afterPx: ptToPx(props.spaceAfter ?? 0),
            },
          }
        : {}),
      // An empty paragraph renders as a blank line — the strut keeps its
      // height (the default style supplies the measuring font).
      ...(inline.length === 0 ? { defaultTextStyle: runStyle(undefined) } : {}),
    };
  });
}

function runStyle(rp: TextCharacterPropertiesOptions | undefined): LayoutTextStyle {
  const family = familyOf(rp?.font);
  const color = colorOf(rp?.fill);
  return {
    family: family ?? DEFAULT_FAMILY,
    sizePx: rp?.size != null ? ptToPx(rp.size) : DEFAULT_SIZE_PX,
    ...(rp?.bold ? { bold: true } : {}),
    ...(rp?.italic ? { italic: true } : {}),
    ...(rp?.underline && rp.underline !== "none" ? { underline: true } : {}),
    ...(rp?.strike === "singleStrike" || rp?.strike === "doubleStrike"
      ? { strikethrough: true }
      : {}),
    ...(color ? { color } : {}),
  };
}

/** A bare string sets latin + ea to the same face; the object form reads
 *  latin first, then eastAsia (each slot is a face string or a full font). */
function familyOf(font: RunFont | undefined): string | undefined {
  if (typeof font === "string") return font || undefined;
  const face = (f: TextFont | undefined): string | undefined =>
    typeof f === "string" ? f || undefined : f?.typeface;
  return face(font?.latin) ?? face(font?.eastAsia);
}
