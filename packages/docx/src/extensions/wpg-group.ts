import {
  convertEmuToPixels,
  convertEmuToPoints,
  convertUniversalMeasureToEmu,
  convertUniversalMeasureToPt,
  encodeBase64,
} from "@office-open/core";
import type { FillOptions, OutlineOptions, SolidFillOptions } from "@office-open/core";
import type {
  ChildOffset,
  GroupChildMediaData,
  ParagraphOptions,
  WpgGroupRunOptions,
  WpsShapeCoreOptions,
  WpsShapeRunOptions,
} from "@office-open/docx";
import type { DOMOutputSpec } from "@tiptap/pm/model";

import { Node } from "../core";
import type { ParseInlineRule } from "./types";
import { floatAnchorScope, floatingToStyles, normalizeColorToHex, renderRunStyles } from "./utils";

/**
 * wpgGroup — inline atom carrying a DOCX drawing group (wpg: wordprocessingGroup)
 * as an opaque blob. Mirrors the office-open `WpgGroupRunOptions` / ParagraphChild
 * `wpgGroup` field verbatim in attrs.wpgGroup, so the node name and attr stay
 * aligned with the OOXML concept (CT_WordprocessingGroup).
 *
 * A group bundles pictures (pic), shapes (wps), and nested groups (wpg) behind a
 * shared coordinate space (grpSpPr chOff/chExt → extent). renderHTML lays each
 * child out at its transformed position/size so the group renders as Word draws
 * it (e.g. a row of colored decoration bars), instead of a single picture.
 */

/** number is EMU; a string is a UniversalMeasure ("1pt"/"0.5in"/…). → px. */
function measureToPx(v: number | string | undefined): number {
  if (v == null) return 0;
  const emu = typeof v === "number" ? v : convertUniversalMeasureToEmu(v);
  return convertEmuToPixels(emu);
}
/** number is EMU; a string is a UniversalMeasure. → pt (for border widths). */
function measureToPt(v: number | string | undefined, fallback = 0): number {
  if (typeof v === "number") return convertEmuToPoints(v);
  if (typeof v === "string") return convertUniversalMeasureToPt(v);
  return fallback;
}

// ── types reused from @office-open (no hand-written duplicates) ──
//
// Shape/group model types come straight from office-open:
//   WpsShapeRunOptions  — standalone wps text box (wp:anchor > wps:wsp)
//   WpsShapeCoreOptions — wps interior (also a group's wps child .data)
//   WpgGroupRunOptions  — wpg drawing group (CT_WordprocessingGroup)
//   GroupChildMediaData — a group child (wps | wpg | pic media-type union)
//   BodyPropertiesOptions — a:bodyPr (wrap/spAutoFit/vert/anchor/insets/…)
// Fill/outline (FillOptions/OutlineOptions) are discriminated unions; the helpers
// below narrow them to the solid branch docen renders.

/** A standalone wps text-box shape (the editable wpsShape node's payload). */
export type WpsShapeStandalone = WpsShapeRunOptions;
/** wps interior data (fill/outline/bodyProperties/text body). */
export type WpsData = WpsShapeCoreOptions;

type Spec = ReadonlyArray<unknown>;

/** SolidFillOptions union → its `.value` when that is a plain hex string (sRgb).
 *  Scheme/system/preset colors carry theme ids docen can't resolve to hex, so
 *  they fall through to the caller's default. */
function solidColorValue(color: SolidFillOptions | undefined): string | undefined {
  const v = (color as { value?: unknown } | undefined)?.value;
  return typeof v === "string" ? v : undefined;
}

/** Solid fill → CSS color (noFill/none/gradient/pattern → undefined). FillOptions
 *  is `string | { type: "solid" | "none" | … }`; docen renders only solid (a bare
 *  hex string or the solid branch). */
function fillToCss(fill: FillOptions | undefined): string | undefined {
  if (!fill) return undefined;
  if (typeof fill === "string") return normalizeColorToHex(fill);
  if (fill.type !== "solid") return undefined;
  const color = fill.color; // string | SolidFillOptions
  const hex = typeof color === "string" ? color : solidColorValue(color);
  return normalizeColorToHex(hex);
}

/** Shape outline → CSS border (EMU width → pt). noFill → undefined.
 *  OutlineOptions = OutlineAttributes & OutlineFillProperties; OutlineFillProperties
 *  type is "noFill" | "solidFill" | "gradFill" | "pattFill". */
function outlineToCss(outline: OutlineOptions | undefined): string | undefined {
  if (!outline || outline.type === "noFill") return undefined;
  const color = normalizeColorToHex(solidColorValue(outline.color)) ?? "black";
  const width = `${measureToPt(outline.width, 0.75)}pt`;
  const style = outline.dash === "sysDot" || outline.dash === "sysDash" ? "dashed" : "solid";
  return `border:${width} ${style} ${color}`;
}

/** Group extent in px: office-open 0.10.4+ parses wp:extent as EMU verbatim on
 *  MediaTransformation.width/height; convert to px (matching child transforms). */
function groupExtent(group: WpgGroupRunOptions): { w: number; h: number } {
  const t = group.transformation;
  return {
    w: t && typeof t.width === "number" ? convertEmuToPixels(t.width) : 0,
    h: t && typeof t.height === "number" ? convertEmuToPixels(t.height) : 0,
  };
}

/** Child coord (EMU, group's chOff/chExt space) → group-local px. */
function childBox(
  child: GroupChildMediaData,
  chOff: ChildOffset,
  scaleX: number,
  scaleY: number,
): { x: number; y: number; w: number; h: number } {
  const t = child.transformation;
  const off = t?.offset?.emus ?? { x: 0, y: 0 };
  const size = t?.emus ?? { x: 0, y: 0 };
  return {
    x: (off.x - chOff.x) * scaleX,
    y: (off.y - chOff.y) * scaleY,
    w: size.x * scaleX,
    h: size.y * scaleY,
  };
}

/** pic media type + raw bytes → data URL (bytes may arrive as Uint8Array or a
 *  plain object of byte values across the JSON boundary). */
function picSrc(mediaType: string, data: unknown): string | null {
  if (!data) return null;
  const bytes =
    data instanceof Uint8Array
      ? data
      : new Uint8Array(Object.values(data as Record<string, number>));
  return `data:image/${mediaType};base64,${encodeBase64(bytes)}`;
}

function renderChild(
  child: GroupChildMediaData,
  chOff: ChildOffset,
  scaleX: number,
  scaleY: number,
): Spec {
  const box = childBox(child, chOff, scaleX, scaleY);
  // box-sizing:border-box so a shape's width/height is its outer box (matching
  // Word's extent), not content-box (which adds border+padding on top, growing
  // the box ~23×13px past the declared extent).
  const base = `position:absolute;left:${box.x}px;top:${box.y}px;width:${box.w}px;height:${box.h}px;box-sizing:border-box`;

  if (child.type === "wpg") {
    // Nested group: recurse with the box the parent transform mapped it to —
    // a nested group's own transformation is in child-coordinate space (small
    // pixels/emus), so its real size is the parent's mapped box, not its own.
    return renderGroup(child, box.w, box.h, base);
  }

  if (child.type === "wps") {
    // wps shape: a colored rect (office-open default geometry is rect; a present
    // customGeometry would need an SVG node view for path fidelity). A text-box
    // shape also carries body paragraphs (children) rendered as inline text.
    const data = (child.data ?? {}) as WpsShapeCoreOptions;
    // rotation may come from the group transform rather than the shape's bodyPr.
    const rotation = data.bodyProperties?.rotation ?? child.transformation?.rotation;
    return renderWpsInterior(data, base, { rotation, attrs: { "data-wpg-wps": "" } });
  }

  // pic: child.type is the media type (jpg/png/…), child.data is the raw bytes.
  if (child.type) {
    const src = picSrc(child.type, child.data);
    if (src) return ["img", { src, alt: "", style: base }];
  }

  return [];
}

/** wps text-body paragraphs (office-open ParagraphOptions[]) → inline HTML runs
 *  so a text-box shape shows its text. Each paragraph's run defaults merge with
 *  each text run; renderRunStyles converts font/color/size/… to CSS. Advanced run
 *  props (shadow w14RawXml, kern) are carried in attrs but not rendered. */
export function renderWpsText(
  children: readonly (ParagraphOptions | string)[] | undefined,
): Spec[] {
  if (!Array.isArray(children)) return [];
  return children.map((para): Spec => {
    if (typeof para === "string") return ["p", { style: "margin:0" }, para];
    const defaultRun = para.run ?? {};
    const runs = (
      (para.children as readonly (string | ({ text?: string } & Record<string, unknown>))[]) ?? []
    ).map((r): Spec | string => {
      if (typeof r === "string") return r;
      const merged = { ...defaultRun, ...r };
      const css = renderRunStyles(merged as Record<string, unknown>);
      const attrs = css.length ? { style: css.join(";") } : {};
      return ["span", attrs, r.text ?? ""];
    });
    const style = para.alignment ? `margin:0;text-align:${para.alignment}` : "margin:0";
    return ["p", { style }, ...runs];
  });
}

/** Inner style of a wps text-box: fill + outline + textbox insets (padding).
 *  Lives on the contentDOM (the editable interior). Never carries rotation or
 *  writing-mode — those go on the outer positioning wrapper (see
 *  wpsRotationVert), because transform/writing-mode on an editable region
 *  distort the caret rect and break CJK IME composition. */
export function wpsInnerStyle(data: WpsShapeCoreOptions): string {
  const parts: string[] = [];
  const fill = fillToCss(data.fill);
  if (fill) parts.push(`background-color:${fill}`);
  const outline = outlineToCss(data.outline);
  if (outline) parts.push(outline);
  // textbox insets (bodyProperties lIns/tIns/rIns/bIns, EMU → px) keep the text
  // inside the shape's padding, matching Word's text-box insets.
  const bp = data.bodyProperties;
  if (bp) {
    // lIns/tIns/rIns/bIns: EMU numbers (or UniversalMeasure strings) → px.
    const ins = (v: number | string | undefined) => `${measureToPx(v).toFixed(1)}px`;
    parts.push(`padding:${ins(bp.tIns)} ${ins(bp.rIns)} ${ins(bp.bIns)} ${ins(bp.lIns)}`);
  }
  return parts.join(";");
}

/** Rotation + writing-mode (text direction) for a wps shape. Lives on the
 *  positioning wrapper OUTSIDE the contentDOM. `rotationOverride` covers a
 *  group child whose rotation comes from the group transform, not bodyPr. */
export function wpsRotationVert(data: WpsShapeCoreOptions, rotationOverride?: number): string {
  const parts: string[] = [];
  const rotation = rotationOverride ?? data.bodyProperties?.rotation;
  if (rotation) parts.push(`transform:rotate(${rotation}deg)`);
  // text direction (bodyPr vert): vert/eaVert/mongolianVert → vertical-rl,
  // vert270 → vertical-lr, so CJK vertical text boxes render top-to-bottom.
  const vert = data.bodyProperties?.vert;
  if (vert && vert !== "horz") {
    parts.push(`writing-mode:${vert === "vert270" ? "vertical-lr" : "vertical-rl"}`);
  }
  return parts.join(";");
}

/**
 * Render a wps shape's interior (fill/outline/insets/rotation/writing-mode + text
 * body) as a positioned div. Shared by the wpg group's inline wps children and
 * the standalone wpsShape node, so a text-box shape renders identically whether
 * it floats alone (wp:anchor > wps:wsp) or sits inside a group. `positionStyle`
 * carries the placement — absolute group coords for a group child, the floating
 * anchor CSS for a standalone shape. `opts.rotation` overrides bodyPr rotation
 * (a group child's rotation may come from the group transform, not bodyPr).
 *
 * The wpg path merges placement + interior into ONE div (an atom has no
 * contentDOM, so rotation/writing-mode on the element is safe — no caret inside).
 * The editable wpsShape node instead splits these across two elements via
 * wpsShapeStyles (outer = position+rotation+vert, inner = contentDOM).
 */
export function renderWpsInterior(
  data: WpsShapeCoreOptions,
  positionStyle: string,
  opts?: { rotation?: number; attrs?: Record<string, string> },
): Spec {
  const inner = wpsInnerStyle(data);
  const rotVert = wpsRotationVert(data, opts?.rotation);
  const style = [positionStyle, inner, rotVert].filter(Boolean).join(";");
  const attrs: Record<string, unknown> = { style };
  if (opts?.attrs) Object.assign(attrs, opts.attrs);
  return ["div", attrs, ...renderWpsText(data.children)];
}

/** Two-element style split for an editable standalone wpsShape: `outer` for the
 *  positioning wrapper (dom), `inner` for the contentDOM. rotation/writing-mode
 *  stay on `outer` so the editable interior has a clean caret/IME rect. The
 *  geometry (EMU extent → px, floating anchor CSS) is computed here so the
 *  editor's NodeView and generateHTML render identically without re-deriving
 *  the engine's EMU/floating math. */
export function wpsShapeStyles(ws: WpsShapeRunOptions): WpsShapeStyles {
  const tw = ws.transformation?.width;
  const th = ws.transformation?.height;
  const w = typeof tw === "number" ? convertEmuToPixels(tw) : 0;
  const h = typeof th === "number" ? convertEmuToPixels(th) : 0;
  // bodyPr@wrap="none": text never wraps and the shape grows to the text width
  // (Word renders the text box one line wide). Fixed extent width + overflow:
  // hidden would wrap it (extent 192px can't fit a 24pt one-liner) — so drop the
  // fixed width (max-content shrinks to the text) and force nowrap. wrap="square"
  // (default) keeps the extent width and clips overflow (Word's noAutoFit default).
  const noWrap = ws.bodyProperties?.wrap === "none";
  const sizeStyle = noWrap
    ? `width:max-content;height:${h}px;box-sizing:border-box;white-space:nowrap`
    : `width:${w}px;height:${h}px;box-sizing:border-box;overflow:hidden`;
  const rotVert = wpsRotationVert(ws);
  const widthNum = typeof tw === "number" ? tw : undefined;
  let outer: string;
  let paragraphAnchor = false;
  if (ws.floating) {
    // A floating text box (wp:anchor wrapNone) overlays its anchor paragraph
    // instead of claiming a line in the flow — same anchor CSS as images and
    // wpg groups (position:absolute at the EMU offset).
    outer = [...floatingToStyles(ws.floating, undefined, widthNum), sizeStyle, rotVert]
      .filter(Boolean)
      .join(";");
    paragraphAnchor = floatAnchorScope(ws.floating) === "paragraph";
  } else {
    // An inline wps (rare — no wp:anchor) flows with the text as an inline block.
    outer = [`display:inline-block;vertical-align:middle`, sizeStyle, rotVert]
      .filter(Boolean)
      .join(";");
  }
  // Vertical anchor (bodyPr anchor: t/ctr/b) — Word positions the text block at
  // the box's top/middle/bottom. flex column + justify-content reproduces it on
  // the contentDOM; flex lives on the container, not the editable text, so the
  // caret/IME rect is unaffected. height:100% fills the outer box so justify
  // has the full extent to distribute against (otherwise the contentDOM
  // shrinks to its text height and anchor has no room to push within).
  const anchor = ws.bodyProperties?.anchor;
  const justify = anchor === "ctr" ? "center" : anchor === "b" ? "flex-end" : "flex-start";
  const inner = `box-sizing:border-box;display:flex;flex-direction:column;justify-content:${justify};height:100%;${wpsInnerStyle(ws)}`;
  return { outer, inner, paragraphAnchor };
}

export interface WpsShapeStyles {
  /** outer wrapper (dom): floating/inline placement + size + rotation + writing-mode. */
  outer: string;
  /** inner contentDOM: box-sizing + fill + outline + textbox insets (padding). */
  inner: string;
  /** true when the floating anchor resolves against the anchor paragraph
   *  (vRelative "paragraph") — the editor marks the anchor <p> relative. */
  paragraphAnchor: boolean;
}

/** Render a wpg group (top-level or nested) as a positioned container.
 *  actualW/H are the group's real pixel size — top-level groups read them from
 *  transformation.width/height; nested groups receive the box their parent
 *  transform mapped them to. containerStyle (when set) carries the absolute
 *  placement for a nested group; otherwise the group is an inline-block. */
function renderGroup(
  group: {
    children?: GroupChildMediaData[];
    childOffset?: ChildOffset;
    childExtent?: { cx: number; cy: number };
  },
  actualW: number,
  actualH: number,
  containerStyle?: string,
  extraAttrs?: Record<string, string>,
): Spec {
  const chOff = group.childOffset ?? { x: 0, y: 0 };
  const chExt = group.childExtent ?? { cx: 1, cy: 1 };
  const scaleX = actualW / chExt.cx;
  const scaleY = actualH / chExt.cy;

  const style =
    containerStyle ??
    `position:relative;display:inline-block;vertical-align:middle;width:${actualW}px;height:${actualH}px`;
  const children = (group.children ?? [])
    .map((c) => renderChild(c, chOff, scaleX, scaleY))
    .filter((c) => Array.isArray(c) && (c as unknown[]).length > 0);

  // Serialize the full group model so generateHTML→parseHTML round-trips it
  // (parseHTML does JSON.parse(data-wpg-group); an empty string would throw and
  // drop children/offsets/extent/floating/grpSpPr on every HTML round-trip).
  const attrs: Record<string, unknown> = { "data-wpg-group": JSON.stringify(group), style };
  if (extraAttrs) Object.assign(attrs, extraAttrs);
  return ["span", attrs, ...children] as Spec;
}

const attrWpgGroup = () => ({
  default: null,
  rendered: false,
  parseHTML: (element: HTMLElement) => {
    const raw = element.getAttribute("data-wpg-group");
    if (!raw) return null;
    try {
      return JSON.parse(raw);
    } catch {
      return null;
    }
  },
});

// DOCX drawing group (wpg) → opaque atom: full WpgGroupRunOptions rides on
// attrs.wpgGroup (the editor doesn't model the group interior).
export const parseDocxInline: ParseInlineRule = {
  match: (child) => "wpgGroup" in child,
  convert: (child) => ({
    type: "wpgGroup",
    attrs: { wpgGroup: (child as { wpgGroup: unknown }).wpgGroup },
  }),
};

export const WpgGroup = Node.create({
  name: "wpgGroup",
  group: "inline",
  inline: true,
  atom: true,

  addAttributes() {
    return {
      wpgGroup: attrWpgGroup(),
    };
  },

  parseHTML() {
    return [{ tag: "span[data-wpg-group]" }];
  },

  renderHTML({
    node,
  }: {
    node: { attrs: Record<string, unknown> };
    HTMLAttributes: Record<string, unknown>;
  }) {
    const wpg = (node.attrs.wpgGroup as WpgGroupRunOptions | null) ?? ({} as WpgGroupRunOptions);
    const { w, h } = groupExtent(wpg);
    const tw = wpg.transformation?.width;
    const widthNum = typeof tw === "number" ? tw : undefined;
    // A floating group (wp:anchor) overlays its anchor paragraph instead of
    // claiming a line in the flow — render it with the same anchor CSS as
    // images (position:absolute for wrapNone "in front/behind text"), so the
    // group floats over the text below instead of pushing it down.
    if (wpg.floating) {
      const containerStyle = [
        ...floatingToStyles(wpg.floating, undefined, widthNum),
        `width:${w}px`,
        `height:${h}px`,
      ].join(";");
      // A paragraph-anchored wrapNone group resolves its absolute top/left from
      // the anchor <p> (data-float-anchor → editor CSS makes the <p> relative);
      // otherwise it anchors to the page box and floats over the heading/body.
      const extraAttrs =
        floatAnchorScope(wpg.floating) === "paragraph"
          ? { "data-float-anchor": "paragraph" }
          : undefined;
      return renderGroup(wpg, w, h, containerStyle, extraAttrs) as unknown as DOMOutputSpec;
    }
    return renderGroup(wpg, w, h) as unknown as DOMOutputSpec;
  },

  parseDocxInline,
});
