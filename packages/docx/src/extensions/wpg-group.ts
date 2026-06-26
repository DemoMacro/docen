import { encodeBase64 } from "@office-open/core";
import type { DOMOutputSpec } from "@tiptap/pm/model";

import { Node } from "../core";
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

const EMU_PER_PT = 12700;
const EMU_PER_PX = 9525;

// ── opaque group model (attrs.wpgGroup is JSON, typed loosely) ──

interface PointEmu {
  x: number;
  y: number;
}

interface ChildTransformation {
  pixels?: PointEmu; // width/height (child coord space, EMU/9525)
  emus?: PointEmu; // width/height (child coord space, EMU)
  offset?: { emus?: PointEmu; pixels?: PointEmu };
  rotation?: number;
}

interface FillColor {
  value?: string;
}

interface Fill {
  type?: string; // "solid" | "none" | "noFill" | …
  color?: FillColor;
}

interface Outline {
  type?: string; // "solidFill" | "noFill" | "none" | …
  color?: FillColor;
  width?: number; // EMU
  dash?: string;
}

/** wps shape data (WpsShapeCoreOptions subset used for rendering). */
interface WpsBodyProperties {
  rotation?: number;
  // textbox insets (EMU): left/top/right/bottom inner padding
  lIns?: number;
  tIns?: number;
  rIns?: number;
  bIns?: number;
  // text direction (a:bodyPr vert): horz default; vert/vert270/eaVert → vertical
  vert?: string;
}
export interface WpsData {
  fill?: Fill;
  outline?: Outline;
  bodyProperties?: WpsBodyProperties & Record<string, unknown>;
  // text-body paragraphs (office-open ParagraphOptions[]) — the text a text-box
  // shape carries; omitted by non-text shapes.
  children?: unknown[];
}

interface GroupChild {
  // "wps" | "wpg" | a pic media type ("jpg"/"png"/…). office-open tags pic
  // children by their image media type rather than a literal "pic".
  type?: string;
  // wps → WpsData; pic → raw image bytes (Uint8Array); wpg → none (uses children).
  data?: unknown;
  transformation?: ChildTransformation;
  children?: GroupChild[];
  childOffset?: PointEmu;
  childExtent?: { cx: number; cy: number };
}

interface WpgGroup {
  children?: GroupChild[];
  transformation?: { width?: number; height?: number } & ChildTransformation;
  childOffset?: PointEmu;
  childExtent?: { cx: number; cy: number };
  // office-open Floating anchor (wp:anchor) — verbatim; rendered via
  // floatingToStyles so an anchored group floats over text instead of flowing.
  floating?: unknown;
}

type Spec = ReadonlyArray<unknown>;

/** Solid fill → CSS color (noFill/none/gradient → undefined). */
function fillToCss(fill: Fill | undefined): string | undefined {
  if (!fill || fill.type !== "solid") return undefined;
  return normalizeColorToHex(fill.color?.value);
}

/** Shape outline → CSS border (EMU width → pt). noFill → undefined. */
function outlineToCss(outline: Outline | undefined): string | undefined {
  if (!outline || outline.type === "noFill" || outline.type === "none") return undefined;
  const color = normalizeColorToHex(outline.color?.value) ?? "black";
  const width = outline.width != null ? `${outline.width / EMU_PER_PT}pt` : "0.75pt";
  const style = outline.dash === "sysDot" || outline.dash === "sysDash" ? "dashed" : "solid";
  return `border:${width} ${style} ${color}`;
}

/** Group extent in px: top-level groups carry {width,height}; nested wpg children
 *  carry a ChildTransformation whose pixels/emus give the size. */
function groupExtent(group: WpgGroup): { w: number; h: number } {
  const t = group.transformation;
  return {
    w: t?.width ?? t?.pixels?.x ?? 0,
    h: t?.height ?? t?.pixels?.y ?? 0,
  };
}

/** Child coord (EMU, group's chOff/chExt space) → group-local px. */
function childBox(
  child: GroupChild,
  chOff: PointEmu,
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

function renderChild(child: GroupChild, chOff: PointEmu, scaleX: number, scaleY: number): Spec {
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
    const data = (child.data ?? {}) as WpsData;
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

interface WpsRunShape {
  text?: string;
  [key: string]: unknown;
}
interface WpsParaShape {
  alignment?: string | null;
  run?: Record<string, unknown>;
  children?: Array<WpsRunShape | string>;
}

/** wps text-body paragraphs (office-open ParagraphOptions[]) → inline HTML runs
 *  so a text-box shape shows its text. Each paragraph's run defaults merge with
 *  each text run; renderRunStyles converts font/color/size/… to CSS. Advanced run
 *  props (shadow w14RawXml, kern) are carried in attrs but not rendered. */
export function renderWpsText(children: unknown): Spec[] {
  if (!Array.isArray(children)) return [];
  return (children as WpsParaShape[]).map((para): Spec => {
    if (typeof para === "string") return ["p", { style: "margin:0" }, para];
    const defaultRun = para.run ?? {};
    const runs = (para.children ?? []).map((r): Spec | string => {
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

/**
 * Render a wps shape's interior (fill/outline/insets/rotation/writing-mode + text
 * body) as a positioned div. Shared by the wpg group's inline wps children and
 * the standalone wpsShape node, so a text-box shape renders identically whether
 * it floats alone (wp:anchor > wps:wsp) or sits inside a group. `positionStyle`
 * carries the placement — absolute group coords for a group child, the floating
 * anchor CSS for a standalone shape. `opts.rotation` overrides bodyPr rotation
 * (a group child's rotation may come from the group transform, not bodyPr).
 */
export function renderWpsInterior(
  data: WpsData,
  positionStyle: string,
  opts?: { rotation?: number; attrs?: Record<string, string> },
): Spec {
  const styles = [positionStyle];
  const fill = fillToCss(data.fill);
  if (fill) styles.push(`background-color:${fill}`);
  const outline = outlineToCss(data.outline);
  if (outline) styles.push(outline);
  // textbox insets (bodyProperties lIns/tIns/rIns/bIns, EMU → px) keep the text
  // inside the shape's padding, matching Word's text-box insets.
  const bp = data.bodyProperties;
  if (bp) {
    const ins = (v: number | undefined) => (v != null ? `${(v / EMU_PER_PX).toFixed(1)}px` : "0px");
    styles.push(`padding:${ins(bp.tIns)} ${ins(bp.rIns)} ${ins(bp.bIns)} ${ins(bp.lIns)}`);
  }
  const rotation = opts?.rotation ?? bp?.rotation;
  if (rotation) styles.push(`transform:rotate(${rotation}deg)`);
  // text direction (bodyPr vert): vert/eaVert/mongolianVert → vertical-rl,
  // vert270 → vertical-lr, so CJK vertical text boxes render top-to-bottom.
  const vert = bp?.vert;
  if (vert && vert !== "horz") {
    styles.push(`writing-mode:${vert === "vert270" ? "vertical-lr" : "vertical-rl"}`);
  }
  const attrs: Record<string, unknown> = { style: styles.join(";") };
  if (opts?.attrs) Object.assign(attrs, opts.attrs);
  return ["div", attrs, ...renderWpsText(data.children)];
}

/** Render a wpg group (top-level or nested) as a positioned container.
 *  actualW/H are the group's real pixel size — top-level groups read them from
 *  transformation.width/height; nested groups receive the box their parent
 *  transform mapped them to. containerStyle (when set) carries the absolute
 *  placement for a nested group; otherwise the group is an inline-block. */
function renderGroup(
  group: WpgGroup,
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

  const attrs: Record<string, unknown> = { "data-wpg-group": "", style };
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
    const wpg = (node.attrs.wpgGroup as WpgGroup | null) ?? {};
    const { w, h } = groupExtent(wpg);
    // A floating group (wp:anchor) overlays its anchor paragraph instead of
    // claiming a line in the flow — render it with the same anchor CSS as
    // images (position:absolute for wrapNone "in front/behind text"), so the
    // group floats over the text below instead of pushing it down.
    if (wpg.floating) {
      const containerStyle = [
        ...floatingToStyles(wpg.floating, undefined, w),
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
});
