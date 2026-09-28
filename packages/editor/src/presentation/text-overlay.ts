import type { LayoutDrawingFill } from "@docen/layout";
import type { ProjectedSlideBackground } from "@docen/pptx";

/** A projected paint translated to the textarea's background properties.
 *  The edit surface is opaque, so it needs the same visual source the canvas
 *  painted instead of falling back to the stylesheet's white box. */
export function textBoxFillStyle(
  fill: LayoutDrawingFill | undefined,
  opacity?: number,
  base = "#FFFFFF",
): Partial<CSSStyleDeclaration> {
  if (typeof fill === "string") return { backgroundColor: opaqueColor(fill, opacity, base) };
  if (!fill) return { backgroundColor: base };

  if (fill.type === "linear") {
    const stops = cssStopsOf(fill.stops);
    const angle = (Math.atan2(fill.to.x - fill.from.x, -(fill.to.y - fill.from.y)) * 180) / Math.PI;
    return {
      backgroundColor: base,
      backgroundImage: `linear-gradient(${angle.toFixed(4)}deg, ${stops})`,
    };
  }
  if (fill.type === "radial") {
    const stops = cssStopsOf(fill.stops);
    const radius = Math.hypot(fill.to.x - fill.from.x, fill.to.y - fill.from.y);
    return {
      backgroundColor: base,
      backgroundImage: `radial-gradient(circle ${radius.toFixed(4)}px at ${fill.from.x.toFixed(4)}px ${fill.from.y.toFixed(4)}px, ${stops})`,
    };
  }

  const scale = fill.scale ?? { x: 1, y: 1 };
  // The painter stretches a plain image over the member box; the textarea is
  // that box, so 100% 100% is the faithful translation (intrinsic-size auto
  // would repaint the canvas's stretch as an untiled natural-size image).
  if (fill.mode !== "repeat") {
    return {
      backgroundColor: base,
      backgroundImage: `url("${fill.url.replaceAll('"', "%22")}")`,
      backgroundRepeat: "no-repeat",
      backgroundSize: "100% 100%",
    };
  }
  const at = cssAlignmentOf(fill.align);
  return {
    backgroundColor: base,
    backgroundImage: `url("${fill.url.replaceAll('"', "%22")}")`,
    backgroundPosition:
      fill.offset?.x || fill.offset?.y
        ? `${at.x} ${fill.offset?.x ?? 0}px ${at.y} ${fill.offset?.y ?? 0}px`
        : `${at.x} ${at.y}`,
    backgroundRepeat: fill.mode === "repeat" ? "repeat" : "no-repeat",
    backgroundSize: `${scale.x === 1 ? "auto" : `${scale.x * 100}%`} ${scale.y === 1 ? "auto" : `${scale.y * 100}%`}`,
  };
}

function cssColor(paint: string, opacity?: number): string {
  if (paint.startsWith("rgb")) return paint;
  const hex = paint.replace("#", "");
  if (!/^[0-9a-f]{6}$/i.test(hex)) return paint;
  const raw = Number.parseInt(hex, 16);
  const alpha = opacity ?? 1;
  return alpha >= 1
    ? `#${hex.toUpperCase()}`
    : `rgba(${raw >> 16}, ${(raw >> 8) & 255}, ${raw & 255}, ${alpha})`;
}

const textMirrors = new WeakMap<HTMLTextAreaElement, HTMLDivElement>();

/** Measure a textarea's text stack without the control's row/intrinsic-height
 *  feedback. Reading a textarea's scrollHeight after it has had a large
 *  height can change between keystrokes; the mirror uses the same text styles
 *  and content width, so vertical anchoring stays stable while deleting. */
export function measureTextAreaContent(editor: HTMLTextAreaElement): number {
  let mirror = textMirrors.get(editor);
  if (!mirror) {
    mirror = document.createElement("div");
    mirror.setAttribute("aria-hidden", "true");
    Object.assign(mirror.style, {
      position: "fixed",
      left: "-999999px",
      top: "0",
      visibility: "hidden",
      pointerEvents: "none",
      boxSizing: "content-box",
      whiteSpace: "pre-wrap",
    });
    editor.before(mirror);
    textMirrors.set(editor, mirror);
  }

  const style = getComputedStyle(editor);
  const rect = editor.getBoundingClientRect();
  const horizontalPadding =
    Number.parseFloat(style.paddingLeft) + Number.parseFloat(style.paddingRight);
  Object.assign(mirror.style, {
    width: `${Math.max(0, rect.width - horizontalPadding)}px`,
    fontFamily: style.fontFamily,
    fontSize: style.fontSize,
    fontWeight: style.fontWeight,
    fontStyle: style.fontStyle,
    lineHeight: style.lineHeight,
    letterSpacing: style.letterSpacing,
    textAlign: style.textAlign,
    textTransform: style.textTransform,
    textIndent: style.textIndent,
    overflowWrap: style.overflowWrap,
    wordBreak: style.wordBreak,
    tabSize: style.tabSize,
    direction: style.direction,
  });
  mirror.textContent = editor.value;
  return mirror.getBoundingClientRect().height;
}

export function disposeTextAreaMirror(editor: HTMLTextAreaElement | null): void {
  const mirror = editor && textMirrors.get(editor);
  mirror?.remove();
  if (editor) textMirrors.delete(editor);
}

/** The edit surface's viewport onto the slide: its box origin in slide px
 *  plus the zoom scale turning slide px into CSS px. */
export interface SlideFrame {
  x: number;
  y: number;
  scale: number;
}

/** A shape without its own fill is transparent over the slide; the opaque
 *  edit surface has to inherit that slide paint rather than assuming white.
 *  The paint anchors to the slide, not the textarea — CSS gradients and
 *  backgrounds resolve against the element box, so the slide-space geometry
 *  is re-expressed as the slice this box covers (frame), in element px. */
export function slideFillStyle(
  background: ProjectedSlideBackground | undefined,
  width: number,
  height: number,
  frame?: SlideFrame,
): Partial<CSSStyleDeclaration> {
  if (!background) return { backgroundColor: "#FFFFFF" };
  if (background.kind === "solid") return { backgroundColor: cssColor(background.color) };
  const s = frame?.scale ?? 1;
  const originX = (frame?.x ?? 0) * s;
  const originY = (frame?.y ?? 0) * s;
  if (background.kind === "image") {
    const tile = background.tile;
    if (!tile) {
      // The canvas stretches a plain picture fill over the whole slide; the
      // textarea shows exactly the slice its box covers.
      return {
        backgroundColor: "#FFFFFF",
        backgroundImage: `url("${background.src.replaceAll('"', "%22")}")`,
        backgroundPosition: `${-originX}px ${-originY}px`,
        backgroundRepeat: "no-repeat",
        backgroundSize: `${width * s}px ${height * s}px`,
      };
    }
    return {
      backgroundColor: "#FFFFFF",
      backgroundImage: `url("${background.src.replaceAll('"', "%22")}")`,
      backgroundPosition: `${(tile.offset?.x ?? 0) * s - originX}px ${(tile.offset?.y ?? 0) * s - originY}px`,
      backgroundRepeat: "repeat",
      backgroundSize: `${(tile.scale?.x ?? 1) * 100}% ${(tile.scale?.y ?? 1) * 100}%`,
    };
  }

  const stopsPx = (lengthPx: number): string =>
    background.stops
      .map((stop) => `${cssColor(stop.color)} ${((stop.position / 100) * lengthPx).toFixed(2)}px`)
      .join(", ");
  if (background.path) {
    // Radial: the canvas paints center → bottom-edge radius; re-center that
    // circle on the box's view of the slide.
    const radius = (height / 2) * s;
    const centerX = (width / 2) * s - originX;
    const centerY = (height / 2) * s - originY;
    return {
      backgroundColor: "#FFFFFF",
      backgroundImage: `radial-gradient(circle ${radius.toFixed(2)}px at ${centerX.toFixed(2)}px ${centerY.toFixed(2)}px, ${stopsPx(radius)})`,
    };
  }
  // Linear: project the box's top-left onto the slide's gradient line and
  // hang px stops off that offset — CSS px stops extrapolate past the
  // gradient line, so the box shows exactly its slice of the slide's fade.
  const angle = (((90 + (background.angle ?? 0)) % 360) + 360) % 360;
  const theta = ((background.angle ?? 0) * Math.PI) / 180;
  const length = width * Math.abs(Math.cos(theta)) + height * Math.abs(Math.sin(theta));
  const startX = width / 2 - (Math.cos(theta) * length) / 2;
  const startY = height / 2 - (Math.sin(theta) * length) / 2;
  const boxOffset =
    ((frame?.x ?? 0) - startX) * Math.cos(theta) + ((frame?.y ?? 0) - startY) * Math.sin(theta);
  const stops = background.stops
    .map(
      (stop) =>
        `${cssColor(stop.color)} ${(((stop.position / 100) * length - boxOffset) * s).toFixed(2)}px`,
    )
    .join(", ");
  return {
    backgroundColor: "#FFFFFF",
    backgroundImage: `linear-gradient(${angle.toFixed(4)}deg, ${stops})`,
  };
}

function opaqueColor(paint: string, opacity: number | undefined, base: string): string {
  const color = cssColor(paint, opacity);
  if (!color.startsWith("rgba(")) return color;
  const [red, green, blue, alpha] = color
    .slice(5, -1)
    .split(",")
    .map((part) => Number.parseFloat(part));
  const [baseRed, baseGreen, baseBlue] = base
    .replace("#", "")
    .match(/.{2}/g)!
    .map((hex) => Number.parseInt(hex, 16));
  const blend = (foreground: number, background: number) =>
    Math.round(foreground * alpha + background * (1 - alpha));
  return `#${[blend(red, baseRed), blend(green, baseGreen), blend(blue, baseBlue)]
    .map((value) => value.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase()}`;
}

function cssStopsOf(stops: readonly { offset: number; color: string }[]): string {
  return stops
    .map((stop) => `${cssColor(stop.color)} ${(stop.offset * 100).toFixed(2)}%`)
    .join(", ");
}

function cssAlignmentOf(
  align:
    | "top-left"
    | "top"
    | "top-right"
    | "left"
    | "center"
    | "right"
    | "bottom-left"
    | "bottom"
    | "bottom-right"
    | undefined,
) {
  const parts = new Set((align ?? "").split("-"));
  return {
    x: parts.has("left") ? "left" : parts.has("right") ? "right" : "center",
    y: parts.has("top") ? "top" : parts.has("bottom") ? "bottom" : "center",
  };
}
