import { presetShapePaths } from "@docen/core/geometry";

export interface ShapeDrawFrame {
  slide: number;
  left: number;
  top: number;
  width: number;
  height: number;
  scale: number;
}

export interface ShapeDrawRect {
  slide: number;
  x: number;
  y: number;
  w: number;
  h: number;
  flipH?: boolean;
  flipV?: boolean;
}

export interface ShapeDrawerHost {
  frameAt(clientX: number, clientY: number): ShapeDrawFrame | null;
  apply(preset: string, rect: ShapeDrawRect): void;
}

const DEFAULT_WIDTH_PX = 192;
const DEFAULT_HEIGHT_PX = 192;
const MOVE_THRESHOLD_PX = 3;
const STRAIGHT_PRESETS = new Set(["line", "straightConnector1"]);

interface DrawState {
  frame: ShapeDrawFrame;
  sx: number;
  sy: number;
  ax: number;
  ay: number;
  moved: boolean;
}

export class ShapeDrawer {
  readonly el: HTMLDivElement;
  #host: ShapeDrawerHost;
  #surface: HTMLElement;
  #preset: string | null = null;
  #draw: DrawState | null = null;
  #rect: HTMLDivElement;
  #line: HTMLDivElement;
  #svg: SVGSVGElement;
  #path: SVGPathElement;
  #onPointerMove = (event: PointerEvent): void => this.#move(event);
  #onPointerUp = (event: PointerEvent): void => this.#finish(event);

  constructor(host: ShapeDrawerHost, surface: HTMLElement) {
    this.#host = host;
    this.#surface = surface;
    this.el = document.createElement("div");
    Object.assign(this.el.style, {
      position: "absolute",
      inset: "0",
      pointerEvents: "none",
      zIndex: "30",
      display: "none",
    } satisfies Partial<CSSStyleDeclaration>);
    this.#rect = document.createElement("div");
    Object.assign(this.#rect.style, {
      position: "absolute",
      boxSizing: "border-box",
      border: "1px solid var(--docen-color-primary, #2b579a)",
      background: "rgba(43, 87, 154, 0.06)",
    } satisfies Partial<CSSStyleDeclaration>);
    this.#line = document.createElement("div");
    Object.assign(this.#line.style, {
      position: "absolute",
      height: "0",
      borderTop: "2px solid var(--docen-color-primary, #2b579a)",
      transformOrigin: "0 0",
    } satisfies Partial<CSSStyleDeclaration>);
    this.#svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    Object.assign(this.#svg.style, {
      position: "absolute",
      overflow: "visible",
    } satisfies Partial<CSSStyleDeclaration>);
    this.#path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    Object.assign(this.#path.style, {
      fill: "rgba(43, 87, 154, 0.06)",
      stroke: "var(--docen-color-primary, #2b579a)",
      strokeWidth: "1",
      vectorEffect: "non-scaling-stroke",
    } satisfies Partial<CSSStyleDeclaration>);
    this.#svg.append(this.#path);
    this.el.append(this.#rect, this.#line, this.#svg);
    surface.append(this.el);
  }

  get armed(): string | null {
    return this.#preset;
  }

  arm(preset: string): void {
    this.#preset = preset;
    this.#surface.style.cursor = "crosshair";
  }

  disarm(): void {
    this.#preset = null;
    this.#surface.style.cursor = "";
    this.hide();
  }

  destroy(): void {
    if (this.#draw) {
      this.#draw = null;
      document.removeEventListener("pointermove", this.#onPointerMove);
      document.removeEventListener("pointerup", this.#onPointerUp);
      document.removeEventListener("pointercancel", this.#onPointerUp);
    }
    this.disarm();
    this.el.remove();
  }

  /** Start a drag when the tool is armed. False leaves the surface's normal
   *  selection/gesture path to run. */
  startFromPointer(event: PointerEvent): boolean {
    if (!this.#preset || event.button !== 0) return false;
    const frame = this.#host.frameAt(event.clientX, event.clientY);
    if (!frame) return false;
    this.#draw = {
      frame,
      sx: event.clientX,
      sy: event.clientY,
      ax: (event.clientX - frame.left) / frame.scale,
      ay: (event.clientY - frame.top) / frame.scale,
      moved: false,
    };
    event.preventDefault();
    document.addEventListener("pointermove", this.#onPointerMove);
    document.addEventListener("pointerup", this.#onPointerUp);
    document.addEventListener("pointercancel", this.#onPointerUp);
    return true;
  }

  #move(event: PointerEvent): void {
    const draw = this.#draw;
    if (!draw) return;
    if (
      !draw.moved &&
      Math.hypot(event.clientX - draw.sx, event.clientY - draw.sy) < MOVE_THRESHOLD_PX
    )
      return;
    draw.moved = true;
    const px = this.#pointerX(event, draw);
    const py = this.#pointerY(event, draw);
    if (this.#isLine) this.#showLine(px, py);
    else if (!this.#showPreset(px, py)) this.#showRect(px, py);
  }

  #finish(event: PointerEvent): void {
    const draw = this.#draw;
    if (!draw) return;
    this.#draw = null;
    document.removeEventListener("pointermove", this.#onPointerMove);
    document.removeEventListener("pointerup", this.#onPointerUp);
    document.removeEventListener("pointercancel", this.#onPointerUp);
    this.hide();
    const preset = this.#preset ?? "rect";
    const px = this.#pointerX(event, draw);
    const py = this.#pointerY(event, draw);
    const rect =
      draw.moved && (Math.abs(px - draw.ax) >= 1 || Math.abs(py - draw.ay) >= 1)
        ? {
            slide: draw.frame.slide,
            x: Math.min(draw.ax, px),
            y: Math.min(draw.ay, py),
            w: Math.abs(px - draw.ax),
            h: Math.abs(py - draw.ay),
            ...(this.#isLine && px < draw.ax ? { flipH: true } : {}),
            ...(this.#isLine && py < draw.ay ? { flipV: true } : {}),
          }
        : this.#defaultRect(draw, px, py);
    this.#host.apply(preset, rect);
    this.disarm();
  }

  get #isLine(): boolean {
    return STRAIGHT_PRESETS.has(this.#preset ?? "");
  }

  #pointerX(event: PointerEvent, draw: DrawState): number {
    return Math.min(
      Math.max(draw.ax + (event.clientX - draw.sx) / draw.frame.scale, 0),
      draw.frame.width,
    );
  }

  #pointerY(event: PointerEvent, draw: DrawState): number {
    return Math.min(
      Math.max(draw.ay + (event.clientY - draw.sy) / draw.frame.scale, 0),
      draw.frame.height,
    );
  }

  #defaultRect(draw: DrawState, px: number, py: number): ShapeDrawRect {
    if (this.#isLine) {
      const x = Math.min(draw.ax, Math.max(0, draw.frame.width - DEFAULT_WIDTH_PX));
      return { slide: draw.frame.slide, x, y: py, w: DEFAULT_WIDTH_PX, h: 0 };
    }
    const width = Math.min(DEFAULT_WIDTH_PX, draw.frame.width);
    const height = Math.min(DEFAULT_HEIGHT_PX, draw.frame.height);
    return {
      slide: draw.frame.slide,
      x: Math.min(Math.max(px - width / 2, 0), Math.max(0, draw.frame.width - width)),
      y: Math.min(Math.max(py - height / 2, 0), Math.max(0, draw.frame.height - height)),
      w: width,
      h: height,
    };
  }

  #showRect(px: number, py: number): void {
    const draw = this.#draw;
    if (!draw) return;
    this.#place(
      this.#rect,
      Math.min(draw.ax, px),
      Math.min(draw.ay, py),
      Math.abs(px - draw.ax),
      Math.abs(py - draw.ay),
    );
  }

  #showLine(px: number, py: number): void {
    const draw = this.#draw;
    if (!draw) return;
    const dx = (px - draw.ax) * draw.frame.scale;
    const dy = (py - draw.ay) * draw.frame.scale;
    this.#line.style.left = `${draw.frame.left + draw.ax * draw.frame.scale}px`;
    this.#line.style.top = `${draw.frame.top + draw.ay * draw.frame.scale}px`;
    this.#line.style.width = `${Math.hypot(dx, dy)}px`;
    this.#line.style.transform = `rotate(${(Math.atan2(dy, dx) * 180) / Math.PI}deg)`;
    this.#line.style.display = "block";
  }

  #showPreset(px: number, py: number): boolean {
    const draw = this.#draw;
    if (!draw) return false;
    const x = Math.min(draw.ax, px);
    const y = Math.min(draw.ay, py);
    const w = Math.abs(px - draw.ax);
    const h = Math.abs(py - draw.ay);
    const paths = presetShapePaths(this.#preset ?? "rect", w, h);
    if (!paths?.length) return false;
    this.#svg.style.left = `${draw.frame.left + x * draw.frame.scale}px`;
    this.#svg.style.top = `${draw.frame.top + y * draw.frame.scale}px`;
    this.#svg.style.width = `${w * draw.frame.scale}px`;
    this.#svg.style.height = `${h * draw.frame.scale}px`;
    this.#svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
    this.#path.setAttribute("d", paths.map((part) => part.d).join(" "));
    this.#svg.style.display = "block";
    return true;
  }

  #place(el: HTMLElement, x: number, y: number, w: number, h: number): void {
    const draw = this.#draw;
    if (!draw) return;
    el.style.left = `${draw.frame.left + x * draw.frame.scale}px`;
    el.style.top = `${draw.frame.top + y * draw.frame.scale}px`;
    el.style.width = `${w * draw.frame.scale}px`;
    el.style.height = `${h * draw.frame.scale}px`;
    el.style.display = "block";
  }

  hide(): void {
    this.#rect.style.display = "none";
    this.#line.style.display = "none";
    this.#svg.style.display = "none";
  }
}
