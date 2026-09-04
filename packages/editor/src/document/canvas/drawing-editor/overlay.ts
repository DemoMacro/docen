import { HANDLES, resizeBox, type Box, type HandleId } from "./geometry";

/**
 * The drawing selection frame — Word's picture selection: a border plus eight
 * resize handles, with the pointer gestures that drag them. Format-independent:
 * it works on a plain box and reports the resized box through {@link
 * DrawingOverlayCallbacks}; the host adapter owns what a box change means.
 *
 * The overlay is a DOM layer over the canvas (like the caret/selection
 * overlays) — synthetic pointer events reach it directly, where Leafer's
 * scene graph would swallow them.
 */

/** Host-provided I/O: read the scale (page px → screen px), and apply a box
 *  the user dragged to. `applyBox` returns false to reject (e.g. a read-only
 *  doc) — the frame snaps back on the next show/refresh. */
export interface DrawingOverlayCallbacks {
  scale(): number;
  applyBox(box: Box): void;
}

export class DrawingOverlay {
  readonly el: HTMLDivElement;
  #callbacks: DrawingOverlayCallbacks;
  #box: Box | null = null;
  #drag: { handle: HandleId; startX: number; startY: number; origin: Box } | null = null;
  #handles = new Map<HandleId, HTMLDivElement>();

  constructor(callbacks: DrawingOverlayCallbacks) {
    this.#callbacks = callbacks;
    this.el = document.createElement("div");
    // The bridge's takeFocus treats any click inside a `data-docen-overlay`
    // widget as the widget's own — a handle grab must not drop a caret.
    this.el.setAttribute("data-docen-overlay", "");
    Object.assign(this.el.style, {
      position: "absolute",
      pointerEvents: "none",
      zIndex: "6",
      display: "none",
    } satisfies Partial<CSSStyleDeclaration>);
    const frame = document.createElement("div");
    Object.assign(frame.style, {
      position: "absolute",
      inset: "0",
      border: "1.5px solid #2b7cd3",
    } satisfies Partial<CSSStyleDeclaration>);
    this.el.append(frame);
    for (const id of HANDLES) {
      const h = document.createElement("div");
      h.dataset.handle = id;
      Object.assign(h.style, {
        position: "absolute",
        width: "8px",
        height: "8px",
        boxSizing: "border-box",
        background: "#fff",
        border: "1.5px solid #2b7cd3",
        pointerEvents: "auto",
        cursor: HANDLE_CURSORS[id],
      } satisfies Partial<CSSStyleDeclaration>);
      // Word's corner handles sit ON the frame corners; edges on their midpoints.
      const at = HANDLE_ANCHORS[id];
      h.style.left = at.x;
      h.style.top = at.y;
      this.#handles.set(id, h);
      this.el.append(h);
      h.addEventListener("pointerdown", (e) => this.#onHandleDown(id, e));
    }
    this.el.addEventListener("pointermove", (e) => this.#onPointerMove(e));
    for (const kind of ["pointerup", "pointercancel"] as const)
      this.el.addEventListener(kind, () => this.#onDragEnd());
  }

  /** Show the frame over `box` (page-local px at scale 1). */
  show(box: Box): void {
    this.#box = box;
    this.el.style.display = "block";
    this.#place();
  }

  /** Refresh after the box changed underneath (a re-render, an applied drag). */
  refresh(box: Box | null): void {
    if (!box) return this.hide();
    this.show(box);
  }

  hide(): void {
    this.#box = null;
    this.#drag = null;
    this.el.style.display = "none";
  }

  #place(): void {
    const box = this.#box;
    if (!box) return;
    const scale = this.#callbacks.scale();
    this.el.style.left = `${box.x * scale}px`;
    this.el.style.top = `${box.y * scale}px`;
    this.el.style.width = `${box.width * scale}px`;
    this.el.style.height = `${box.height * scale}px`;
  }

  #onHandleDown(handle: HandleId, event: PointerEvent): void {
    if (!this.#box) return;
    event.preventDefault();
    event.stopPropagation();
    // Capture the pointer on the handle so the drag keeps receiving moves
    // after the cursor leaves the handle (and the frame) — without it a
    // corner drag past the frame edge stalls the resize mid-gesture.
    (event.currentTarget as HTMLElement).setPointerCapture(event.pointerId);
    this.#drag = {
      handle,
      startX: event.clientX,
      startY: event.clientY,
      origin: { ...this.#box },
    };
  }

  #onPointerMove(event: PointerEvent): void {
    const drag = this.#drag;
    if (!drag) return;
    const scale = this.#callbacks.scale();
    const dx = (event.clientX - drag.startX) / scale;
    const dy = (event.clientY - drag.startY) / scale;
    const next = resizeBox(drag.origin, drag.handle, dx, dy);
    this.#box = next;
    this.#place();
  }

  #onDragEnd(): void {
    const drag = this.#drag;
    if (!drag || !this.#box) return;
    this.#drag = null;
    this.#callbacks.applyBox(this.#box);
  }
}

const HANDLE_CURSORS: Record<HandleId, string> = {
  nw: "nwse-resize",
  se: "nwse-resize",
  ne: "nesw-resize",
  sw: "nesw-resize",
  n: "ns-resize",
  s: "ns-resize",
  e: "ew-resize",
  w: "ew-resize",
};

const HANDLE_ANCHORS: Record<HandleId, { x: string; y: string }> = {
  nw: { x: "-4px", y: "-4px" },
  n: { x: "calc(50% - 4px)", y: "-4px" },
  ne: { x: "calc(100% - 4px)", y: "-4px" },
  e: { x: "calc(100% - 4px)", y: "calc(50% - 4px)" },
  se: { x: "calc(100% - 4px)", y: "calc(100% - 4px)" },
  s: { x: "calc(50% - 4px)", y: "calc(100% - 4px)" },
  sw: { x: "-4px", y: "calc(100% - 4px)" },
  w: { x: "-4px", y: "calc(50% - 4px)" },
};
