import { HANDLES, resizeBox, rotateDelta, type Box, type HandleId } from "./geometry";

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
 *  doc) — the frame snaps back on the next show/refresh. `applyOffset` moves
 *  the drawing by a drag delta (a floating drawing's move); `applyRotation`
 *  spins it by a handle-swept delta (degrees, clockwise); absent, the frame
 *  stays put on a body drag. */
export interface DrawingOverlayCallbacks {
  scale(): number;
  applyBox(box: Box): void;
  applyOffset?(dx: number, dy: number): void;
  applyRotation?(delta: number): void;
}

export class DrawingOverlay {
  readonly el: HTMLDivElement;
  #callbacks: DrawingOverlayCallbacks;
  #box: Box | null = null;
  #rotation = 0;
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
    // Word's rotate handle: a round grip on a stem above the frame's top
    // edge, spinning the drawing about its center.
    const rot = document.createElement("div");
    rot.dataset.handle = "rot";
    Object.assign(rot.style, {
      position: "absolute",
      left: "calc(50% - 5px)",
      top: "-28px",
      width: "10px",
      height: "10px",
      boxSizing: "border-box",
      borderRadius: "50%",
      background: "#fff",
      border: "1.5px solid #2b7cd3",
      pointerEvents: "auto",
      cursor: "grab",
    } satisfies Partial<CSSStyleDeclaration>);
    rot.addEventListener("pointerdown", (e) => this.#onRotateDown(e));
    this.el.append(rot);
    this.el.addEventListener("pointermove", (e) => this.#onPointerMove(e));
    for (const kind of ["pointerup", "pointercancel"] as const)
      this.el.addEventListener(kind, () => this.#onDragEnd());
  }

  /** Show the frame over `box` (page-local px at scale 1), tilted by
   *  `rotation` degrees when the drawing is rotated. */
  show(box: Box, rotation = 0): void {
    this.#box = box;
    this.#rotation = rotation;
    this.el.style.display = "block";
    this.#place();
  }

  /** Refresh after the box changed underneath (a re-render, an applied drag). */
  refresh(box: Box | null, rotation = 0): void {
    if (!box) return this.hide();
    this.show(box, rotation);
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
    // The frame tilts with the drawing (Word's rotated selection): a CSS
    // rotate about the frame's center matches the painter's pivot, and the
    // handles ride along.
    this.el.style.transform = `rotate(${this.#rotation}deg)`;
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

  /** Start a rotate drag from the rotate handle: each pointer move adds its
   *  swept angle around the frame center (screen px, so zoom-independent)
   *  to the spin, and release commits the accumulated degrees. */
  #onRotateDown(event: PointerEvent): void {
    if (!this.#box) return;
    event.preventDefault();
    event.stopPropagation();
    const rect = this.el.getBoundingClientRect();
    const cx = rect.left + rect.width / 2;
    const cy = rect.top + rect.height / 2;
    const origin = this.#rotation;
    let prevX = event.clientX;
    let prevY = event.clientY;
    let total = 0;
    const onMove = (e: PointerEvent): void => {
      total += rotateDelta(cx, cy, prevX, prevY, e.clientX, e.clientY);
      prevX = e.clientX;
      prevY = e.clientY;
      this.#rotation = origin + total;
      this.#place();
    };
    const onUp = (): void => {
      document.removeEventListener("pointermove", onMove);
      document.removeEventListener("pointerup", onUp);
      document.removeEventListener("pointercancel", onUp);
      const delta = Math.round(total);
      if (delta !== 0) this.#callbacks.applyRotation?.(delta);
    };
    document.addEventListener("pointermove", onMove);
    document.addEventListener("pointerup", onUp);
    document.addEventListener("pointercancel", onUp);
  }

  /** True while a frame is shown (a selected drawing on screen). */
  get active(): boolean {
    return this.#box != null;
  }

  /** Begin a move drag from a pointerdown on the selected drawing itself
   *  (the bridge owns that hit chain). The frame trails the pointer via
   *  document-level listeners — the gesture starts on the canvas, not this
   *  element — and a release past {@link MOVE_THRESHOLD} commits as an
   *  offset; a plain click (no real move) leaves everything untouched, so
   *  clicking the selection keeps it (Word). */
  beginMove(clientX: number, clientY: number): void {
    if (!this.#box) return;
    const startX = clientX;
    const startY = clientY;
    const origin = { ...this.#box };
    let moved = false;
    let last: Box = origin;
    const scale = (): number => this.#callbacks.scale() || 1;
    const onMove = (event: PointerEvent): void => {
      const dx = (event.clientX - startX) / scale();
      const dy = (event.clientY - startY) / scale();
      if (!moved && Math.hypot(event.clientX - startX, event.clientY - startY) < MOVE_THRESHOLD)
        return;
      moved = true;
      last = { ...origin, x: origin.x + dx, y: origin.y + dy };
      this.#box = last;
      this.#place();
    };
    const onUp = (): void => {
      document.removeEventListener("pointermove", onMove);
      document.removeEventListener("pointerup", onUp);
      document.removeEventListener("pointercancel", onUp);
      if (moved) this.#callbacks.applyOffset?.(last.x - origin.x, last.y - origin.y);
    };
    document.addEventListener("pointermove", onMove);
    document.addEventListener("pointerup", onUp);
    document.addEventListener("pointercancel", onUp);
  }
}

/** Pointer travel (screen px) that separates a move drag from a plain click
 *  on the selection. */
const MOVE_THRESHOLD = 3;

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
