import { resizeCrop, type Box, type CropFractions, type HandleId } from "./geometry";

/**
 * The crop-mode layer — Word's picture crop: the source shows in full (the
 * cropped-off bands dimmed) with black crop handles, a handle drag moves the
 * crop edge, and Enter (or a press outside) commits while Esc cancels.
 *
 * Format-independent like the selection overlay: it works on the visible box
 * plus crop fractions and reports the new crop through {@link
 * CropOverlayCallbacks}; the host adapter owns what a crop means. The layer
 * paints its own full-source preview (an <img> the host hands in), so the
 * canvas — which renders the cropped result — needs no crop-mode branch.
 */

/** Host-provided I/O: read the scale (page px → screen px), apply the
 *  cropped fraction set the user committed, and reclaim the layer's needs on
 *  exit (the selection frame it displaces comes back through this). */
export interface CropOverlayCallbacks {
  scale(): number;
  applyCrop(crop: CropFractions): void;
  onExit?(): void;
}

/** The crop handles, clockwise from the top-left corner — the resize
 *  handle ids reused so the drag math keys off the same vocabulary. */
const CROP_HANDLES: readonly HandleId[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];

export class CropOverlay {
  readonly el: HTMLDivElement;
  #callbacks: CropOverlayCallbacks;
  /** The visible (cropped) box — page-local px at scale 1. */
  #box: Box | null = null;
  #rotation = 0;
  #crop: CropFractions | null = null;
  /** The crop when the mode entered — a release matching it commits nothing. */
  #origin: CropFractions | null = null;
  #full: Box | null = null;
  #img = new Image();
  #shades = new Map<HandleId /* axis face */, HTMLDivElement>();
  #handles = new Map<HandleId, HTMLDivElement>();
  #onDocDown = (e: PointerEvent): void => this.#onDocumentPointerDown(e);
  #onDocKey = (e: KeyboardEvent): void => this.#onDocumentKeyDown(e);

  constructor(callbacks: CropOverlayCallbacks) {
    this.#callbacks = callbacks;
    this.el = document.createElement("div");
    // The bridge's takeFocus treats a `data-docen-overlay` widget's press as
    // its own — leaning on the crop layer must not drop a caret.
    this.el.setAttribute("data-docen-overlay", "");
    Object.assign(this.el.style, {
      position: "absolute",
      pointerEvents: "none",
      zIndex: "6",
      display: "none",
      overflow: "visible",
    } satisfies Partial<CSSStyleDeclaration>);
    this.#img.style.cssText = "width:100%;height:100%;display:block;user-select:none";
    this.#img.draggable = false;
    this.el.append(this.#img);
    // The dimmed cropped-off bands: full-width top/bottom, side bands between
    // them (percent-of-source geometry, so a drag re-places by one #place).
    for (const face of ["n", "s", "w", "e"] as const) {
      const shade = document.createElement("div");
      Object.assign(shade.style, {
        position: "absolute",
        background: "rgba(0,0,0,0.5)",
        pointerEvents: "none",
      } satisfies Partial<CSSStyleDeclaration>);
      this.#shades.set(face, shade);
      this.el.append(shade);
    }
    // The visible-frame outline (Word's crop line) on top of the shades.
    const frame = document.createElement("div");
    frame.dataset.cropFrame = "";
    Object.assign(frame.style, {
      position: "absolute",
      border: "1px solid rgba(255,255,255,0.9)",
      boxShadow: "0 0 0 1px rgba(0,0,0,0.35)",
      pointerEvents: "none",
    } satisfies Partial<CSSStyleDeclaration>);
    this.el.append(frame);
    for (const id of CROP_HANDLES) {
      const h = document.createElement("div");
      h.dataset.cropHandle = id;
      Object.assign(h.style, {
        position: "absolute",
        width: "8px",
        height: "8px",
        boxSizing: "border-box",
        background: "#000",
        border: "1px solid #fff",
        pointerEvents: "auto",
        cursor: CROP_CURSORS[id],
      } satisfies Partial<CSSStyleDeclaration>);
      h.addEventListener("pointerdown", (e) => this.#onHandleDown(id, e));
      this.#handles.set(id, h);
      this.el.append(h);
    }
  }

  /** Enter crop mode over `box` (page-local px at scale 1), tilted by
   *  `rotation` degrees; `src` is the full source URL the layer previews. */
  show(box: Box, rotation: number, crop: CropFractions, src: string): void {
    this.#box = box;
    this.#rotation = rotation;
    this.#crop = { ...crop };
    this.#origin = { ...crop };
    this.#img.src = src;
    this.el.style.display = "block";
    this.#place();
    document.addEventListener("pointerdown", this.#onDocDown, true);
    document.addEventListener("keydown", this.#onDocKey, true);
  }

  /** True while crop mode is on. */
  get active(): boolean {
    return this.#box != null;
  }

  /** Commit (a changed crop applies through the host) and leave crop mode. */
  commit(): void {
    if (this.#crop && this.#origin && this.#box) {
      const moved =
        this.#crop.left !== this.#origin.left ||
        this.#crop.top !== this.#origin.top ||
        this.#crop.right !== this.#origin.right ||
        this.#crop.bottom !== this.#origin.bottom;
      if (moved) this.#callbacks.applyCrop({ ...this.#crop });
    }
    this.hide();
  }

  /** Leave crop mode without applying (Esc). */
  cancel(): void {
    this.hide();
  }

  hide(): void {
    if (this.#box == null) return;
    this.#box = null;
    this.#crop = null;
    this.#origin = null;
    this.el.style.display = "none";
    document.removeEventListener("pointerdown", this.#onDocDown, true);
    document.removeEventListener("keydown", this.#onDocKey, true);
    this.#callbacks.onExit?.();
  }

  /** Lay the layer out over the full-source rectangle: everything inside is
   *  percent-of-source positioned, so one outer geometry write re-places the
   *  preview, the shades, the frame, and the handles together. */
  #place(): void {
    const box = this.#box;
    const crop = this.#crop;
    if (!box || !crop) return;
    const w = 1 - crop.left - crop.right;
    const h = 1 - crop.top - crop.bottom;
    const full: Box = {
      x: box.x - box.width * (w > 0 ? crop.left / w : 0),
      y: box.y - box.height * (h > 0 ? crop.top / h : 0),
      width: w > 0 ? box.width / w : box.width,
      height: h > 0 ? box.height / h : box.height,
    };
    this.#full = full;
    const scale = this.#callbacks.scale() || 1;
    this.el.style.left = `${full.x * scale}px`;
    this.el.style.top = `${full.y * scale}px`;
    this.el.style.width = `${full.width * scale}px`;
    this.el.style.height = `${full.height * scale}px`;
    this.el.style.transform = `rotate(${this.#rotation}deg)`;
    const pct = (v: number): string => `${v * 100}%`;
    const placeBand = (
      face: "n" | "s" | "w" | "e",
      left: string,
      top: string,
      width: string,
      height: string,
    ): void => {
      const band = this.#shades.get(face)!;
      band.style.left = left;
      band.style.top = top;
      band.style.width = width;
      band.style.height = height;
    };
    placeBand("n", "0", "0", "100%", pct(crop.top));
    placeBand("s", "0", pct(1 - crop.bottom), "100%", pct(crop.bottom));
    placeBand("w", "0", pct(crop.top), pct(crop.left), pct(h));
    placeBand("e", pct(1 - crop.right), pct(crop.top), pct(crop.right), pct(h));
    const frame = this.el.querySelector<HTMLElement>("[data-crop-frame]")!;
    frame.style.left = pct(crop.left);
    frame.style.top = pct(crop.top);
    frame.style.width = pct(w);
    frame.style.height = pct(h);
    for (const id of CROP_HANDLES) {
      const h2 = this.#handles.get(id)!;
      const at = CROP_ANCHORS[id](crop);
      h2.style.left = at.x;
      h2.style.top = at.y;
    }
  }

  #onHandleDown(handle: HandleId, event: PointerEvent): void {
    if (!this.#crop || !this.#full) return;
    event.preventDefault();
    event.stopPropagation();
    (event.currentTarget as HTMLElement).setPointerCapture(event.pointerId);
    const startX = event.clientX;
    const startY = event.clientY;
    const origin = { ...this.#crop };
    // A screen delta maps to a source fraction through the full box (scaled
    // to screen px), un-rotated so a tilted crop drags along its own axes.
    const scale = (this.#callbacks.scale() || 1) * (this.#full.width || 1);
    const scaleV = (this.#callbacks.scale() || 1) * (this.#full.height || 1);
    const rad = (this.#rotation * Math.PI) / 180;
    const cos = Math.cos(rad);
    const sin = Math.sin(rad);
    const onMove = (e: PointerEvent): void => {
      const sx = e.clientX - startX;
      const sy = e.clientY - startY;
      // Un-rotate: rotate the screen delta by −rotation into layer space.
      const lx = (sx * cos + sy * sin) / scale;
      const ly = (-sx * sin + sy * cos) / scaleV;
      this.#crop = resizeCrop(origin, handle, lx, ly, lx, ly);
      this.#place();
    };
    const onUp = (): void => {
      document.removeEventListener("pointermove", onMove);
      document.removeEventListener("pointerup", onUp);
      document.removeEventListener("pointercancel", onUp);
    };
    document.addEventListener("pointermove", onMove);
    document.addEventListener("pointerup", onUp);
    document.addEventListener("pointercancel", onUp);
  }

  /** A press outside the crop layer commits (Word: clicking away applies the
   *  crop). Capture phase + stopPropagation: the press belongs to the commit
   *  — the canvas chain must not also drop a caret or drop the selection. */
  #onDocumentPointerDown(event: PointerEvent): void {
    // The layer lives inside the editor's shadow tree, but this listener sits
    // on `document` — across that boundary the event target retargets to the
    // host element, so only the composed path still names the real handle.
    if (event.composedPath().includes(this.el)) return;
    event.stopPropagation();
    event.preventDefault();
    this.commit();
  }

  #onDocumentKeyDown(event: KeyboardEvent): void {
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      this.cancel();
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault();
      event.stopPropagation();
      this.commit();
    }
  }
}

const CROP_CURSORS: Record<HandleId, string> = {
  nw: "nwse-resize",
  se: "nwse-resize",
  ne: "nesw-resize",
  sw: "nesw-resize",
  n: "ns-resize",
  s: "ns-resize",
  e: "ew-resize",
  w: "ew-resize",
};

/** Percent-of-source anchor for each crop handle (the visible frame's
 *  corners and edge midpoints), each offset half a handle so the grip's
 *  center sits ON the frame point. */
const CROP_ANCHORS: Record<HandleId, (c: CropFractions) => { x: string; y: string }> = {
  nw: (c) => ({ x: at(c.left), y: at(c.top) }),
  n: (c) => ({ x: "calc(50% - 4px)", y: at(c.top) }),
  ne: (c) => ({ x: at(1 - c.right), y: at(c.top) }),
  e: (c) => ({ x: at(1 - c.right), y: "calc(50% - 4px)" }),
  se: (c) => ({ x: at(1 - c.right), y: at(1 - c.bottom) }),
  s: (c) => ({ x: "calc(50% - 4px)", y: at(1 - c.bottom) }),
  sw: (c) => ({ x: at(c.left), y: at(1 - c.bottom) }),
  w: (c) => ({ x: at(c.left), y: "calc(50% - 4px)" }),
};

const at = (v: number): string => `calc(${v * 100}% - 4px)`;
