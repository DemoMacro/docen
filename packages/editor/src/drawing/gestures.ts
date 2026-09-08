import type { Editor } from "@docen/docx/core";
import { EMU_PER_PX } from "@docen/layout";
import { NodeSelection } from "@tiptap/pm/state";

import { CropOverlay } from "./crop-overlay";
import { DrawingOverlay } from "./overlay";
import type { DrawingHit } from "./target";

/** What the gestures need from their host: the viewless editor (its state and
 *  commands), the PM position a hit resolves to, the re-resolved painted box
 *  after a re-render, the page frames the overlays mount in, and the zoom
 *  factor (semantic page px → screen px). All injected — the gestures own no
 *  document state beyond the selection they manage. */
export interface DrawingGesturesHost {
  editor(): Editor;
  drawingSelection(hit: DrawingHit): number | null;
  drawingBoxOf(
    para: unknown,
    index: number,
    kind: "drawing" | "inline",
    childPath?: readonly number[],
  ): DrawingHit | null;
  pageHost(page: number): HTMLElement | null;
  scale(): number;
}

/**
 * The drawing selection state machine — Word's picture selection plus the
 * gestures that edit it. Owns the selection frame (resize/move/rotate) and
 * the crop layer; every write-back lands through the host's editor commands,
 * so the class stays format-independent (docx today, pptx/xlsx tomorrow).
 */
export class DrawingGestures {
  #host: DrawingGesturesHost;
  #overlay: DrawingOverlay;
  #crop: CropOverlay;
  /** The selected drawing — the hit box carries the laid host paragraph + its
   *  drawing index (how the PM node was found); after a re-render the box
   *  re-resolves from the stage table, and a drawing that no longer paints
   *  drops the selection. */
  #sel: DrawingHit | null = null;

  constructor(host: DrawingGesturesHost) {
    this.#host = host;
    // The drawing-editor adapter: what a dragged box means. The box is
    // page-local px at scale 1; the PM node's width/height attrs are px too,
    // so the write-back is a direct setNodeMarkup (inline pictures size
    // through the same attrs; a re-render re-anchors the frame via
    // drawingBoxOf). A body drag moves the drawing: the px delta converts to
    // EMU (the floating attrs' unit) and lands through the engine's
    // move-drawing command, which adds it to the current offsets.
    this.#overlay = new DrawingOverlay({
      scale: () => this.#host.scale(),
      applyBox: (box) => this.#applyBox(box),
      applyOffset: (dx, dy) => this.#applyOffset(dx, dy),
      applyRotation: (delta) => this.#applyRotation(delta),
    });
    // The crop layer: the selected image's source shows in full with black
    // crop handles; a commit writes the dragged insets through the crop
    // command (fractions — the command converts to the attrs' raw ints). It
    // displaces the selection frame while on — its handles share the frame's
    // grip points, and a resize landing under a crop drag would re-render and
    // kill the mode — so exiting hands the frame back through onExit.
    this.#crop = new CropOverlay({
      scale: () => this.#host.scale(),
      applyCrop: (crop) => {
        if (!this.#sel) return;
        const nodePos = this.#host.drawingSelection(this.#sel);
        if (nodePos == null) return;
        this.#host.editor().commands["drawing-crop-apply"](JSON.parse(JSON.stringify(crop)));
      },
      onExit: () => this.place(),
    });
  }

  /** Mount both overlay layers into the positioned host covering the canvas
   *  surface (they re-parent into per-page frames as selections move). */
  mount(host: HTMLElement): void {
    host.append(this.#overlay.el, this.#crop.el);
  }

  /** The selected drawing's hit, if one is selected. */
  get selected(): DrawingHit | null {
    return this.#sel;
  }

  /** True while the selection frame is shown (a selected drawing on screen). */
  get frameActive(): boolean {
    return this.#overlay.active;
  }

  /** Whether the selection is a floating drawing a pointer drag can move. */
  movableFloating(): boolean {
    return this.#floatingAnchors() != null;
  }

  /** Begin a move drag from a pointerdown on the selected drawing itself (the
   *  bridge owns that hit chain). */
  beginMove(clientX: number, clientY: number): void {
    this.#overlay.beginMove(clientX, clientY);
  }

  /** Select the drawing a click hit: the NodeSelection lands first, then the
   *  frame shows. A member hit that resolved to the group (the group was not
   *  entered — Word's first click selects the whole group) frames the group's
   *  box, not the member's. False when the PM side cannot pair the hit (the
   *  caller falls through to text placement). */
  select(hit: DrawingHit): boolean {
    const nodePos = this.#host.drawingSelection(hit);
    const node = nodePos != null ? this.#host.editor().state.doc.nodeAt(nodePos) : null;
    if (nodePos == null || !node) return false;
    this.#host.editor().commands.command(({ state, dispatch }) => {
      dispatch?.(state.tr.setSelection(NodeSelection.create(state.doc, nodePos) as never));
      return true;
    });
    this.#sel =
      hit.childPath && node.type.name === "wpgGroup"
        ? (this.#host.drawingBoxOf(hit.para, hit.index, hit.kind) ?? hit)
        : hit;
    this.place();
    return true;
  }

  /** Re-place the frame against the fresh geometry (after a re-render or a
   *  zoom change). A selected drawing that no longer paints drops. */
  place(): void {
    if (this.#sel) {
      this.#sel =
        this.#host.drawingBoxOf(
          this.#sel.para,
          this.#sel.index,
          this.#sel.kind,
          this.#sel.childPath,
        ) ?? null;
    }
    const frame = this.#sel ? this.#host.pageHost(this.#sel.page) : null;
    if (!this.#sel || !frame) {
      this.#overlay.hide();
      return;
    }
    if (frame !== this.#overlay.el.parentElement) frame.append(this.#overlay.el);
    this.#overlay.refresh(this.#sel, this.#sel.rotation);
  }

  /** Drop the selection state (the frame itself hides on the next place). */
  clear(): void {
    this.#sel = null;
  }

  /** Enter crop mode on the selected image (the context menu's Crop). False
   *  when the selection frame isn't showing or the node carries no source —
   *  shapes and source-less images have nothing to crop. */
  enterCropMode(): boolean {
    if (!this.#sel || this.#crop.active) return false;
    const nodePos = this.#host.drawingSelection(this.#sel);
    const editor = this.#host.editor();
    const node = nodePos != null ? editor.state.doc.nodeAt(nodePos) : null;
    const attrs = node?.attrs as Record<string, unknown> | undefined;
    const src = typeof attrs?.src === "string" ? attrs.src : null;
    if (!src) return false;
    // The layer positions page-locally, like the selection frame — mount in
    // the drawing's page frame.
    const frame = this.#host.pageHost(this.#sel.page);
    if (!frame) return false;
    if (frame !== this.#crop.el.parentElement) frame.append(this.#crop.el);
    this.#overlay.hide();
    // The attrs carry the raw ST_Percentage ints (100000 = 100%) — the same
    // contract breach cropOf reads through; divide back to fractions here.
    const raw = (attrs?.crop ?? {}) as Record<string, unknown>;
    const fraction = (v: unknown): number =>
      typeof v === "number" && Number.isFinite(v) ? v / 100000 : 0;
    this.#crop.show(
      this.#sel,
      this.#sel.rotation ?? 0,
      {
        left: fraction(raw.left),
        top: fraction(raw.top),
        right: fraction(raw.right),
        bottom: fraction(raw.bottom),
      },
      src,
    );
    return true;
  }

  /** A re-render under an open crop layer orphans its geometry — drop the
   *  mode (the drag commits through Enter/click, never mid-transaction) and
   *  re-place the frame. */
  replaceOverlays(): void {
    if (this.#crop.active) this.#crop.cancel();
    this.place();
  }

  destroy(): void {
    this.#overlay.hide();
    this.#overlay.el.remove();
    this.#crop.el.remove();
  }

  /** The selected floating drawing's position anchors (null on any other
   *  selection) — any float can start a pointer drag; whether the drop adds
   *  a delta or lands absolute depends on the anchor shape. */
  #floatingAnchors(): {
    h: Record<string, unknown> | undefined;
    v: Record<string, unknown> | undefined;
  } | null {
    const sel = this.#host.editor().state.selection;
    if (!(sel instanceof NodeSelection)) return null;
    const attrs = sel.node.attrs as Record<string, unknown>;
    const floating =
      sel.node.type.name === "image"
        ? (attrs.floating as Record<string, unknown> | null)
        : ((attrs.wpsShape as Record<string, unknown> | null | undefined)?.floating as Record<
            string,
            unknown
          > | null);
    if (!floating) return null;
    return {
      h: floating.horizontalPosition as Record<string, unknown> | undefined,
      v: floating.verticalPosition as Record<string, unknown> | undefined,
    };
  }

  /** A handle drag's box: the selected image's px attrs resize in place,
   *  keeping the NodeSelection (a setNodeMarkup that changes attrs demotes a
   *  NodeSelection to a caret — re-create it so the resize keeps the drawing
   *  selected, Word's picture stays selected after a handle drag). */
  #applyBox(box: { x: number; y: number; width: number; height: number }): void {
    if (!this.#sel) return;
    const nodePos = this.#host.drawingSelection(this.#sel);
    if (nodePos == null) return;
    const editor = this.#host.editor();
    if (!editor.state.doc.nodeAt(nodePos)) return;
    editor.commands.command(({ tr, dispatch }) => {
      tr.setNodeMarkup(nodePos, undefined, {
        ...editor.state.doc.nodeAt(nodePos)!.attrs,
        width: box.width,
        height: box.height,
      }).setSelection(NodeSelection.create(tr.doc, nodePos) as never);
      dispatch?.(tr as never);
      return true;
    });
  }

  /** A body drag's offset: an offset-anchored float adds the drag delta to
   *  its offsets; an align-anchored one has none to add to — the drop lands
   *  absolute, page-anchored, at the dragged spot (Word: dragging breaks the
   *  alignment, the drawn result doesn't shift). */
  #applyOffset(dx: number, dy: number): void {
    if (!this.#sel) return;
    const nodePos = this.#host.drawingSelection(this.#sel);
    if (nodePos == null) return;
    const hEmu = Math.round(dx * EMU_PER_PX);
    const vEmu = Math.round(dy * EMU_PER_PX);
    const anchors = this.#floatingAnchors();
    if (typeof anchors?.h?.offset === "number" && typeof anchors.v?.offset === "number") {
      this.#host.editor().commands["move-drawing"](JSON.stringify({ h: hEmu, v: vEmu }));
      return;
    }
    this.#host.editor().commands["place-drawing"](
      JSON.stringify({
        h: Math.round(this.#sel.x * EMU_PER_PX) + hEmu,
        v: Math.round(this.#sel.y * EMU_PER_PX) + vEmu,
      }),
    );
  }

  /** A rotate handle's swept delta (degrees, clockwise). */
  #applyRotation(delta: number): void {
    if (!this.#sel) return;
    const nodePos = this.#host.drawingSelection(this.#sel);
    if (nodePos == null) return;
    this.#host.editor().commands["rotate-drawing"](JSON.stringify(delta));
  }
}
