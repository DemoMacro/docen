import type { Editor } from "@docen/docx/core";
import { EMU_PER_PX } from "@docen/layout";
import type { Node as PmNode } from "@tiptap/pm/model";
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
  /** `enter` marks the entry double click on a group — the member hit
   *  resolves to the member instead of the group (Word: the second click
   *  enters the group). */
  drawingSelection(hit: DrawingHit, enter?: boolean): number | null;
  drawingBoxOf(
    para: unknown,
    index: number,
    kind: "drawing" | "inline",
    childPath?: readonly number[],
  ): DrawingHit | null;
  /** The paragraph under a page-local drop point: its content-end insertion
   *  position and laid box origin — the drop re-anchor's target. Null on
   *  bare geometry (margin, furniture) where no paragraph can host one. */
  paragraphAt(
    page: number,
    x: number,
    y: number,
  ): { contentPos: number; xPx: number; yPx: number } | null;
  /** The page a pointer position sits over, with the point as page-local px
   *  at scale 1 plus the page's own rect (the drop resolution's frame — a
   *  drag may cross pages, and only the host knows the pages' screen
   *  geometry). Null off the pages. */
  pageAtPoint(
    clientX: number,
    clientY: number,
  ): { page: number; x: number; y: number; w: number; h: number } | null;
  /** Whether the drop point resolves to a different table cell than the
   *  selected drawing's anchor paragraph (out of the cell, across cells, or
   *  into one): the cell-anchored box's layoutInCell clamp pins it at the
   *  cell, eating any committed offset — the drop re-homes beside the point
   *  instead (Word re-anchors on drag). Absent: no cell tracking, drops keep
   *  the anchors. */
  crossesCell?(hit: DrawingHit, page: number, x: number, y: number): boolean;
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
  /** The Shift+Click multi-selection beyond the primary: floating drawings
   *  framed lightly, consumed by group/distribute/align (Word's multi-object
   *  selection). A doc change re-validates in place(). */
  #multi: DrawingHit[] = [];
  #multiBoxes: HTMLDivElement[] = [];
  /** A move drag awaiting its landing check: the drop's delta writes against
   *  the pre-drag layout, but re-wrapping shifts the anchor paragraph, so the
   *  painted box lands off the drop point by that shift. place() measures the
   *  painted box after the re-layout and re-commits the residue as another
   *  delta (Word corrects the offset against the final anchor the same way).
   *  Page-local px, the page the drop targeted, whether the anchor was
   *  re-homed, the passes spent, the painted box at the last commit, and the
   *  last committed delta (a commit that paints no movement is being eaten
   *  by a clamp — a cell-anchored box pinned at its cell's edge — and the
   *  delta retreats instead of piling onto a stuck box). */
  #dropCorrection: {
    x: number;
    y: number;
    page: number;
    tries: number;
    reanchored: boolean;
    box?: { x: number; y: number };
    last?: { h: number; v: number };
  } | null = null;
  /** Where the grab point sat inside the drawing (page-local px): Word's
   *  drag keeps that offset — the drawing trails the pointer, so its top
   *  left lands at the release point minus this offset, on either side of
   *  a page crossing. */
  #grab: { x: number; y: number } | null = null;

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
      applyOffset: (dx, dy, clientX, clientY) => this.#applyOffset(dx, dy, clientX, clientY),
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

  /** The hit's PM node position, null when the host cannot pair it (the
   *  caller falls back to its own handling). */
  nodePosOf(hit: DrawingHit): number | null {
    return this.#host.drawingSelection(hit);
  }

  /** Begin a move drag from a pointerdown on the selected drawing itself (the
   *  bridge owns that hit chain). */
  beginMove(clientX: number, clientY: number): void {
    const down = this.#sel ? this.#host.pageAtPoint(clientX, clientY) : null;
    this.#grab = down && this.#sel ? { x: down.x - this.#sel.x, y: down.y - this.#sel.y } : null;
    this.#overlay.beginMove(clientX, clientY);
  }

  /** Select the drawing a click hit: the NodeSelection lands first, then the
   *  frame shows. A member hit that resolved to the group (the group was not
   *  entered — Word's first click selects the whole group) frames the group's
   *  box, not the member's. `enter` (the entry double click) targets the
   *  member instead. False when the PM side cannot pair the hit (the caller
   *  falls through to text placement). */
  select(hit: DrawingHit, enter = false): boolean {
    const nodePos = this.#host.drawingSelection(hit, enter);
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

  /** Word's Shift+Click multi-selection: toggle a floating drawing in/out of
   *  the set beyond the primary (which keeps the frame). With nothing framed
   *  yet the click just selects — the primary starts the set. False when the
   *  hit cannot pair to the PM side or its node is not floating (inline art
   *  rides the text stream); the caller falls through to the plain click. */
  toggleMulti(hit: DrawingHit): boolean {
    const nodePos = this.#host.drawingSelection(hit);
    const node = nodePos != null ? this.#host.editor().state.doc.nodeAt(nodePos) : null;
    if (nodePos == null || !node) return false;
    if (!this.#floatingOf(node)) return false;
    if (!this.#sel) return this.select(hit);
    const i = this.#multi.findIndex((m) => sameHit(m, hit));
    if (i >= 0) this.#multi.splice(i, 1);
    else this.#multi.push(hit);
    this.place();
    return true;
  }

  /** The multi-selection's members for the group/distribute payloads: the
   *  primary plus every toggled member, each with its PM position and page
   *  box. Null when fewer than two resolve (nothing to group or distribute). */
  multiPayload():
    | { pos: number; box: { x: number; y: number; width: number; height: number } }[]
    | null {
    const all = [this.#sel, ...this.#multi].filter((hit): hit is DrawingHit => hit != null);
    const members: { pos: number; box: { x: number; y: number; width: number; height: number } }[] =
      [];
    for (const hit of all) {
      const pos = this.#host.drawingSelection(hit);
      if (pos == null) continue;
      members.push({ pos, box: { x: hit.x, y: hit.y, width: hit.width, height: hit.height } });
    }
    return members.length >= 2 ? members : null;
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
    if (this.#dropCorrection) this.#correctDrop();
    const frame = this.#sel ? this.#host.pageHost(this.#sel.page) : null;
    if (!this.#sel || !frame) {
      this.#overlay.hide();
    } else {
      if (frame !== this.#overlay.el.parentElement) frame.append(this.#overlay.el);
      this.#overlay.refresh(this.#sel, this.#sel.rotation);
    }
    this.#placeMulti();
  }

  /** Re-resolve the multi-selection against the fresh geometry and redraw the
   *  light frames (a member that no longer paints, or one that became the
   *  primary, drops out). */
  #placeMulti(): void {
    this.#multi = this.#multi
      .map((hit) => this.#host.drawingBoxOf(hit.para, hit.index, hit.kind, hit.childPath))
      .filter((hit): hit is DrawingHit => hit != null && !(this.#sel && sameHit(hit, this.#sel)));
    for (const el of this.#multiBoxes) el.remove();
    this.#multiBoxes = [];
    if (!this.#multi.length) return;
    const scale = this.#host.scale();
    for (const hit of this.#multi) {
      const host = this.#host.pageHost(hit.page);
      if (!host) continue;
      const el = document.createElement("div");
      Object.assign(el.style, {
        position: "absolute",
        left: `${hit.x * scale}px`,
        top: `${hit.y * scale}px`,
        width: `${hit.width * scale}px`,
        height: `${hit.height * scale}px`,
        border: "1px solid #2b7cd3",
        pointerEvents: "none",
        zIndex: "6",
      } satisfies Partial<CSSStyleDeclaration>);
      host.append(el);
      this.#multiBoxes.push(el);
    }
  }

  /** Drop the selection state (the frame itself hides on the next place). */
  clear(): void {
    this.#sel = null;
    this.#multi = [];
    for (const el of this.#multiBoxes) el.remove();
    this.#multiBoxes = [];
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
    this.clear();
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
    const floating = this.#floatingOf(sel.node);
    if (!floating) return null;
    return {
      h: floating.horizontalPosition as Record<string, unknown> | undefined,
      v: floating.verticalPosition as Record<string, unknown> | undefined,
    };
  }

  /** A node's floating attrs carrier (null on any non-floating node) — the
   *  position anchors a pointer drag moves. */
  #floatingOf(node: PmNode): Record<string, unknown> | null {
    const attrs = node.attrs as Record<string, unknown>;
    const carrier =
      node.type.name === "image"
        ? attrs.floating
        : node.type.name === "wpsShape"
          ? (attrs.wpsShape as Record<string, unknown> | undefined)?.floating
          : node.type.name === "wpgGroup"
            ? (attrs.wpgGroup as Record<string, unknown> | undefined)?.floating
            : null;
    return (carrier as Record<string, unknown> | null | undefined) ?? null;
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

  /** A body drag's offset: the release point resolves the drop page first —
   *  a drag may cross pages, and the raw delta is a coordinate from the
   *  origin page's frame, so committing it against the old anchor pushes the
   *  drawing past its page top where nothing paints it (it "vanishes").
   *  Every drop quantity is the release point minus the grab offset (the
   *  drawing trails the pointer — Word), which on the same page is algebra
   *  equal to the plain delta. The drop box clamps into the page rect first
   *  — Word's drag keeps the drawing on its page, and a commit whose box
   *  paints off-page can never be corrected (there is no painted box left
   *  to measure), so it must not land. A cross-page drop, or one landing in
   *  a different table cell than the anchor, re-homes beside the drop point
   *  instead (Word re-anchors on drag; the cell clamp eats offsets).
   *  Otherwise the drop keeps the anchors: an offset-anchored float takes
   *  the bounded delta toward the clamped target, an align-anchored one
   *  lands absolute at the dragged spot (Word: dragging breaks the
   *  alignment, the drawn result doesn't shift). A release off every page
   *  commits nothing. The release commits once, so the whole drag is ONE
   *  undo step. */
  #applyOffset(_dx: number, _dy: number, clientX?: number, clientY?: number): void {
    if (!this.#sel) return;
    const nodePos = this.#host.drawingSelection(this.#sel);
    if (nodePos == null) return;
    const target =
      clientX == null || clientY == null ? null : this.#host.pageAtPoint(clientX, clientY);
    if (!target) return;
    const dropX = Math.min(
      Math.max(target.x - (this.#grab?.x ?? 0), 0),
      Math.max(0, target.w - this.#sel.width),
    );
    const dropY = Math.min(
      Math.max(target.y - (this.#grab?.y ?? 0), 0),
      Math.max(0, target.h - this.#sel.height),
    );
    if (
      target.page !== this.#sel.page ||
      this.#host.crossesCell?.(this.#sel, target.page, target.x, target.y)
    ) {
      this.#reanchorDrop(target, dropX, dropY);
      return;
    }
    const anchors = this.#floatingAnchors();
    if (typeof anchors?.h?.offset === "number" && typeof anchors.v?.offset === "number") {
      // Same-page algebra: the clamped target minus the current painted box
      // is the delta toward it — bounded by the page, unlike the raw pointer
      // delta (the re-wrap may shift the anchor; the correction settles it).
      const h = Math.round((dropX - this.#sel.x) * EMU_PER_PX);
      const v = Math.round((dropY - this.#sel.y) * EMU_PER_PX);
      this.#dropCorrection = {
        x: dropX,
        y: dropY,
        page: target.page,
        tries: 0,
        reanchored: false,
        box: { x: this.#sel.x, y: this.#sel.y },
        last: { h, v },
      };
      this.#host.editor().commands["move-drawing"](JSON.stringify({ h, v }));
      return;
    }
    this.#host.editor().commands["place-drawing"](
      JSON.stringify({
        h: Math.round(dropX * EMU_PER_PX),
        v: Math.round(dropY * EMU_PER_PX),
      }),
    );
  }

  /** Re-home the drawing beside a drop point (Word re-anchors on drag): a
   *  cross-page drop, or one whose paragraph sits in a different table cell
   *  than the anchor. The anchor paragraph becomes the drop point's nearest
   *  paragraph and the offsets are seeded from that paragraph's laid origin,
   *  then the correction loop settles the residual as usual. Null when no
   *  paragraph can host the drawing — the caller commits nothing. */
  #reanchorDrop(
    target: { page: number; x: number; y: number },
    dropX: number,
    dropY: number,
  ): void {
    const at = this.#host.paragraphAt(target.page, target.x, target.y);
    if (!at) return;
    this.#dropCorrection = {
      x: dropX,
      y: dropY,
      page: target.page,
      tries: 0,
      reanchored: true,
    };
    this.#host.editor().commands["reanchor-drawing"](
      JSON.stringify({
        to: at.contentPos,
        h: Math.round((dropX - at.xPx) * EMU_PER_PX),
        v: Math.round((dropY - at.yPx) * EMU_PER_PX),
      }),
    );
  }

  /** The landing check after a move-drawing drop: measure the painted box
   *  against the drop point and re-commit the residue. The corrected offset
   *  re-wraps the page around a box already at the drop point, so the anchor
   *  settles and the residue is zero on the next pass — bounded anyway, an
   *  unterminated layout gap must not loop the correction. A box on another
   *  page than the drop means the re-wrap pushed the anchor paragraph itself
   *  off its page; the drawing then re-homes beside the drop point (Word
   *  re-anchors on drag), once — further cross-page drift just gives up. */
  #correctDrop(): void {
    const c = this.#dropCorrection!;
    if (!this.#sel) {
      this.#dropCorrection = null;
      return;
    }
    if (this.#sel.page !== c.page) {
      if (c.reanchored || c.tries > 0) {
        this.#dropCorrection = null;
        return;
      }
      this.#reanchorDrop(c, c.x, c.y);
      return;
    }
    const rx = c.x - this.#sel.x;
    const ry = c.y - this.#sel.y;
    if (Math.hypot(rx, ry) < 1 || c.tries >= 3) {
      this.#dropCorrection = null;
      return;
    }
    // A commit that painted no movement is being eaten by a clamp (a
    // cell-anchored box pinned inside its cell, Word's layoutInCell) —
    // repeating the residue only piles offsets onto a stuck box. With the
    // anchor kept, the last delta retreats so the attrs land back where the
    // paint sits; a re-homed anchor can only give up (its seed was
    // page-clamped, so the residue stays bounded by the anchor's shift).
    if (c.box && Math.hypot(c.box.x - this.#sel.x, c.box.y - this.#sel.y) < 1) {
      if (!c.reanchored && c.last) {
        this.#host
          .editor()
          .commands["move-drawing"](JSON.stringify({ h: -c.last.h, v: -c.last.v }));
      }
      this.#dropCorrection = null;
      return;
    }
    c.box = { x: this.#sel.x, y: this.#sel.y };
    c.last = { h: Math.round(rx * EMU_PER_PX), v: Math.round(ry * EMU_PER_PX) };
    c.tries++;
    this.#host
      .editor()
      .commands["move-drawing"](
        JSON.stringify({ h: Math.round(rx * EMU_PER_PX), v: Math.round(ry * EMU_PER_PX) }),
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

/** Two hits identify the same drawing: same host paragraph, drawing index,
 *  kind, and group-member path. */
function sameHit(a: DrawingHit, b: DrawingHit): boolean {
  const path = (p?: readonly number[]): string => (p ?? []).join(",");
  return (
    a.para === b.para &&
    a.index === b.index &&
    a.kind === b.kind &&
    path(a.childPath) === path(b.childPath)
  );
}
