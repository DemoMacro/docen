import { emuToPx, pxToEmu } from "@docen/core";
import { Image, renderImageStyles, floatAnchorScope } from "@docen/docx";
import type { ImageAttrs } from "@docen/docx";
import dismissIcon from "@fluentui/svg-icons/icons/dismiss_24_regular.svg?raw";

import "../../ui/components/drawings/image";

/**
 * Editor-only NodeView for the image node.
 *
 * EXTENDS the engine's `Image` — does not re-create it — so its schema
 * (parseHTML/renderHTML/renderDocx/parseDocx + all fidelity attrs) is inherited;
 * only `addNodeView` is added. The engine node is UI-free, so without this the
 * image renders via `renderHTML` (which works and stays SSR-safe) but has no
 * resize handles, no rotation UI, no interactive editing.
 *
 * This NodeView is a thin adapter: it mounts a `<docen-image>` web component
 * (a first-class UI element shared with property panes and the future pptx
 * canvas), forwards the node's `ImageAttrs` as the component's JSON `attrs`,
 * and listens for the component's `change` event to write the edited
 * `{ width, height, rotation }` back into the node via a ProseMirror transaction.
 *
 * Office model: the image is always selectable inline — click it to get
 * resize/rotate handles + the floating toolbar in place, no separate mode.
 *
 * SSR safety: ProseMirror instantiates NodeViews only in a browser view, so
 * this never runs server-side. `generateHTML` / read-mode rendering still uses
 * the engine's `renderHTML` three branches (vector / crop / plain), which are
 * SSR-safe and LeaferJS-free. The LeaferJS canvas lives entirely inside this
 * NodeView's shadow DOM.
 *
 * ImageCap (the append-transaction plugin that auto-scales over-wide images on
 * import) stays registered alongside this NodeView — the two are complementary:
 * ImageCap normalizes import-time sizing; this NodeView handles interactive
 * resize/rotate during editing.
 */
export const ImageView = Image.extend({
  // Disable ProseMirror's node-level drag-and-drop: the engine's Image sets
  // draggable:true by default, so dragging the resize handle (or the image
  // body) triggers a node-drag instead of LeaferJS's resize. The image stays
  // in place — move/cut is done via the document text cursor, not drag.
  draggable: false,

  addNodeView() {
    return ({ node, getPos, editor }) => {
      // Outer wrapper — applies the SAME floating CSS as renderHTML so the
      // editing surface matches the rendered/exported layout (float /
      // position:absolute / z-index / shape-outside). For non-floating images
      // renderImageStyles yields display:inline-block, matching the old
      // hard-coded style. This mirrors WpsShapeView's pattern (it uses
      // wpsShapeStyles → dom.style.cssText) — the editor never re-derives
      // EMU/floating geometry, it reuses the engine's style mapper.
      const dom = document.createElement("span");
      applyFloatingStyles(dom, node.attrs as Record<string, unknown>);

      // Track the latest node so the change/drag handlers read CURRENT attrs
      // (not this create-time snapshot) — an external attrs change between edits
      // (property pane crop/anchor, ImageCap, undo) would otherwise be overwritten
      // by the closed-over node.attrs on the next resize/drag.
      let current = node;

      const editorEl = document.createElement("docen-image") as HTMLElement & {
        attrs: string;
        selected: boolean;
        addEventListener: (type: string, listener: (e: CustomEvent) => void) => void;
      };
      // Forward the current node attrs as the editor component's JSON input.
      editorEl.attrs = JSON.stringify(toImageInput(node.attrs as ImageAttrs));

      // Inject a Delete button into the component's toolbar-actions slot.
      // The component itself is generic (no document semantics); only a
      // document NodeView knows how to delete a ProseMirror node.
      const deleteBtn = document.createElement("fluent-button");
      deleteBtn.setAttribute("appearance", "subtle");
      deleteBtn.setAttribute("title", "Delete");
      deleteBtn.slot = "toolbar-actions";
      deleteBtn.style.cssText = "min-height:26px;min-width:26px;padding:0 3px;";
      deleteBtn.innerHTML = dismissIcon;
      deleteBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        editorEl.dispatchEvent(new CustomEvent("delete"));
      });
      editorEl.append(deleteBtn);

      dom.append(editorEl);

      // Kill native HTML5 drag-and-drop at the source. LeaferJS's canvases have
      // `pointer-events: none`, so a mousedown on a resize handle is dispatched
      // to the element BELOW the canvas. If that element is an <img> (or any
      // natively draggable node), the browser fires a native `dragstart` →
      // ProseMirror's `editHandlers.drop` → `handleDrop` → tries to insert the
      // dragged slice into a paragraph → `TransformError: Invalid content for
      // node paragraph`. Preventing dragstart here stops the whole chain before
      // it starts; the resize drag is a pure pointer gesture, never a DnD drop.
      dom.draggable = false;
      dom.addEventListener("dragstart", (e) => e.preventDefault());

      // Listen for resize/rotate → write back into the node attrs.
      editorEl.addEventListener("change", (event: CustomEvent) => {
        const detail = event.detail as {
          width?: number;
          height?: number;
          rotation?: number;
        } | null;
        if (!detail || detail.width == null) return;
        const pos = getPos();
        if (typeof pos !== "number") return;
        const { width, height, rotation } = detail;
        editor
          .chain()
          .focus()
          .command(({ tr }) => {
            tr.setNodeMarkup(pos, undefined, {
              ...(current.attrs as ImageAttrs),
              width,
              height,
              rotation,
            });
            return true;
          })
          .run();
      });

      // Delete intent from the component → dispatch a `command` event so it
      // routes through the host's #onCommand → Tiptap "delete-picture" command
      // (same path as a ribbon button), instead of mutating the node attrs.
      editorEl.addEventListener("delete", () => {
        dom.dispatchEvent(
          new CustomEvent("command", {
            bubbles: true,
            composed: true,
            detail: { event: "delete-picture" },
          }),
        );
      });

      // Floating drag-to-move: for wrapNone images (position:absolute), let the
      // user drag the image body to reposition it. Drag updates dom left/top
      // live; pointerup dispatches a "position-picture" command event (same
      // routing as a ribbon button → host #onCommand → Tiptap command). A getter
      // reads the CURRENT floating config each gesture, so a property-pane anchor
      // change between drags takes effect without rebinding listeners.
      setupFloatingDrag(dom, editorEl, () => (current.attrs as ImageAttrs).floating);

      return {
        dom,
        // atom — no editable content inside the image.
        contentDOM: null,
        // ProseMirror NodeSelection → show toolbar + LeaferJS handles.
        selectNode: () => {
          editorEl.selected = true;
        },
        deselectNode: () => {
          editorEl.selected = false;
        },
        // Re-forward attrs when the node is updated externally (e.g. ImageCap,
        // property pane, or undo/redo). Re-apply floating CSS so position/wrap
        // stays in sync with the rendered layout.
        update: (updated) => {
          if (updated.type !== node.type) return false;
          current = updated;
          const newAttrs = updated.attrs as Record<string, unknown>;
          editorEl.attrs = JSON.stringify(toImageInput(updated.attrs as ImageAttrs));
          applyFloatingStyles(dom, newAttrs);
          return true;
        },
        // Control which events ProseMirror sees vs. LeaferJS handles.
        //
        // The LeaferJS canvas lives inside <docen-image>'s SHADOW DOM with
        // `pointer-events: none`, so composedPath() must be used to detect
        // whether an event originated from the editor component.
        //
        // Two categories are blocked once the image is selected:
        //  1. Native DnD (dragstart/dragover/drop/…) — would otherwise trigger
        //     ProseMirror's handleDrop → TransformError when the user drags a
        //     resize handle (the browser sees a drag on the <img> under the
        //     pointer-events:none canvas).
        //  2. Pointer/mouse down/move/up — so ProseMirror's contenteditable
        //     doesn't start a text-selection drag that competes with LeaferJS's
        //     resize drag (the cause of the handle "sticking" after a few px).
        stopEvent: (event) => {
          const type = event.type;
          if (type === "dragstart" || type === "dragend") return true;
          if (!editorEl.selected) return false;
          if (
            type === "dragenter" ||
            type === "dragover" ||
            type === "dragleave" ||
            type === "drop"
          )
            return true;
          if (
            type === "mousedown" ||
            type === "mousemove" ||
            type === "mouseup" ||
            type === "pointerdown" ||
            type === "pointermove" ||
            type === "pointerup"
          ) {
            const path = event.composedPath();
            if (!path.includes(editorEl) || path[0] === dom) return false;
            return true;
          }
          return false;
        },
        ignoreMutation: () => true,
      };
    };
  },
});

/** Reduce full ImageAttrs to the subset the editor component renders from. */
const toImageInput = (
  attrs: ImageAttrs,
): {
  src: string;
  width: number | null;
  height: number | null;
  rotation: number | null;
  crop: unknown;
  outline: unknown;
} => ({
  src: attrs.src,
  width: attrs.width,
  height: attrs.height,
  rotation: attrs.rotation,
  crop: attrs.crop,
  outline: attrs.outline,
});

/**
 * Apply the engine's image style (display + rotation + floating placement) to
 * the NodeView's outer dom, so the editing surface matches renderHTML exactly.
 *
 * For non-floating images this yields `display:inline-block` + optional
 * `transform:rotate()` — same as the old hard-coded style. For floating images
 * it yields the full `floatingToStyles` output (`float` / `position:absolute` /
 * `z-index` / `shape-outside`), so a wrapSquare/anchor image stays in place
 * during editing instead of collapsing to inline-block.
 *
 * The `data-float-anchor` attribute is set for paragraph-anchored wrapNone
 * images so the editor's CSS rule `p:has([data-float-anchor])` makes the anchor
 * `<p>` the offsetParent (otherwise the absolute image floats to the page top).
 */
const applyFloatingStyles = (dom: HTMLElement, attrs: Record<string, unknown>): void => {
  // marginOrigin=true: the editor lives inside <docen-document-area>, which
  // defines --docen-page-margin-*, so a margin-anchored floating drawing can
  // use the calc() compensation. generateHTML (the engine's other caller) runs
  // outside the editor and leaves this false for host-agnostic output.
  dom.style.cssText = renderImageStyles(attrs, true).join(";");
  const floating = attrs.floating as unknown;
  if (floating && floatAnchorScope(floating) === "paragraph") {
    dom.setAttribute("data-float-anchor", "paragraph");
  } else {
    dom.removeAttribute("data-float-anchor");
  }
};

/** Handle inset (px) — matches HANDLE_PAD in the docen-image component.
 *  Pointer-downs within this band of the edge are resize-handle drags handled
 *  by LeaferJS; the center region is the move-drag surface. */
const HANDLE_PAD = 8;

/**
 * Enable drag-to-move on a wrapNone floating image.
 *
 * Uses screen-px delta with `dom.offsetLeft/Top` as the base — both are in the
 * same CSS-px coordinate space (offsetLeft is unaffected by CSS zoom, and we
 * divide the screen-px pointer delta by zoom to get CSS px). On pointerdown we
 * snapshot the element's offsetLeft/Top; on pointermove we add the zoomed
 * delta. This avoids the pitfalls of `getBoundingClientRect()` under CSS zoom
 * and of reading `style.left` (which may be `50%` or a keyword).
 *
 * `setPointerCapture` keeps pointermove/up firing even outside the element.
 *
 * Only fires when the image is selected (first click selects). The edge band
 * (HANDLE_PAD px) is left to LeaferJS resize handles.
 */
const setupFloatingDrag = (
  dom: HTMLElement,
  editorEl: HTMLElement,
  getFloating: () =>
    | {
        horizontalPosition: { relative?: string };
        verticalPosition: { relative?: string };
        wrap?: { type?: number };
      }
    | null
    | undefined,
): void => {
  let originX = 0;
  let originY = 0;
  let pointerX = 0;
  let pointerY = 0;
  let zoom = 1;
  let activePointer: number | null = null;
  let savedTransform = "";
  let savedLeft = "";
  let savedTop = "";

  editorEl.addEventListener("pointerdown", (e: PointerEvent) => {
    if (!(editorEl as { selected?: boolean }).selected) return;
    const floating = getFloating();
    // Only wrapNone (position:absolute) images drag-to-move; square/tight/
    // through wrap types use CSS float and have no meaningful free position.
    if (!floating || (floating.wrap?.type ?? 0) !== 0) return;
    const r = editorEl.getBoundingClientRect();
    const fromEdge = Math.min(
      e.clientX - r.left,
      r.right - e.clientX,
      e.clientY - r.top,
      r.bottom - e.clientY,
    );
    if (fromEdge <= HANDLE_PAD) return;
    // Snapshot the live style BEFORE mutating so a sub-threshold click can
    // restore it (the align-clearing + px-freeze below would otherwise strip a
    // centered image's transform and leave it at the frozen px position).
    savedTransform = dom.style.transform;
    savedLeft = dom.style.left;
    savedTop = dom.style.top;
    zoom = readZoom(dom);
    // When the image uses align (e.g. center → left:50% + translateX(-50%)),
    // offsetLeft is the percentage-resolved value and does NOT match the visual
    // position. Freeze the actual visual position into explicit px so the
    // drag delta is measured from where the user sees the image.
    const parent = dom.offsetParent as HTMLElement | null;
    if (parent && dom.style.transform) {
      const pr = parent.getBoundingClientRect();
      dom.style.left = `${Math.round((r.left - pr.left) / zoom)}px`;
      dom.style.top = `${Math.round((r.top - pr.top) / zoom)}px`;
      dom.style.transform = "";
    }
    originX = dom.offsetLeft;
    originY = dom.offsetTop;
    pointerX = e.clientX;
    pointerY = e.clientY;
    activePointer = e.pointerId;
    editorEl.setPointerCapture(e.pointerId);
    e.preventDefault();
  });

  editorEl.addEventListener("pointermove", (e: PointerEvent) => {
    if (activePointer !== e.pointerId) return;
    const dx = (e.clientX - pointerX) / zoom;
    const dy = (e.clientY - pointerY) / zoom;
    dom.style.left = `${Math.round(originX + dx)}px`;
    dom.style.top = `${Math.round(originY + dy)}px`;
  });

  editorEl.addEventListener("pointerup", (e: PointerEvent) => {
    if (activePointer !== e.pointerId) return;
    activePointer = null;
    editorEl.releasePointerCapture(e.pointerId);
    const totalDx = (e.clientX - pointerX) / zoom;
    const totalDy = (e.clientY - pointerY) / zoom;
    // Sub-threshold movement = a click, not a drag: restore the pre-drag style
    // so align (transform) / percentage positions survive the aborted gesture.
    if (Math.abs(totalDx) < 2 && Math.abs(totalDy) < 2) {
      dom.style.transform = savedTransform;
      dom.style.left = savedLeft;
      dom.style.top = savedTop;
      return;
    }
    const floating = getFloating();
    if (!floating) return;
    // Convert dom.offsetLeft/Top (CSS px, relative to offsetParent padding-box =
    // page edge) to OOXML posOffset EMU. When relativeFrom is "margin" the origin
    // is the margin-box (page edge + page padding), so subtract the page margin.
    // The --docen-page-margin-* CSS vars are authored in mm (from OOXML EMU), so
    // bridge mm→EMU (1mm = 36000 EMU) then EMU→px via @docen/core — the same
    // EMU/px source as everywhere else, avoiding a hard-coded 9525 literal.
    const isMargin = (rel: string | undefined): boolean =>
      rel === "margin" || rel === "insideMargin" || rel === "outsideMargin";
    const page = dom.closest(".docen-page") as HTMLElement | null;
    const marginPx = (cssVar: string): number => {
      if (!page) return 0;
      const raw = getComputedStyle(page).getPropertyValue(cssVar).trim();
      const val = parseFloat(raw);
      if (!val) return 0;
      return raw.endsWith("mm") ? emuToPx(val * 36000) : val;
    };
    const padL = marginPx("--docen-page-margin-left");
    const padT = marginPx("--docen-page-margin-top");
    const hPx = dom.offsetLeft - (isMargin(floating.horizontalPosition.relative) ? padL : 0);
    const vPx = dom.offsetTop - (isMargin(floating.verticalPosition.relative) ? padT : 0);
    dom.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: {
          event: "position-picture",
          value: JSON.stringify({
            hOffset: Math.round(pxToEmu(hPx)),
            vOffset: Math.round(pxToEmu(vPx)),
          }),
        },
      }),
    );
  });
};

/** Read the CSS `zoom` factor of the document area so pointer deltas are
 *  converted from screen px back to CSS px. Returns 1 if unset. */
const readZoom = (el: HTMLElement): number => {
  const area = el.closest("docen-document-area") as HTMLElement | null;
  if (!area) return 1;
  const z = parseFloat(getComputedStyle(area).zoom || "1");
  return Number.isFinite(z) && z > 0 ? z : 1;
};
