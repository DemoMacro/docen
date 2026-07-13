// leafer-editor bundles editor + resize + text + view + export plugins — one
// import auto-creates editor/tree/sky layers on App({ editor: {} }). Kept as a
// static side-effect import: its module augmentation types the App `editor`
// config, so a dynamic import would lose type-safety on IAppConfig.editor.
import "leafer-editor";
import { hasCrop } from "@docen/core/geometry";
import { parseImage, renderCropBox, renderImage, type RenderImageInput } from "@docen/core/image";
// Fluent System Icons — same ?raw import pattern as ribbon/icons.ts.
import arrowExpand from "@fluentui/svg-icons/icons/arrow_expand_24_regular.svg?raw";
import resetIcon from "@fluentui/svg-icons/icons/arrow_reset_24_regular.svg?raw";
import arrowRotateClockwise from "@fluentui/svg-icons/icons/arrow_rotate_clockwise_24_regular.svg?raw";
import arrowRotateCounterclockwise from "@fluentui/svg-icons/icons/arrow_rotate_counterclockwise_24_regular.svg?raw";
import dismissIcon from "@fluentui/svg-icons/icons/dismiss_24_regular.svg?raw";
import flipHorizontalIcon from "@fluentui/svg-icons/icons/flip_horizontal_24_regular.svg?raw";
import flipVerticalIcon from "@fluentui/svg-icons/icons/flip_vertical_24_regular.svg?raw";
import zoomFit from "@fluentui/svg-icons/icons/zoom_fit_24_regular.svg?raw";
import zoomIn from "@fluentui/svg-icons/icons/zoom_in_24_regular.svg?raw";
import zoomOut from "@fluentui/svg-icons/icons/zoom_out_24_regular.svg?raw";
import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";
import { App, Box, Image, type IAppConfig } from "leafer-ui";

/**
 * `<docen-image>` — Office-style image editor.
 *
 * Follows the MS Office model: the image is always selectable/editable inline.
 * Click the image → LeaferJS editor shows 8-way resize + rotation handles in
 * place (no modal, no separate "edit mode"); a floating Fluent toolbar appears
 * at the bottom-inside of the canvas with zoom / rotate / flip / reset / expand
 * / delete. A Format-pane or future pptx slide canvas can mount the same
 * component standalone.
 *
 * The Expand button opens an OPTIONAL fullscreen overlay (dark mask + centered
 * image + close + bottom toolbar) for adjusting on a large canvas — a
 * convenience, not the primary editing path. The overlay is a
 * `position:fixed; inset:0` div, NOT browser-native requestFullscreen, so it
 * works inside iframes / shadow DOM and doesn't trap ESC at the OS level.
 *
 * OOXML image data arrives as JSON `attrs`; user edits emit `change` with
 * `{ x, y, width, height, rotation }`.
 */

const ICONS = {
  zoomIn,
  zoomOut,
  zoomFit,
  reset: resetIcon,
  rotateLeft: arrowRotateCounterclockwise,
  rotateRight: arrowRotateClockwise,
  flipH: flipHorizontalIcon,
  flipV: flipVerticalIcon,
  expand: arrowExpand,
  dismiss: dismissIcon,
} as const;

/** Title → icon SVG mapping, used by #injectToolbarIcons. */
const TITLE_ICON: Record<string, string> = {
  "Zoom out": ICONS.zoomOut,
  "Zoom in": ICONS.zoomIn,
  "Fit to screen": ICONS.zoomFit,
  "Rotate left 90°": ICONS.rotateLeft,
  "Rotate right 90°": ICONS.rotateRight,
  "Flip horizontal": ICONS.flipH,
  "Flip vertical": ICONS.flipV,
  "Reset position & size": ICONS.reset,
  Expand: ICONS.expand,
  Delete: ICONS.dismiss,
  Close: ICONS.dismiss,
};

const template = html<DocenImage>`
  <!-- Floating toolbar — appears ABOVE the image when selected. -->
  <div
    class="toolbar"
    ?hidden="${(x) => !x.selected}"
    @mousedown="${(_, c) => c.event.stopPropagation()}"
    @click="${(_, c) => c.event.stopPropagation()}"
  >
    <fluent-button
      appearance="subtle"
      title="Zoom out"
      @click="${(x) => x.zoomOut()}"
    ></fluent-button>
    <fluent-button
      appearance="subtle"
      title="Zoom in"
      @click="${(x) => x.zoomIn()}"
    ></fluent-button>
    <fluent-button
      appearance="subtle"
      title="Fit to screen"
      @click="${(x) => x.fitToScreen()}"
    ></fluent-button>
    <span class="divider"></span>
    <fluent-button
      appearance="subtle"
      title="Rotate left 90°"
      @click="${(x) => x.rotate(-90)}"
    ></fluent-button>
    <fluent-button
      appearance="subtle"
      title="Rotate right 90°"
      @click="${(x) => x.rotate(90)}"
    ></fluent-button>
    <fluent-button
      appearance="subtle"
      title="Flip horizontal"
      @click="${(x) => x.flipHorizontal()}"
    ></fluent-button>
    <fluent-button
      appearance="subtle"
      title="Flip vertical"
      @click="${(x) => x.flipVertical()}"
    ></fluent-button>
    <span class="divider"></span>
    <fluent-button
      appearance="subtle"
      title="Reset position & size"
      @click="${(x) => x.reset()}"
    ></fluent-button>
    <fluent-button
      appearance="subtle"
      title="Expand"
      @click="${(x) => x.openOverlay()}"
    ></fluent-button>
    <slot name="toolbar-actions"></slot>
  </div>
  <div class="canvas-host" ${ref("host")}></div>
  <!-- Fullscreen preview overlay (fixed-position overlay, NOT requestFullscreen). -->
  <div
    class="overlay"
    ?hidden="${(x) => !x.overlayOpen}"
    @click="${(x, c) => x.onOverlayClick(c.event as MouseEvent)}"
  >
    <div class="overlay-close">
      <fluent-button
        appearance="subtle"
        title="Close"
        @click="${(x) => x.closeOverlay()}"
      ></fluent-button>
    </div>
    <div class="overlay-canvas" ${ref("overlayHost")}></div>
    <div class="overlay-toolbar">
      <fluent-button
        appearance="subtle"
        title="Zoom out"
        @click="${(x) => x.zoomOut()}"
      ></fluent-button>
      <fluent-button
        appearance="subtle"
        title="Zoom in"
        @click="${(x) => x.zoomIn()}"
      ></fluent-button>
      <fluent-button
        appearance="subtle"
        title="Fit to screen"
        @click="${(x) => x.fitToScreen()}"
      ></fluent-button>
      <span class="divider"></span>
      <fluent-button
        appearance="subtle"
        title="Rotate left 90°"
        @click="${(x) => x.rotate(-90)}"
      ></fluent-button>
      <fluent-button
        appearance="subtle"
        title="Rotate right 90°"
        @click="${(x) => x.rotate(90)}"
      ></fluent-button>
      <fluent-button
        appearance="subtle"
        title="Flip horizontal"
        @click="${(x) => x.flipHorizontal()}"
      ></fluent-button>
      <fluent-button
        appearance="subtle"
        title="Flip vertical"
        @click="${(x) => x.flipVertical()}"
      ></fluent-button>
      <span class="divider"></span>
      <fluent-button
        appearance="subtle"
        title="Reset position & size"
        @click="${(x) => x.reset()}"
      ></fluent-button>
    </div>
  </div>
`;

const styles = css`
  :host {
    display: inline-block;
    position: relative;
    width: fit-content;
    vertical-align: middle;
  }
  /* canvas-host wraps the LeaferJS App. overflow:visible so the editor's
   * resize handles (which render slightly outside the image bounds) are not
   * clipped. user-select:none + -webkit-user-drag:none kill the browser's
   * native text-selection and image-drag gestures that compete with LeaferJS's
   * resize drag (a native dragstart on the <img> under the pointer-events:none
   * canvas is what triggers ProseMirror's TransformError on drop). */
  .canvas-host {
    display: inline-block;
    box-sizing: content-box;
    overflow: visible;
    user-select: none;
    -webkit-user-select: none;
    -webkit-user-drag: none;
  }
  /* Toolbar — a centered card floating ABOVE the image edge (mini-toolbar /
   * BubbleMenu pattern). Uses Fluent UI tokens so it follows the active theme.
   * Absolute so it doesn't add to the host's height when it appears. */
  .toolbar {
    position: absolute;
    bottom: 100%;
    left: 50%;
    transform: translateX(-50%);
    margin-bottom: 4px;
    display: flex;
    align-items: center;
    gap: 1px;
    padding: 2px 4px;
    background: var(--docen-color-bg, #fff);
    border: 1px solid var(--docen-color-divider, #e2e2e2);
    border-radius: 4px;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.12);
    white-space: nowrap;
  }
  .toolbar[hidden] {
    display: none;
  }
  /* Ribbon-style buttons: 26px height, subtle (no border), 16px icon. */
  .toolbar fluent-button,
  .overlay-toolbar fluent-button,
  .overlay-close fluent-button {
    min-height: 26px;
    min-width: 26px;
    padding: 0 3px;
    font-size: 12px;
  }
  .toolbar fluent-button::part(content),
  .overlay-toolbar fluent-button::part(content),
  .overlay-close fluent-button::part(content) {
    display: flex;
    justify-content: center;
    align-items: center;
  }
  .toolbar fluent-button svg,
  .overlay-toolbar fluent-button svg,
  .overlay-close fluent-button svg {
    width: 16px;
    height: 16px;
    fill: currentColor;
  }
  .toolbar .divider,
  .overlay-toolbar .divider {
    width: 1px;
    height: 16px;
    background: var(--docen-color-divider, #e2e2e2);
    margin: 0 3px;
  }
  /* Fullscreen preview overlay (fixed-position). */
  .overlay {
    position: fixed;
    inset: 0;
    z-index: 3000;
    background: rgba(0, 0, 0, 0.85);
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .overlay[hidden] {
    display: none;
  }
  .overlay-close {
    position: absolute;
    top: 16px;
    right: 16px;
    z-index: 2;
  }
  .overlay-canvas {
    width: 100%;
    height: 100%;
  }
  /*
   * The overlay toolbar reuses the SAME ribbon toolbar style (white card +
   * dark icons) as the editor-mode .toolbar — no frosted/transparent variant.
   * Keeps button appearance consistent between normal and fullscreen modes.
   */
  .overlay-toolbar {
    position: absolute;
    bottom: 24px;
    left: 50%;
    transform: translateX(-50%);
    display: flex;
    align-items: center;
    gap: 1px;
    padding: 2px 4px;
    background: var(--docen-color-bg, #fff);
    border: 1px solid var(--docen-color-divider, #e2e2e2);
    border-radius: 4px;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.12);
    white-space: nowrap;
    z-index: 2;
  }
`;

/** Padding around the canvas so LeaferJS editor handles (which render at the
 *  element bounds, slightly outside) have clickable room. */
const HANDLE_PAD = 8;

@customElement({ name: "docen-image", template, styles })
class DocenImage extends FASTElement {
  /** OOXML image data as JSON string. */
  @attr attrs?: string;
  /**
   * Whether the image is selected (toolbar visible). Toggled by the host:
   * a ProseMirror NodeView drives this from selectNode()/deselectNode();
   * standalone usage toggles it on click. The component itself does NOT
   * manage selection — it only reacts to this property.
   *
   * Uses @observable (not @attr) so template bindings update on direct property
   * assignment — @attr's boolean mode routes through attributeChangedCallback
   * which doesn't reliably notify the ?hidden binding on property set.
   */
  @observable selected = false;
  @observable host?: HTMLElement;
  @observable overlayHost?: HTMLElement;
  @observable overlayOpen = false;

  /** Show/hide the LeaferJS editor frame + handles when selected toggles.
   *  When `selected` flips true before #app exists (a NodeView's selectNode can
   *  fire before hostChanged mounts the canvas), this early-returns; #render
   *  applies the pending selection once the canvas is ready. */
  selectedChanged(): void {
    const editor = this.#app?.editor;
    if (!editor) return;
    editor.visible = this.selected;
    if (this.selected && this.#image) editor.select(this.#image);
  }

  #app?: App;
  #overlayApp?: App;
  // The edited element is a plain `Image`, or — when an OOXML srcRect crop is
  // present — a `Box { overflow:'hide' }` wrapping a larger inner Image (see
  // #createImage). Either way #emit reads it via parseImage, which uses local
  // width/height so rotation round-trips without dimension drift.
  #image?: Image | Box;
  #overlayImage?: Image | Box;
  #mounted = false;

  attrsChanged(): void {
    this.#render();
  }

  hostChanged(): void {
    if (this.host && !this.#mounted) this.#mount();
  }

  connectedCallback(): void {
    super.connectedCallback();
    if (this.host && !this.#mounted) this.#mount();
    this.#render();
    this.#injectIcons(".toolbar");
  }

  disconnectedCallback(): void {
    this.#destroy();
    this.#destroyOverlay();
    document.removeEventListener("keydown", this.#onKeyDown);
    super.disconnectedCallback();
  }

  /** Inject SVG icons into toolbar buttons matching their title attr. */
  #injectIcons(scope: string): void {
    const buttons = this.shadowRoot?.querySelectorAll<HTMLElement>(`${scope} fluent-button`);
    buttons?.forEach((btn) => {
      const title = btn.getAttribute("title") ?? "";
      const svg = TITLE_ICON[title];
      if (svg) btn.innerHTML = svg;
    });
  }

  #mount(): void {
    if (!this.host || this.#mounted) return;
    this.#mounted = true;
    // editor: {} makes LeaferJS auto-create the tree + sky render layers and
    // wires selection/resize/rotate handles. Without it App creates no canvas.
    // The editor frame/handle colors use Fluent UI brand tokens so they follow
    // the active light/dark theme (set via applyTheme → CSS custom properties).
    const brandStroke = this.#cssVar("--colorBrandStroke1") ?? "#0f6cbd";
    const brandBg = this.#cssVar("--colorBrandBackground") ?? "#0078d4";
    const neutralBg = this.#cssVar("--colorNeutralBackground1") ?? "#ffffff";
    const config: IAppConfig = {
      view: this.host,
      fill: "transparent",
      editor: {
        stroke: brandStroke,
        pointFill: brandBg,
        pointStroke: neutralBg,
        // Office model: an inline image can be resized/rotated via handles but
        // NOT moved by dragging its body (moving is meaningless — the canvas IS
        // the image frame, and ProseMirror owns document-flow placement).
        moveable: false,
      },
    };
    this.#app = new App(config);
    // LeaferJS's Interaction listens for native `wheel` on the view div to zoom
    // the canvas. That swallows the wheel event (preventDefault + stopPropaga-
    // tion), so scrolling the document while the cursor is over an image gets
    // stuck. Intercept wheel at the very first capture phase and let it pass
    // through — the image is a static document element, not a zoom surface.
    this.host.addEventListener("wheel", (e) => e.stopImmediatePropagation(), true);
    const editor = this.#app.editor;
    if (editor) {
      // Hidden until the host sets selected=true (selectedChanged → visible).
      editor.visible = false;
      // LeaferJS editor fires `editor.scale` / `editor.rotate` every frame DURING
      // a handle drag (there is no endScale/endRotate event). Touching app.resize()
      // or img.x/y mid-drag corrupts the editor's coordinate state (→ NaN warnings
      // and a broken resize). So we only mark dirty during the drag, and commit
      // the final size on pointerup — matches MS Office, which resizes the frame
      // live but commits layout/reflow on mouse-up.
      editor.on("editor.scale", () => {
        this.#dirty = true;
      });
      editor.on("editor.rotate", () => {
        this.#dirty = true;
      });
    }
    // pointerup ends any editor drag — commit the pending resize/rotate.
    // Listen on window (not host) because the drag's pointerup may fire outside
    // the host div (LeaferJS captures the drag via its own interaction layer).
    window.addEventListener("pointerup", this.#onHostPointerUp);
    // Render the image now that the app exists — connectedCallback's #render
    // call ran before #mount (hostChanged fires after the first render), so
    // without this the canvas stays empty.
    this.#render();
  }

  /** Dirty flag set during editor drag; flushed on pointerup. */
  #dirty = false;

  /** Commit any pending resize/rotate once the drag handle is released. */
  #onHostPointerUp = (): void => {
    if (!this.#dirty) return;
    this.#dirty = false;
    // Defer to the next frame so LeaferJS's editor finishes its own drag
    // state cleanup before we read boxBounds + resize the canvas.
    requestAnimationFrame(() => {
      this.#syncCanvasSize();
      this.#emit();
    });
  };

  /** Read a CSS custom property from the host element's computed style. */
  #cssVar(name: string): string | undefined {
    const val = getComputedStyle(this).getPropertyValue(name).trim();
    return val || undefined;
  }

  #render(): void {
    const app = this.#app;
    if (!app) return;
    const input = this.#parsedAttrs();
    if (!input?.src) {
      this.#image?.remove();
      this.#image = undefined;
      return;
    }
    // Rebuild only when the wrapping structure flips (plain Image ↔ crop Box);
    // otherwise update in place via set() so the bitmap is NOT re-decoded on
    // every attrs change (a data: URL re-load would flicker on each resize).
    const wantBox = hasCrop(input.crop);
    const isBox = this.#image instanceof Box;
    if (this.#image && isBox !== wantBox) {
      this.#image.remove();
      this.#image = undefined;
    }
    if (!this.#image) {
      this.#image = this.#createImage(input);
      app.tree?.add(this.#image);
    } else {
      this.#updateImage(this.#image, input);
    }
    // Size the canvas to the image FIRST — Office model: the image frame IS
    // the canvas, handles sit on the edge, no empty margin around it.
    // Must resize before select so the editor lays out handles on the final
    // canvas size (selecting first then resizing leaves the editor's sky-layer
    // rendering stale → handles invisible).
    this.#syncCanvasSize();
    // Apply the current selection once the canvas is ready. `selected` may have
    // been set true BEFORE #app existed (a NodeView's selectNode can fire before
    // hostChanged mounts the canvas) — selectedChanged early-returns then, so
    // this is the catch-up that shows the editor frame + handles.
    if (this.selected && this.#image) {
      const editor = app.editor;
      if (editor) {
        editor.visible = true;
        editor.select(this.#image);
      }
    }
  }

  /** Build a fresh editable element for `input`: a plain `Image`, or — when an
   *  OOXML srcRect crop is present — a `Box { overflow:'hide' }` wrapping a
   *  larger inner `Image` sized/offset via {@link renderCropBox}. The Box is the
   *  editable frame (resize/rotate handles bind to it); the inner image realizes
   *  the crop, mirroring `@docen/docx`'s span[crop]>img math (renderCropAttrs). */
  #createImage(input: RenderImageInput): Image | Box {
    if (hasCrop(input.crop)) return this.#createCroppedBox(input);
    return new Image(renderImage(input));
  }

  /** Update an existing element in place without re-decoding the bitmap. */
  #updateImage(el: Image | Box, input: RenderImageInput): void {
    if (el instanceof Box) {
      this.#updateCroppedBox(el, input);
    } else {
      el.set(renderImage(input));
    }
  }

  #createCroppedBox(input: RenderImageInput): Box {
    const outer = renderImage(input);
    // renderImage clamps width/height to its 400×300 fallback, so both are
    // always set — assert for the crop math below.
    const w = outer.width as number;
    const h = outer.height as number;
    const { innerWidth, innerHeight, offsetX, offsetY } = renderCropBox(input.crop!, w, h);
    return new Box({
      width: w,
      height: h,
      x: 0,
      y: 0,
      rotation: outer.rotation ?? 0,
      // overflow:'hide' clips the inner image to the Box extent — the canvas
      // equivalent of the HTML overflow:hidden crop frame in renderCropAttrs.
      overflow: "hide",
      editable: true,
      children: [
        new Image({
          url: input.src,
          width: innerWidth,
          height: innerHeight,
          x: offsetX,
          y: offsetY,
          // The inner image is not directly selectable — the editor operates on
          // the wrapping Box so resize/rotate keep the crop frame as the unit.
          editable: false,
        }),
      ],
    });
  }

  #updateCroppedBox(box: Box, input: RenderImageInput): void {
    const outer = renderImage(input);
    const w = outer.width as number;
    const h = outer.height as number;
    const { innerWidth, innerHeight, offsetX, offsetY } = renderCropBox(input.crop!, w, h);
    box.set({ width: w, height: h, rotation: outer.rotation ?? 0 });
    const inner = box.children?.[0] as Image | undefined;
    // set() on the inner image updates crop geometry without re-decoding when
    // the url is unchanged (LeaferJS caches by url) — no flicker on resize.
    inner?.set({ url: input.src, width: innerWidth, height: innerHeight, x: offsetX, y: offsetY });
  }

  /**
   * Sync the canvas size to the image's current size + HANDLE_PAD on each side.
   *
   * Uses app.tree.width/height to set the CONTENT layer's element bounds (this
   * is what LeaferJS's editor reads for handle placement / resize math), and
   * directly sets the `.leafer-app-view` wrapper div's CSS size + each layer
   * canvas's CSS size so the visible frame grows with the image.
   *
   * We deliberately AVOID app.resize() here: it re-allocates every child
   * layer's canvas backing store and resets the editor's sky-layer coordinate
   * state, which corrupts the editor's drag-start bounds and makes every
   * subsequent handle resize compute boxBounds.width / NaN. Setting tree.width
   * + the view div's CSS size achieves the same visible result without
   * touching the editor's internal coordinate system.
   */
  #syncCanvasSize(): void {
    const app = this.#app;
    const img = this.#image;
    if (!app?.tree || !img || !this.host) return;
    const b = img.boxBounds as { width: number; height: number };
    const w = Math.max(1, Math.round(b.width));
    const h = Math.max(1, Math.round(b.height));
    // Keep the image offset by HANDLE_PAD so the editor's handles sit inside.
    img.x = HANDLE_PAD;
    img.y = HANDLE_PAD;
    const full = w + HANDLE_PAD * 2;
    const fullH = h + HANDLE_PAD * 2;
    // Content layer bounds — the editor reads these for handle layout / resize.
    app.tree.width = full;
    app.tree.height = fullH;
    // Wrapper div CSS size (LeaferJS's auto-layout would do this, but once we
    // set tree.width it stops tracking — so set it manually).
    const appView = this.host.querySelector(".leafer-app-view") as HTMLElement | null;
    if (appView) {
      appView.style.width = `${full}px`;
      appView.style.height = `${fullH}px`;
    }
  }

  #emit(): void {
    if (!this.#image) return;
    this.$emit("change", parseImage(this.#image));
  }

  #parsedAttrs(): RenderImageInput | null {
    try {
      return JSON.parse(this.attrs ?? "{}") as RenderImageInput;
    } catch {
      return null;
    }
  }

  #destroy(): void {
    window.removeEventListener("pointerup", this.#onHostPointerUp);
    this.#app?.destroy();
    this.#app = undefined;
    this.#image = undefined;
    this.#mounted = false;
  }

  // ── Zoom ──

  /** Get the active app/tree for zoom (overlay or main depending on mode). */
  #activeApp(): App | undefined {
    return this.overlayOpen ? this.#overlayApp : this.#app;
  }

  zoomIn(): void {
    this.#activeApp()?.tree?.zoom("in");
  }
  zoomOut(): void {
    this.#activeApp()?.tree?.zoom("out");
  }
  fitToScreen(): void {
    this.#activeApp()?.tree?.zoom("fit");
  }

  // ── Rotate / Flip ──

  rotate(delta: number): void {
    const img = this.overlayOpen ? this.#overlayImage : this.#image;
    if (!img) return;
    // rotateOf('center') pivots around the image center so it stays in frame.
    // Do NOT call #syncCanvasSize here — it would reset img.x/y to HANDLE_PAD,
    // overriding the center-anchored position rotateOf just computed (making
    // the rotation appear to pivot from the top-left, not the center). The
    // canvas size stays at the unrotated boxBounds + pad; rotated corners
    // extend past it but .canvas-host{overflow:visible} shows them.
    img.rotateOf("center", delta);
    this.#emit();
  }
  flipHorizontal(): void {
    const img = this.overlayOpen ? this.#overlayImage : this.#image;
    if (!img) return;
    img.flip("x");
    this.#emit();
  }
  flipVertical(): void {
    const img = this.overlayOpen ? this.#overlayImage : this.#image;
    if (!img) return;
    img.flip("y");
    this.#emit();
  }

  // ── Reset / Delete ──

  reset(): void {
    const input = this.#parsedAttrs();
    if (!input?.src) return;
    if (this.overlayOpen) {
      const app = this.#overlayApp;
      if (!app) return;
      // Rebuild the overlay image from the original attrs. A bare
      // img.set(renderImage(input)) would push Image-only options onto a crop
      // Box, so reuse the same factory as #render.
      this.#overlayImage?.remove();
      this.#overlayImage = this.#createImage(input);
      app.tree?.add(this.#overlayImage);
      app.editor?.select(this.#overlayImage);
    } else {
      this.#render();
    }
    this.#emit();
  }
  /**
   * Request deletion of the image. The component does NOT delete itself — it
   * only emits a `delete` intent event. The host (a ProseMirror NodeView, or a
   * standalone container) listens and performs the actual removal. This mirrors
   * the Office.js model where `InlinePicture.delete()` is called on the
   * document object, not inside the picture's own render surface.
   */
  deleteSelected(): void {
    this.$emit("delete");
  }

  // ── Fullscreen overlay ──

  openOverlay(): void {
    if (this.overlayOpen) return;
    this.overlayOpen = true;
    // Mount the overlay LeaferJS app + image after the DOM updates.
    requestAnimationFrame(() => {
      // Guard against a rapid closeOverlay() scheduled before this frame: rAF
      // cannot be cancelled, so re-check overlayOpen (plus host + existing app)
      // before creating — otherwise a fast open→close would resurrect a
      // destroyed overlay on the next frame.
      if (!this.overlayOpen || !this.overlayHost || this.#overlayApp) return;
      this.#overlayApp = new App({
        view: this.overlayHost,
        fill: "transparent",
        editor: {},
      });
      const input = this.#parsedAttrs();
      if (input?.src) {
        this.#overlayImage = this.#createImage(input);
        this.#overlayApp.tree?.add(this.#overlayImage);
        this.#overlayApp.editor?.select(this.#overlayImage);
      }
      this.#injectIcons(".overlay-toolbar");
      this.#injectIcons(".overlay-close");
      // Fit after the canvas has its real dimensions.
      requestAnimationFrame(() => this.fitToScreen());
    });
    document.addEventListener("keydown", this.#onKeyDown);
  }

  closeOverlay(): void {
    if (!this.overlayOpen) return;
    this.overlayOpen = false;
    this.#destroyOverlay();
    document.removeEventListener("keydown", this.#onKeyDown);
  }

  #destroyOverlay(): void {
    this.#overlayApp?.destroy();
    this.#overlayApp = undefined;
    this.#overlayImage = undefined;
  }

  /**
   * Click on the dark mask closes the overlay. The LeaferJS canvas fills the
   * whole overlay (so it, not .overlay, is the click target for empty mask
   * area) — treat clicks on .overlay or .overlay-canvas as "mask clicks".
   * Clicks on the close button / toolbar bubble up too, but their target is a
   * fluent-button, so they're ignored.
   */
  onOverlayClick(e: MouseEvent): void {
    const el = e.target as Element;
    if (el === e.currentTarget || el?.classList?.contains("overlay-canvas")) {
      this.closeOverlay();
    }
  }

  /** ESC closes the overlay. */
  #onKeyDown = (e: KeyboardEvent): void => {
    if (e.key === "Escape") this.closeOverlay();
  };
}

export default DocenImage;
