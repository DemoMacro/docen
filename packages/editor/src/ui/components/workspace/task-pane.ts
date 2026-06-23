import { observeLang, t } from "../../i18n/localize";

const template = document.createElement("template");
template.innerHTML = `
  <style>
    :host { display: block; height: 100%; box-sizing: border-box; }
    /* Drawer-wide small text so the pane reads as a compact side rail, not
       body copy. */
    fluent-drawer { font-size: 12px; }
    /* fluent-drawer's <dialog> is the pane's bounding box. Make it a flex
       column so the header stays fixed and the slotted body (outline tree)
       fills the remaining height and scrolls internally instead of growing past
       the editor (a 7888px outline overflowed the 551px pane before this). */
    fluent-drawer::part(dialog) {
      display: flex;
      flex-direction: column;
      overflow: hidden;
      height: 100%;
    }
    /* Visibility is CSS-driven from our own "open" attribute. An inline
       fluent-drawer hide() does not actually close the dialog (its
       display stays flex), so a pane without open — e.g. the Properties
       pane — renders visible by default and the close button cannot
       dismiss it. Keying display on :host(:not([open])) makes the
       attribute the single source of truth: show()/hide() still update
       the drawer internal state, but the canvas reflow follows the
       attribute. */
    :host(:not([open])) fluent-drawer::part(dialog) {
      display: none;
    }
    /* Header row: title on the inline-start, close button on the inline-end. */
    .panel-head {
      display: flex;
      flex: 0 0 auto;
      align-items: center;
      justify-content: space-between;
      padding: 6px 4px 6px 16px;
      border-bottom: 1px solid var(--docen-color-divider, #e2e2e2);
    }
    .panel-head span { font-size: 13px; font-weight: 600; }
    .panel-close { min-width: 28px; height: 28px; padding: 0; }
    .panel-close[hidden] { display: none; }
  </style>
  <fluent-drawer part="drawer" type="inline" size="small">
    <div class="panel-head" part="head">
      <span part="title"></span>
      <fluent-button class="panel-close" part="close" appearance="subtle">✕</fluent-button>
    </div>
    <slot></slot>
  </fluent-drawer>`;

/** Structural subset of fluent-drawer the pane forwards to. */
interface FluentDrawer extends HTMLElement {
  show(): void;
  hide(): void;
  readonly open: boolean;
  readonly dialog: HTMLElement;
}

/**
 * `<docen-task-pane position="start" title="…" open>` — an Office-style Task
 * Pane: a side rail that pushes the canvas (inline drawer), with a header
 * (title + close button) and a default slot for the pane body (outline tree,
 * property fields, …). Closing is only via the header button — inline drawers
 * ignore ESC and outside clicks, and we additionally block Fluent's
 * click-the-dialog dismissal, so nothing else can dismiss the pane.
 *
 * The close button tooltip is localized (`taskPane.close`); set the
 * workspace's `lang` to override the page locale.
 */
class DocenTaskPane extends HTMLElement {
  static get observedAttributes(): string[] {
    return ["position", "title", "open", "closable"];
  }

  #drawer?: FluentDrawer;
  #titleEl?: HTMLElement;
  #closeBtn?: HTMLElement;
  #unsubscribe?: () => void;

  attributeChangedCallback(name: string): void {
    switch (name) {
      case "position":
        this.#applyPosition();
        break;
      case "title":
        this.#applyTitle();
        break;
      case "open":
        this.#applyOpen();
        break;
      case "closable":
        this.#applyClosable();
        break;
    }
  }

  connectedCallback(): void {
    if (!this.shadowRoot) {
      this.attachShadow({ mode: "open" }).append(template.content.cloneNode(true));
    }
    const root = this.shadowRoot!;
    this.#drawer = root.querySelector("fluent-drawer") as FluentDrawer;
    this.#titleEl = root.querySelector('[part="title"]')!;
    this.#closeBtn = root.querySelector('[part="close"]')!;
    this.#applyPosition();
    this.#applyTitle();
    this.#applyClosable();
    this.#closeBtn?.addEventListener("click", () => this.hide());
    this.#lockBlankClose();
    this.#applyOpen();
    this.#applyI18n();
    this.#unsubscribe = observeLang(() => this.#applyI18n());
  }

  disconnectedCallback(): void {
    this.#unsubscribe?.();
  }

  get open(): boolean {
    return this.hasAttribute("open");
  }

  set open(value: boolean) {
    this.toggleAttribute("open", value);
  }

  show(): void {
    this.open = true;
  }

  hide(): void {
    this.open = false;
  }

  // Sync the `open` attribute to fluent-drawer's show/hide. The drawer may not
  // be upgraded on first connect, so retry once it is.
  #applyOpen(): void {
    const drawer = this.#drawer;
    if (!drawer || typeof drawer.show !== "function") {
      if (drawer) requestAnimationFrame(() => this.#applyOpen());
      return;
    }
    // fluent-drawer has no `open` property reflecting the dialog state —
    // show()/hide() drive the <dialog> directly and are idempotent, so branch
    // on our own `open` attribute alone. (The old `drawer.open` checks were
    // always falsy, making the close branch dead code: the pane's `open`
    // attribute flipped but the dialog stayed open — "state changed, pane
    // didn't actually close".)
    if (this.open) drawer.show();
    else drawer.hide();
  }

  #applyPosition(): void {
    const position = this.getAttribute("position") ?? "start";
    this.#drawer?.setAttribute("position", position);
  }

  #applyTitle(): void {
    if (this.#titleEl) this.#titleEl.textContent = this.getAttribute("title") ?? "";
  }

  #applyClosable(): void {
    const closable = this.getAttribute("closable") !== "false";
    this.#closeBtn?.toggleAttribute("hidden", !closable);
  }

  // Inline drawers also close when a click lands on the <dialog> itself
  // (Fluent's clickHandler treats the dialog as a backdrop). Block those so
  // only the header close button dismisses the pane.
  #lockBlankClose(): void {
    const bind = (): boolean => {
      const dialog = this.#drawer?.dialog;
      if (!dialog) return false;
      dialog.addEventListener(
        "click",
        (event: Event) => {
          if (event.target === dialog) event.stopImmediatePropagation();
        },
        true,
      );
      return true;
    };
    if (!bind()) requestAnimationFrame(bind);
  }

  #applyI18n(): void {
    if (this.#closeBtn) this.#closeBtn.setAttribute("title", t("taskPane.close", this));
  }
}

customElements.define("docen-task-pane", DocenTaskPane);

export default DocenTaskPane;
