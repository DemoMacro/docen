const template = document.createElement("template");
template.innerHTML = `
  <style>
    :host { display: contents; }
    fluent-dialog-body { width: 100%; }
    /* fluent-dialog-body only lays the footer out as a right-aligned row once
       the dialog-body itself is ≥480px wide (@container). A compact dialog
       never reaches that, leaving OK/Cancel stacked vertically — force the
       row here so the action footer is always right-aligned. */
    fluent-dialog-body::part(actions) {
      flex-direction: row;
      justify-content: flex-end;
      align-items: center;
      gap: 8px;
      padding-block-start: var(--spacingVerticalXL, 20px);
    }
  </style>
  <fluent-dialog type="modal" part="dialog">
    <fluent-dialog-body part="body">
      <h2 slot="title" part="title"></h2>
      <slot name="title-action" slot="title-action"></slot>
      <fluent-button
        slot="close"
        part="close"
        tabindex="0"
        appearance="transparent"
        icon-only
        aria-label="Close"
      >
        <svg fill="currentColor" aria-hidden="true" width="20" height="20" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg">
          <path d="m4.09 4.22.06-.07a.5.5 0 0 1 .63-.06l.07.06L10 9.29l5.15-5.14a.5.5 0 0 1 .63-.06l.07.06c.18.17.2.44.06.63l-.06.07L10.71 10l5.14 5.15c.18.17.2.44.06.63l-.06.07a.5.5 0 0 1-.63.06l-.07-.06L10 10.71l-5.15 5.14a.5.5 0 0 1-.63.06l-.07-.06a.5.5 0 0 1-.06-.63l.06-.07L9.29 10 4.15 4.85a.5.5 0 0 1-.06-.63l.06-.07-.06.07Z" fill="currentColor" />
        </svg>
      </fluent-button>
      <slot></slot>
      <slot name="action" slot="action"></slot>
    </fluent-dialog-body>
  </fluent-dialog>`;

/** Structural subset of fluent-dialog the dialog forwards to. */
interface FluentDialog extends HTMLElement {
  show(): void;
  hide(): void;
}

interface FluentToggleEvent extends Event {
  detail?: { newState?: string; oldState?: string };
}

/**
 * `<docen-dialog heading="…" open>` — a generic modal dialog wrapping
 * `<fluent-dialog type="modal">` + `<fluent-dialog-body>` (title / content /
 * action regions). The default slot is the body — any fields or content the
 * caller supplies; the `action` slot is the footer (OK/Cancel). `show()`/
 * `hide()` drive the underlying fluent-dialog (modal: showModal → backdrop +
 * ESC). This is a content-agnostic container — it owns no business fields.
 */
class DocenDialog extends HTMLElement {
  static get observedAttributes(): string[] {
    return ["heading", "open"];
  }

  #dialog?: FluentDialog;
  #titleEl?: HTMLElement;

  attributeChangedCallback(name: string): void {
    if (name === "heading") this.#applyHeading();
    if (name === "open") this.#applyOpen();
  }

  connectedCallback(): void {
    if (!this.shadowRoot) {
      this.attachShadow({ mode: "open" }).append(template.content.cloneNode(true));
    }
    this.#dialog = this.shadowRoot!.querySelector("fluent-dialog") as FluentDialog;
    this.#titleEl = this.shadowRoot!.querySelector<HTMLElement>('[part="title"]') ?? undefined;
    // ESC / backdrop close the fluent-dialog directly; sync our `open` attr so
    // state stays consistent. fluent emits `toggle` with newState.
    this.#dialog.addEventListener("toggle", (event) => {
      if ((event as FluentToggleEvent).detail?.newState === "closed") {
        if (this.hasAttribute("open")) this.removeAttribute("open");
      }
    });
    // Office dialogs don't light-dismiss on backdrop click. fluent-dialog's
    // clickHandler hides when the click lands on the native <dialog> itself
    // (the backdrop region); intercept those in capture phase so only ESC, the
    // close button, or Cancel dismisses the dialog.
    this.#disableBackdropDismiss();
    this.#applyHeading();
    this.#applyOpen();
  }

  #disableBackdropDismiss(): void {
    const apply = (): void => {
      const native = this.#dialog?.shadowRoot?.querySelector("dialog");
      if (!native) {
        requestAnimationFrame(apply);
        return;
      }
      native.addEventListener(
        "click",
        (event: Event) => {
          if (event.target === native) event.stopImmediatePropagation();
        },
        true,
      );
    };
    apply();
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

  #applyHeading(): void {
    if (this.#titleEl) this.#titleEl.textContent = this.getAttribute("heading") ?? "";
  }

  #applyOpen(): void {
    const dialog = this.#dialog;
    if (!dialog || typeof dialog.show !== "function") {
      if (dialog) requestAnimationFrame(() => this.#applyOpen());
      return;
    }
    // fluent-dialog, like fluent-drawer, has no reliable `open` reflection —
    // show()/hide() drive the <dialog> directly and are idempotent, so branch
    // on our own `open` attribute alone.
    if (this.open) dialog.show();
    else dialog.hide();
  }
}

customElements.define("docen-dialog", DocenDialog);

export default DocenDialog;
