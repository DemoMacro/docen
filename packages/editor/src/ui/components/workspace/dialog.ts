import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

const styles = css`
  :host {
    display: contents;
  }
  fluent-dialog-body {
    width: 100%;
  }
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
`;

const template = html<DocenDialog>`
  <fluent-dialog type="modal" part="dialog" ${ref("dialog")}>
    <fluent-dialog-body part="body">
      <h2 slot="title" part="title" ${ref("titleEl")}></h2>
      <slot name="title-action" slot="title-action"></slot>
      <fluent-button
        slot="close"
        part="close"
        tabindex="0"
        appearance="transparent"
        icon-only
        aria-label="Close"
      >
        <svg
          fill="currentColor"
          aria-hidden="true"
          width="20"
          height="20"
          viewBox="0 0 20 20"
          xmlns="http://www.w3.org/2000/svg"
        >
          <path
            d="m4.09 4.22.06-.07a.5.5 0 0 1 .63-.06l.07.06L10 9.29l5.15-5.14a.5.5 0 0 1 .63-.06l.07.06c.18.17.2.44.06.63l-.06.07L10.71 10l5.14 5.15c.18.17.2.44.06.63l-.06.07a.5.5 0 0 1-.63.06l-.07-.06L10 10.71l-5.15 5.14a.5.5 0 0 1-.63.06l-.07-.06a.5.5 0 0 1-.06-.63l.06-.07L9.29 10 4.15 4.85a.5.5 0 0 1-.06-.63l.06-.07-.06.07Z"
            fill="currentColor"
          />
        </svg>
      </fluent-button>
      <slot></slot>
      <slot name="action" slot="action"></slot>
    </fluent-dialog-body>
  </fluent-dialog>
`;

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
@customElement({ name: "docen-dialog", template, styles })
class DocenDialog extends FASTElement {
  @attr heading?: string;
  @attr({ mode: "boolean" }) open?: boolean;

  @observable dialog?: FluentDialog;
  @observable titleEl?: HTMLElement;
  #nativeDialog?: HTMLElement;
  #backdropRaf = 0;
  readonly #toggleHandler = (event: Event): void => {
    if ((event as FluentToggleEvent).detail?.newState === "closed") {
      if (this.open) this.open = false;
    }
  };
  readonly #backdropHandler = (event: Event): void => {
    if (event.target === this.#nativeDialog) event.stopImmediatePropagation();
  };

  headingChanged(): void {
    this.#applyHeading();
  }
  openChanged(): void {
    this.#applyOpen();
  }

  connectedCallback(): void {
    super.connectedCallback();
    // ESC / backdrop close the fluent-dialog directly; sync our `open` attr so
    // state stays consistent. fluent emits `toggle` with newState.
    this.dialog?.addEventListener("toggle", this.#toggleHandler);
    // Office dialogs don't light-dismiss on backdrop click. fluent-dialog's
    // clickHandler hides when the click lands on the native <dialog> itself
    // (the backdrop region); intercept those in capture phase so only ESC, the
    // close button, or Cancel dismisses the dialog.
    this.#disableBackdropDismiss();
    this.#applyHeading();
    this.#applyOpen();
  }

  disconnectedCallback(): void {
    cancelAnimationFrame(this.#backdropRaf);
    this.dialog?.removeEventListener("toggle", this.#toggleHandler);
    this.#nativeDialog?.removeEventListener("click", this.#backdropHandler, true);
    super.disconnectedCallback();
  }

  show(): void {
    this.open = true;
  }

  hide(): void {
    this.open = false;
  }

  #disableBackdropDismiss(): void {
    const apply = (): void => {
      const native = this.dialog?.shadowRoot?.querySelector("dialog");
      if (!native) {
        this.#backdropRaf = requestAnimationFrame(apply);
        return;
      }
      this.#nativeDialog = native;
      native.addEventListener("click", this.#backdropHandler, true);
    };
    apply();
  }

  #applyHeading(): void {
    if (this.titleEl) this.titleEl.textContent = this.heading ?? "";
  }

  // Sync the `open` attribute to fluent-dialog's show/hide. The dialog may not
  // be upgraded on first connect, so retry once it is.
  #applyOpen(): void {
    const dialog = this.dialog;
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

export default DocenDialog;
