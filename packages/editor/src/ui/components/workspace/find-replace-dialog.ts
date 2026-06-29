import { observeLang, t } from "../../i18n/localize";

const template = document.createElement("template");
template.innerHTML = `
  <style>
    :host { display: contents; }
    .field {
      display: flex;
      align-items: center;
      gap: 8px;
      margin-block-end: 10px;
    }
    .field > label { width: 96px; font-size: 13px; color: #3b3b3b; }
    fluent-text-input { flex: 1 1 auto; min-width: 0; }
    .options {
      display: flex;
      gap: 16px;
      margin-block-end: 8px;
      font-size: 13px;
    }
    .check-field {
      display: flex;
      align-items: center;
      gap: 6px;
      cursor: pointer;
    }
  </style>
  <docen-dialog heading="Find and Replace" part="dialog">
    <div class="field">
      <label data-i18n="findReplace.find">Find what:</label>
      <fluent-text-input part="find" type="text"></fluent-text-input>
    </div>
    <div class="field">
      <label data-i18n="findReplace.replaceWith">Replace with:</label>
      <fluent-text-input part="replace" type="text"></fluent-text-input>
    </div>
    <div class="options">
      <label class="check-field">
        <fluent-checkbox part="case"></fluent-checkbox>
        <span data-i18n="findReplace.matchCase">Match case</span>
      </label>
      <label class="check-field">
        <fluent-checkbox part="word"></fluent-checkbox>
        <span data-i18n="findReplace.wholeWord">Whole word</span>
      </label>
    </div>
    <fluent-button slot="action" part="find-next" appearance="stealth" data-i18n="findReplace.findNext">Find Next</fluent-button>
    <fluent-button slot="action" part="replace-next" appearance="stealth" data-i18n="findReplace.replace">Replace</fluent-button>
    <fluent-button slot="action" part="replace-all" appearance="accent" data-i18n="findReplace.replaceAll">Replace All</fluent-button>
    <fluent-button slot="action" part="cancel" appearance="stealth" data-i18n="findReplace.cancel">Cancel</fluent-button>
  </docen-dialog>`;

interface DialogEl extends HTMLElement {
  open: boolean;
}
interface InputEl extends HTMLElement {
  value: string;
  focus(): void;
  select(): void;
}
interface CheckEl extends HTMLElement {
  checked: boolean;
}

/**
 * `<docen-find-replace-dialog open>` — a Word-style Find & Replace modal built
 * on `<docen-dialog>`. Two text fields (Find what / Replace with), Match case /
 * Whole word checkboxes, and Find Next / Replace / Replace All / Cancel
 * actions. Every input change or action emits `find-replace:action` with the
 * current fields; the host (`<docen-document>`) drives prosemirror-search.
 */
class DocenFindReplaceDialog extends HTMLElement {
  static get observedAttributes(): string[] {
    return ["open"];
  }

  #dialog?: DialogEl;
  #dialogObserver?: MutationObserver;
  #unsubscribe?: () => void;

  attributeChangedCallback(name: string): void {
    if (name === "open") this.#applyOpen();
  }

  connectedCallback(): void {
    if (!this.shadowRoot) {
      this.attachShadow({ mode: "open" }).append(template.content.cloneNode(true));
    }
    const root = this.shadowRoot!;
    this.#dialog = root.querySelector<DialogEl>("[part='dialog']") ?? undefined;
    const find = root.querySelector<InputEl>("[part='find']")!;
    const replace = root.querySelector<InputEl>("[part='replace']")!;
    const caseCb = root.querySelector<CheckEl>("[part='case']")!;
    const wordCb = root.querySelector<CheckEl>("[part='word']")!;
    const read = (): Record<string, unknown> => ({
      find: find.value ?? "",
      replace: replace.value ?? "",
      caseSensitive: !!caseCb.checked,
      wholeWord: !!wordCb.checked,
    });
    const emit = (action: string): void => {
      this.dispatchEvent(
        new CustomEvent("find-replace:action", {
          bubbles: true,
          composed: true,
          detail: { action, source: this, ...read() },
        }),
      );
    };
    find.addEventListener("input", () => emit("query"));
    replace.addEventListener("input", () => emit("query"));
    root.querySelector("[part='find-next']")?.addEventListener("click", () => emit("find-next"));
    root
      .querySelector("[part='replace-next']")
      ?.addEventListener("click", () => emit("replace-next"));
    root
      .querySelector("[part='replace-all']")
      ?.addEventListener("click", () => emit("replace-all"));
    root.querySelector("[part='cancel']")?.addEventListener("click", () => this.hide());
    // ESC / ✕ inside the inner docen-dialog drop its `open` attr — mirror that
    // here so our state stays consistent.
    this.#dialogObserver = new MutationObserver(() => {
      if (this.#dialog && !this.#dialog.hasAttribute("open") && this.hasAttribute("open")) {
        this.removeAttribute("open");
      }
    });
    this.#dialogObserver.observe(this.#dialog!, { attributes: true, attributeFilter: ["open"] });
    this.#applyOpen();
    this.#applyI18n();
    this.#unsubscribe = observeLang(() => this.#applyI18n());
  }

  disconnectedCallback(): void {
    this.#dialogObserver?.disconnect();
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

  #applyOpen(): void {
    if (this.#dialog) this.#dialog.open = this.open;
    if (this.open) requestAnimationFrame(() => this.#focusFind());
  }

  #focusFind(): void {
    const find = this.shadowRoot?.querySelector<InputEl>("[part='find']");
    find?.focus();
    find?.select?.();
  }

  #applyI18n(): void {
    const root = this.shadowRoot!;
    this.#dialog?.setAttribute("heading", t("findReplace.title", this));
    root.querySelectorAll<HTMLElement>("[data-i18n]").forEach((el) => {
      el.textContent = t(el.dataset.i18n!, this);
    });
  }
}

customElements.define("docen-find-replace-dialog", DocenFindReplaceDialog);

export default DocenFindReplaceDialog;
