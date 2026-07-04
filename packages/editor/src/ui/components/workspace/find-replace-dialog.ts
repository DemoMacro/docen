import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

const styles = css`
  :host {
    display: contents;
  }
  .field {
    display: flex;
    align-items: center;
    gap: 8px;
    margin-block-end: 10px;
  }
  .field > label {
    width: 96px;
    font-size: 13px;
    color: #3b3b3b;
  }
  fluent-text-input {
    flex: 1 1 auto;
    min-width: 0;
  }
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
`;

const template = html<DocenFindReplaceDialog>`
  <docen-dialog heading="Find and Replace" part="dialog" ${ref("dialog")}>
    <div class="field">
      <label data-i18n="findReplace.find">Find what:</label>
      <fluent-text-input part="find" ${ref("find")} type="text"></fluent-text-input>
    </div>
    <div class="field">
      <label data-i18n="findReplace.replaceWith">Replace with:</label>
      <fluent-text-input part="replace" ${ref("replace")} type="text"></fluent-text-input>
    </div>
    <div class="options">
      <label class="check-field">
        <fluent-checkbox part="case" ${ref("case")}></fluent-checkbox>
        <span data-i18n="findReplace.matchCase">Match case</span>
      </label>
      <label class="check-field">
        <fluent-checkbox part="word" ${ref("word")}></fluent-checkbox>
        <span data-i18n="findReplace.wholeWord">Whole word</span>
      </label>
    </div>
    <fluent-button
      slot="action"
      part="find-next"
      appearance="stealth"
      data-i18n="findReplace.findNext"
      >Find Next</fluent-button
    >
    <fluent-button
      slot="action"
      part="replace-next"
      appearance="stealth"
      data-i18n="findReplace.replace"
      >Replace</fluent-button
    >
    <fluent-button
      slot="action"
      part="replace-all"
      appearance="accent"
      data-i18n="findReplace.replaceAll"
      >Replace All</fluent-button
    >
    <fluent-button slot="action" part="cancel" appearance="stealth" data-i18n="findReplace.cancel"
      >Cancel</fluent-button
    >
  </docen-dialog>
`;

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
@customElement({ name: "docen-find-replace-dialog", template, styles })
class DocenFindReplaceDialog extends FASTElement {
  @attr({ mode: "boolean" }) open?: boolean;

  @observable dialog?: DialogEl;
  @observable find?: InputEl;
  @observable replace?: InputEl;
  @observable case?: CheckEl;
  @observable word?: CheckEl;
  #dialogObserver?: MutationObserver;
  #unsubscribe?: () => void;

  openChanged(): void {
    this.#applyOpen();
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.find?.addEventListener("input", () => this.#emit("query"));
    this.replace?.addEventListener("input", () => this.#emit("query"));
    const root = this.shadowRoot!;
    root
      .querySelector("[part='find-next']")
      ?.addEventListener("click", () => this.#emit("find-next"));
    root
      .querySelector("[part='replace-next']")
      ?.addEventListener("click", () => this.#emit("replace-next"));
    root
      .querySelector("[part='replace-all']")
      ?.addEventListener("click", () => this.#emit("replace-all"));
    root.querySelector("[part='cancel']")?.addEventListener("click", () => this.hide());
    // ESC / ✕ inside the inner docen-dialog drop its `open` attr — mirror that
    // here so our state stays consistent.
    this.#dialogObserver = new MutationObserver(() => {
      if (this.dialog && !this.dialog.hasAttribute("open") && this.open) {
        this.open = false;
      }
    });
    if (this.dialog)
      this.#dialogObserver.observe(this.dialog, { attributes: true, attributeFilter: ["open"] });
    this.#applyOpen();
    this.#applyI18n();
    this.#unsubscribe = observeLang(() => this.#applyI18n());
  }

  disconnectedCallback(): void {
    this.#dialogObserver?.disconnect();
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  show(): void {
    this.open = true;
  }

  hide(): void {
    this.open = false;
  }

  #emit(action: string): void {
    this.dispatchEvent(
      new CustomEvent("find-replace:action", {
        bubbles: true,
        composed: true,
        detail: { action, source: this, ...this.#read() },
      }),
    );
  }

  #read(): Record<string, unknown> {
    return {
      find: this.find?.value ?? "",
      replace: this.replace?.value ?? "",
      caseSensitive: !!this.case?.checked,
      wholeWord: !!this.word?.checked,
    };
  }

  #applyOpen(): void {
    if (this.dialog) this.dialog.open = !!this.open;
    if (this.open) requestAnimationFrame(() => this.#focusFind());
  }

  #focusFind(): void {
    this.find?.focus();
    this.find?.select?.();
  }

  #applyI18n(): void {
    const root = this.shadowRoot;
    if (!root) return;
    this.dialog?.setAttribute("heading", t("findReplace.title", this));
    root.querySelectorAll<HTMLElement>("[data-i18n]").forEach((el) => {
      el.textContent = t(el.dataset.i18n!, this);
    });
  }
}

export default DocenFindReplaceDialog;
