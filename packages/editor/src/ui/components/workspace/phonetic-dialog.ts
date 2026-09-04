import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** The alignment choices (ST_RubyAlign's horizontal tokens — rightVertical is
 *  a vertical-text form the dialog does not offer). */
const ALIGNMENTS = ["center", "distributeLetter", "distributeSpace", "left", "right"] as const;

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(380px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .rows {
    max-height: 240px;
    overflow: auto;
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
  .row {
    display: grid;
    grid-template-columns: 32px 1fr;
    align-items: center;
    gap: 8px;
  }
  .row .base {
    text-align: center;
    font-size: 15px;
    border: 1px solid var(--neutral-stroke-rest, #d1d1d1);
    border-radius: 3px;
    padding: 3px 0;
    user-select: none;
  }
  .meta {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .meta select {
    flex: 1;
  }
  .clear {
    align-self: flex-end;
  }
`;

const template = html<DocenPhoneticDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="rows" ${ref("rowsEl")}></div>
      <div class="meta">
        <span>${(x) => t("phoneticDialog.alignment", x)}</span>
        <select ${ref("alignSel")}></select>
      </div>
      <fluent-button
        class="clear"
        ${ref("clearBtn")}
        @click="${(x) => x.clearAll()}"
      ></fluent-button>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyGuide()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-phonetic-dialog>` — Word's Phonetic Guide dialog (拼音指南, Home →
 * Font): per-character readings over the selection, an alignment pick, and a
 * clear-all. Opened with `show(chars, readings, alignment)` — one row per
 * base character, the existing readings prefilled; commits via `phonetic:ok`
 * `{ chars, readings, alignment }` (a blank reading leaves that character
 * unannotated) or `phonetic:clear` (remove the guides from the selection).
 */
@customElement({ name: "docen-phonetic-dialog", template, styles })
class DocenPhoneticDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable rowsEl?: HTMLDivElement;
  @observable alignSel?: HTMLSelectElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;
  @observable clearBtn?: HTMLElement;
  chars: string[] = [];

  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#unobserveLang = observeLang(() => this.#applyLabels());
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  /** Open with one row per base character — plain DOM rows built here (the
   *  readings are user-typed; a template repeat would fight the typing). */
  show(chars: string[], readings: string[], alignment: string | null): void {
    this.chars = [...chars];
    if (this.rowsEl) {
      this.rowsEl.replaceChildren(
        ...chars.map((ch, i) => {
          const row = document.createElement("div");
          row.className = "row";
          const base = document.createElement("div");
          base.className = "base";
          base.textContent = ch;
          const input = document.createElement("input");
          input.type = "text";
          input.value = readings[i] ?? "";
          input.spellcheck = false;
          input.style.width = "100%";
          input.style.boxSizing = "border-box";
          row.append(base, input);
          return row;
        }),
      );
    }
    if (this.alignSel)
      this.alignSel.value = (ALIGNMENTS as readonly string[]).includes(alignment ?? "")
        ? alignment!
        : "center";
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applyGuide(): void {
    const readings = [...(this.rowsEl?.querySelectorAll<HTMLInputElement>("input") ?? [])].map(
      (input) => input.value,
    );
    this.$emit("phonetic:ok", {
      chars: this.chars,
      readings,
      alignment: this.alignSel?.value ?? "center",
    });
    this.hide();
  }

  clearAll(): void {
    this.$emit("phonetic:clear", {});
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("phoneticDialog.title", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    if (this.clearBtn) this.clearBtn.textContent = t("phoneticDialog.clearAll", this);
    // The alignment options are built here (like font-dialog's selects) —
    // a `t()` inside a template repeat binds with the item, not the element,
    // and aborts the whole render.
    if (this.alignSel && this.alignSel.options.length === 0)
      this.alignSel.replaceChildren(
        ...ALIGNMENTS.map((value) => new Option(t(`phoneticDialog.align.${value}`, this), value)),
      );
  }
}

export default DocenPhoneticDialog;
