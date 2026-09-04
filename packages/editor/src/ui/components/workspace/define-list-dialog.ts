import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** The number styles the dialog offers (w:numFmt values the renderer draws)
 *  with the sample shapes the option labels show. */
const NUMBER_FORMATS: { value: string; sample: string; suffix: string }[] = [
  { value: "decimal", sample: "1, 2, 3, …", suffix: "." },
  { value: "lowerLetter", sample: "a, b, c, …", suffix: "." },
  { value: "upperLetter", sample: "A, B, C, …", suffix: "." },
  { value: "lowerRoman", sample: "i, ii, iii, …", suffix: "." },
  { value: "upperRoman", sample: "I, II, III, …", suffix: "." },
  { value: "chineseCounting", sample: "一, 二, 三, …", suffix: "、" },
  { value: "chineseCountingThousand", sample: "壹, 贰, 叁, …", suffix: "、" },
];

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(430px, 94vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    font-size: 13px;
  }
  .level {
    display: flex;
    flex-direction: column;
    gap: 6px;
    border: 1px solid var(--neutral-stroke-rest, #d1d1d1);
    border-radius: 4px;
    padding: 8px 10px;
  }
  .level > .title {
    font-weight: 600;
  }
  .row {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .row > label {
    min-width: 92px;
  }
  .row select,
  fluent-text-input {
    flex: 1;
    min-width: 0;
  }
  .note {
    opacity: 0.75;
  }
`;

// Three explicit level blocks — FAST stringifies a `.map()` of templates
// interpolated into an html`` binding ([object Object] × 3).
const template = html<DocenDefineListDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="level">
        <span class="title">${(x) => `${t("defineList.level", x)} 1`}</span>
        <div class="row">
          <label>${(x) => t("defineList.format", x)}</label>
          <select ${ref("l0Format")}></select>
        </div>
        <div class="row">
          <label>${(x) => t("defineList.text", x)}</label>
          <fluent-text-input ${ref("l0Text")} spellcheck="false"></fluent-text-input>
        </div>
      </div>
      <div class="level">
        <span class="title">${(x) => `${t("defineList.level", x)} 2`}</span>
        <div class="row">
          <label>${(x) => t("defineList.format", x)}</label>
          <select ${ref("l1Format")}></select>
        </div>
        <div class="row">
          <label>${(x) => t("defineList.text", x)}</label>
          <fluent-text-input ${ref("l1Text")} spellcheck="false"></fluent-text-input>
        </div>
      </div>
      <div class="level">
        <span class="title">${(x) => `${t("defineList.level", x)} 3`}</span>
        <div class="row">
          <label>${(x) => t("defineList.format", x)}</label>
          <select ${ref("l2Format")}></select>
        </div>
        <div class="row">
          <label>${(x) => t("defineList.text", x)}</label>
          <fluent-text-input ${ref("l2Text")} spellcheck="false"></fluent-text-input>
        </div>
      </div>
      <span class="note">${(x) => t("defineList.note", x)}</span>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.apply()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-define-list-dialog>` — Word's Define New Multilevel List dialog
 * (定义新多级列表, the List Library's last entry): per-level number style and
 * the marker text for the first three levels (deeper levels extend the
 * third's shape). Opened with `show()`; commits via `define-list:ok`
 * `{ levels: [{ format, text } × 3] }` or cancels.
 */
@customElement({ name: "docen-define-list-dialog", template, styles })
class DocenDefineListDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable l0Format?: HTMLSelectElement;
  @observable l1Format?: HTMLSelectElement;
  @observable l2Format?: HTMLSelectElement;
  // fluent-text-input exposes the string on `value`, like a native input.
  @observable l0Text?: HTMLElement & { value: string };
  @observable l1Text?: HTMLElement & { value: string };
  @observable l2Text?: HTMLElement & { value: string };
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

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

  show(): void {
    // Reset to the cascading decimal starter (Word's dialog opens on the
    // current list's shapes — the editor has no current-list readback, so a
    // fresh cascade it is).
    const reset = (
      format: HTMLSelectElement | undefined,
      text: { value: string } | undefined,
      n: number,
    ) => {
      if (!format || !text) return;
      format.value = "decimal";
      text.value = `%${n}.`;
    };
    reset(this.l0Format, this.l0Text, 1);
    reset(this.l1Format, this.l1Text, 2);
    reset(this.l2Format, this.l2Text, 3);
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  apply(): void {
    const levels = [
      [this.l0Format, this.l0Text],
      [this.l1Format, this.l1Text],
      [this.l2Format, this.l2Text],
    ]
      .map(([format, text]) => ({
        format: (format as HTMLSelectElement)?.value,
        text: text?.value ?? "",
      }))
      .filter((lvl) => lvl.text);
    if (levels.length === 0) return;
    this.$emit("define-list:ok", { levels });
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("defineList.title", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    // The selects are built here (like font-dialog's) — a `t()` inside a
    // template repeat binds with the item, not the element, and aborts the
    // whole render.
    for (const select of [this.l0Format, this.l1Format, this.l2Format]) {
      if (select && select.options.length === 0)
        select.replaceChildren(...NUMBER_FORMATS.map((f) => new Option(f.sample, f.value)));
      if (select)
        select.onchange = () => {
          // Picking a style pre-fills this level's marker text with the
          // style's sample shape (%n + suffix); the text stays editable.
          const input = [this.l0Text, this.l1Text, this.l2Text][
            [this.l0Format, this.l1Format, this.l2Format].indexOf(select)
          ];
          if (!select || !input) return;
          const n = [this.l0Format, this.l1Format, this.l2Format].indexOf(select) + 1;
          const suffix = NUMBER_FORMATS.find((f) => f.value === select.value)?.suffix ?? ".";
          input.value = `%${n}${suffix}`;
        };
    }
  }
}

export default DocenDefineListDialog;
