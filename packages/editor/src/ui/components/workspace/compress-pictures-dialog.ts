import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(360px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    font-size: 13px;
  }
  fluent-radio-group {
    align-items: flex-start;
    gap: 6px;
  }
`;

const template = html<DocenCompressPicturesDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <fluent-radio-group orientation="vertical" ${ref("targetGroup")}>
        <fluent-field label-position="after">
          <fluent-radio slot="input" value="220"></fluent-radio>
          <label slot="label">${(x) => t("compress.print", x)}</label>
        </fluent-field>
        <fluent-field label-position="after">
          <fluent-radio slot="input" value="150"></fluent-radio>
          <label slot="label">${(x) => t("compress.web", x)}</label>
        </fluent-field>
        <fluent-field label-position="after">
          <fluent-radio slot="input" value="keep"></fluent-radio>
          <label slot="label">${(x) => t("compress.keep", x)}</label>
        </fluent-field>
      </fluent-radio-group>
      <fluent-field label-position="after">
        <fluent-checkbox slot="input" ${ref("dropCropBox")}></fluent-checkbox>
        <label slot="label">${(x) => t("compress.drop-crop", x)}</label>
      </fluent-field>
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
 * `<docen-compress-pictures-dialog>` — Word's Compress Pictures dialog for
 * the selected picture: a target-resolution radio (Print 220 / Web 150 ppi
 * / keep current) plus the delete-cropped-areas option. Commits via
 * `compress:ok` `{ ppi: number | null, dropCrop: boolean }` (`null` = keep
 * the current resolution) or cancels.
 */
@customElement({ name: "docen-compress-pictures-dialog", template, styles })
class DocenCompressPicturesDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable targetGroup?: HTMLElement & { value: string | null };
  @observable dropCropBox?: HTMLElement & { checked: boolean };
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
    if (this.targetGroup) this.targetGroup.value = "220";
    if (this.dropCropBox) this.dropCropBox.checked = false;
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  apply(): void {
    const raw = this.targetGroup?.value ?? "220";
    const ppi = raw === "keep" ? null : Number(raw);
    this.$emit("compress:ok", { ppi, dropCrop: this.dropCropBox?.checked === true });
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("compress.title", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenCompressPicturesDialog;
