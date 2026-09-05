import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(360px, 92vw);
  }
  .fe-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
  }
  .fe-row {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .fe-hint {
    font-size: 13px;
    color: var(--colorNeutralForeground2, #444);
  }
  .fe-preview {
    width: 100%;
    height: 120px;
    object-fit: cover;
    border: 1px solid var(--colorNeutralStroke2, #e0e0e0);
    border-radius: 4px;
    background: #fff;
  }
  .fe-remove {
    align-self: flex-start;
  }
`;

const template = html<DocenFillEffectsDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="fe-body">
      <span class="fe-hint" ${ref("hintEl")}></span>
      <div class="fe-row">
        <fluent-button appearance="outline" ${ref("pickBtn")} @click="${(x) => x.pickImage()}">
        </fluent-button>
        <fluent-button
          appearance="outline"
          class="fe-remove"
          ${ref("removeBtn")}
          @click="${(x) => x.removeImage()}"
        ></fluent-button>
      </div>
      <img class="fe-preview" ${ref("previewEl")} hidden alt="" />
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button appearance="accent" ${ref("okBtn")} @click="${(x) => x.submit()}">
      </fluent-button>
    </div>
  </docen-dialog>
  <input type="file" accept="image/*" ${ref("fileInput")} hidden />
`;

/**
 * `<docen-fill-effects-dialog>` — Word's "Fill Effects" page background
 * (Design → Page Color group): pick a picture that fills the page (Word's
 * Fill Effects → Picture pane; gradients/textures/patterns have no
 * structured write path and are not offered). OK emits `fill-effects:ok`
 * with `{ image: { data, type } }`, or `{}` when the fill was removed —
 * the host stamps `background.image` either way. Cancel changes nothing.
 */
@customElement({ name: "docen-fill-effects-dialog", template, styles })
class DocenFillEffectsDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable hintEl?: HTMLElement;
  @observable pickBtn?: HTMLElement;
  @observable removeBtn?: HTMLElement;
  @observable previewEl?: HTMLImageElement;
  @observable cancelBtn?: HTMLElement;
  @observable okBtn?: HTMLElement;
  @observable fileInput?: HTMLInputElement;

  // The picked/current fill as a data URL; null = no fill staged.
  #imageSrc: string | null = null;
  #imageType: string | null = null;
  #removed = false;
  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#unobserveLang = observeLang(() => this.#applyLabels());
    this.fileInput?.addEventListener("change", this.#onFileChange);
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    this.fileInput?.removeEventListener("change", this.#onFileChange);
    super.disconnectedCallback();
  }

  /** Prefill from the current `background.image` ({ data, type }); null
   *  stages no fill. */
  show(current?: unknown): void {
    const cur = (current ?? {}) as { data?: unknown; type?: unknown };
    this.#imageSrc = typeof cur.data === "string" ? cur.data : null;
    this.#imageType = typeof cur.type === "string" ? cur.type : null;
    this.#removed = false;
    this.#syncPreview();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  pickImage(): void {
    this.fileInput?.click();
  }

  removeImage(): void {
    this.#imageSrc = null;
    this.#imageType = null;
    this.#removed = true;
    this.#syncPreview();
  }

  /** OK — emit the staged fill (or the removal) for the host to stamp. */
  submit(): void {
    this.dialogEl?.hide();
    if (this.#removed || !this.#imageSrc) {
      this.$emit("fill-effects:ok", {});
      return;
    }
    this.$emit("fill-effects:ok", {
      image: { data: this.#imageSrc, type: this.#imageType ?? "png" },
    });
  }

  #syncPreview(): void {
    if (!this.previewEl) return;
    this.previewEl.hidden = !this.#imageSrc;
    this.previewEl.src = this.#imageSrc ?? "";
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("fillEffects.title", this);
    if (this.hintEl) this.hintEl.textContent = t("fillEffects.hint", this);
    if (this.pickBtn) this.pickBtn.textContent = t("fillEffects.pick", this);
    if (this.removeBtn) this.removeBtn.textContent = t("fillEffects.remove", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("fillEffects.cancel", this);
    if (this.okBtn) this.okBtn.textContent = t("fillEffects.ok", this);
  }

  readonly #onFileChange = (event: Event): void => {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = "";
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (): void => {
      if (typeof reader.result !== "string") return;
      this.#imageSrc = reader.result;
      // "data:image/png;base64,…" → "png" (the subtype, without parameters).
      const mime = /^data:image\/([a-z0-9+.-]+)[;,]/i.exec(reader.result)?.[1];
      this.#imageType = mime === "jpeg" ? "jpg" : (mime ?? null);
      this.#removed = false;
      this.#syncPreview();
    };
    reader.readAsDataURL(file);
  };
}

export default DocenFillEffectsDialog;
