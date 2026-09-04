import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { FONT_NAMES } from "../../../document/font-lists";
import type { WatermarkPictureSpec, WatermarkTextSpec } from "../../../document/watermark";
import { observeLang, t } from "../../i18n/localize";

/** The picture-scale ladder (Word's 缩放 list; "auto" fits the box width). */
const SCALES: Array<[string, number | "auto"]> = [
  ["auto", "auto"],
  ["0.5", 0.5],
  ["1", 1],
  ["1.5", 1.5],
  ["2", 2],
];

/** Text sizes offered beside Word's 自动 (points). */
const TEXT_SIZES = [24, 32, 40, 48, 72, 96];

/** auto + the eight standard colors, keyed to the Font dialog's color names. */
const COLORS: Array<[string, string | null]> = [
  ["auto", null],
  ["000000", "colorBlack"],
  ["800000", "colorDarkRed"],
  ["008000", "colorGreen"],
  ["000080", "colorDarkBlue"],
  ["FF0000", "colorRed"],
  ["FF00FF", "colorMagenta"],
  ["FFFF00", "colorYellow"],
  ["00FFFF", "colorCyan"],
];

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(400px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    font-size: 13px;
  }
  .choice {
    display: flex;
    align-items: center;
    gap: 8px;
    cursor: pointer;
  }
  .pane {
    display: flex;
    flex-direction: column;
    gap: 10px;
    padding: 10px 12px;
    border: 1px solid var(--docen-color-divider, #e1e1e1);
    border-radius: 4px;
  }
  .pane[hidden] {
    display: none;
  }
  .row {
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .field {
    display: flex;
    align-items: center;
    gap: 6px;
    flex: 1 1 0;
    min-width: 0;
  }
  .field > label {
    white-space: nowrap;
  }
  select {
    min-width: 0;
    flex: 1 1 auto;
    box-sizing: border-box;
    font: inherit;
    padding: 3px 4px;
  }
  fluent-text-input {
    min-width: 0;
    flex: 1 1 auto;
  }
  .filename {
    flex: 1 1 auto;
    min-width: 0;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    color: var(--docen-color-text-3, #616161);
  }
  .check {
    display: flex;
    align-items: center;
    gap: 6px;
    cursor: pointer;
  }
`;

const template = html<DocenWatermarkDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <label class="choice">
        <input
          type="radio"
          name="wm-kind"
          value="none"
          ${ref("noneRadio")}
          @change="${(x) => x.syncPanes()}"
        />
        <span ${ref("noneLabel")}></span>
      </label>
      <label class="choice">
        <input
          type="radio"
          name="wm-kind"
          value="picture"
          ${ref("pictureRadio")}
          @change="${(x) => x.syncPanes()}"
        />
        <span ${ref("pictureLabel")}></span>
      </label>
      <div class="pane" ${ref("picturePane")} hidden>
        <div class="row">
          <fluent-button ${ref("pickBtn")} @click="${(x) => x.pickImage()}"></fluent-button>
          <span class="filename" ${ref("filename")}></span>
        </div>
        <div class="row">
          <div class="field">
            <label ${ref("scaleLabel")}></label>
            <select ${ref("scaleSel")}></select>
          </div>
          <label class="check">
            <input type="checkbox" ${ref("washoutCheck")} />
            <span ${ref("washoutLabel")}></span>
          </label>
        </div>
      </div>
      <label class="choice">
        <input
          type="radio"
          name="wm-kind"
          value="text"
          ${ref("textRadio")}
          @change="${(x) => x.syncPanes()}"
        />
        <span ${ref("textLabel")}></span>
      </label>
      <div class="pane" ${ref("textPane")} hidden>
        <div class="row">
          <div class="field">
            <label ${ref("textLabel2")}></label>
            <fluent-text-input ${ref("textInput")} value="ASAP"></fluent-text-input>
          </div>
        </div>
        <div class="row">
          <div class="field">
            <label ${ref("fontLabel")}></label>
            <select ${ref("fontSel")}></select>
          </div>
          <div class="field">
            <label ${ref("sizeLabel")}></label>
            <select ${ref("sizeSel")}></select>
          </div>
        </div>
        <div class="row">
          <div class="field">
            <label ${ref("colorLabel")}></label>
            <select ${ref("colorSel")}></select>
          </div>
          <div class="field">
            <label ${ref("layoutLabel")}></label>
            <select ${ref("layoutSel")}>
              <option value="diagonal"></option>
              <option value="horizontal"></option>
            </select>
          </div>
        </div>
        <label class="check">
          <input type="checkbox" ${ref("transparentCheck")} />
          <span ${ref("transparentLabel")}></span>
        </label>
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button appearance="accent" ${ref("okBtn")} @click="${(x) => x.applyWatermark()}">
      </fluent-button>
    </div>
  </docen-dialog>
  <input type="file" accept="image/*" ${ref("fileInput")} hidden />
`;

/**
 * `<docen-watermark-dialog>` — Word's "Custom Watermark" dialog (Design →
 * Watermark → Custom Watermark): no watermark / picture watermark (pick an
 * image, scale, 冲蚀 washout) / text watermark (text, font, size, color,
 * layout, 半透明). OK emits `watermark:ok` with a none/text/picture payload
 * for the host to stamp into the header slots. `show(current)` prefills from
 * the current stamp. Rides on `<docen-dialog>` for the modal shell.
 */
@customElement({ name: "docen-watermark-dialog", template, styles })
class DocenWatermarkDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable noneRadio?: HTMLInputElement;
  @observable noneLabel?: HTMLElement;
  @observable pictureRadio?: HTMLInputElement;
  @observable pictureLabel?: HTMLElement;
  @observable picturePane?: HTMLElement;
  @observable pickBtn?: HTMLElement;
  @observable filename?: HTMLElement;
  @observable scaleLabel?: HTMLElement;
  @observable scaleSel?: HTMLSelectElement;
  @observable washoutCheck?: HTMLInputElement;
  @observable washoutLabel?: HTMLElement;
  @observable textRadio?: HTMLInputElement;
  @observable textLabel?: HTMLElement;
  @observable textPane?: HTMLElement;
  @observable textLabel2?: HTMLElement;
  @observable textInput?: HTMLElement & { value: string };
  @observable fontLabel?: HTMLElement;
  @observable fontSel?: HTMLSelectElement;
  @observable sizeLabel?: HTMLElement;
  @observable sizeSel?: HTMLSelectElement;
  @observable colorLabel?: HTMLElement;
  @observable colorSel?: HTMLSelectElement;
  @observable layoutLabel?: HTMLElement;
  @observable layoutSel?: HTMLSelectElement;
  @observable transparentCheck?: HTMLInputElement;
  @observable transparentLabel?: HTMLElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;
  @observable fileInput?: HTMLInputElement;

  /** The picked image as a data URL (null until 选择图片 picks one). */
  #imageSrc: string | null = null;
  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#fillCombos();
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

  /** Prefill from the current stamp ({kind:"text"|"picture"} fields from the
   *  host's header read); null selects 无水印. */
  show(current?: unknown): void {
    const cur = (current ?? {}) as Record<string, unknown>;
    const kind = cur.kind === "text" || cur.kind === "picture" ? cur.kind : "none";
    if (this.textRadio) this.textRadio.checked = kind === "text";
    if (this.pictureRadio) this.pictureRadio.checked = kind === "picture";
    if (this.noneRadio) this.noneRadio.checked = kind === "none";
    if (kind === "text") {
      if (this.textInput) this.textInput.value = (cur.text as string) || "ASAP";
      if (this.fontSel) this.fontSel.value = (cur.font as string) || "";
      if (this.sizeSel)
        this.sizeSel.value =
          typeof cur.size === "number" && cur.size > 0
            ? String(cur.size)
            : (this.sizeSel.options[0]?.value ?? "auto");
      if (this.colorSel) {
        // A preset's silver (C0C0C0) isn't in the standard ladder — fall back
        // to 自动 rather than a silently-unmatched select.
        const c = cur.color as string;
        this.colorSel.value =
          c && [...this.colorSel.options].some((o) => o.value === c) ? c : "auto";
      }
      if (this.layoutSel) this.layoutSel.value = cur.diagonal ? "diagonal" : "horizontal";
      if (this.transparentCheck) this.transparentCheck.checked = !!cur.semiTransparent;
    } else {
      if (this.scaleSel) this.scaleSel.value = "auto";
      if (this.washoutCheck) this.washoutCheck.checked = !!cur.washout;
      if (this.filename) this.filename.textContent = cur.hasImage ? "…" : "";
    }
    this.#imageSrc = null;
    this.syncPanes();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible pane switcher (radios drive which pane shows). */
  syncPanes(): void {
    if (this.picturePane) this.picturePane.hidden = !this.pictureRadio?.checked;
    if (this.textPane) this.textPane.hidden = !this.textRadio?.checked;
  }

  pickImage(): void {
    this.fileInput?.click();
  }

  readonly #onFileChange = (event: Event): void => {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = "";
    if (!file) return;
    this.#imageSrc = null;
    if (this.filename) this.filename.textContent = file.name;
    const reader = new FileReader();
    reader.onload = (): void => {
      if (typeof reader.result !== "string") return;
      this.#imageSrc = reader.result;
    };
    reader.readAsDataURL(file);
  };

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applyWatermark(): void {
    if (this.textRadio?.checked) {
      const spec: WatermarkTextSpec = {
        text: this.textInput?.value || "ASAP",
        font: this.fontSel?.value || undefined,
        size: this.#sizeValue(),
        color: this.#colorValue(),
        diagonal: (this.layoutSel?.value ?? "diagonal") === "diagonal",
        semiTransparent: !!this.transparentCheck?.checked,
      };
      this.$emit("watermark:ok", { kind: "text", spec });
    } else if (this.pictureRadio?.checked) {
      if (!this.#imageSrc) {
        this.hide();
        this.$emit("watermark:ok", { kind: "none" });
        return;
      }
      const spec: WatermarkPictureSpec = {
        src: this.#imageSrc,
        scale: this.#scaleValue(),
        washout: !!this.washoutCheck?.checked,
      };
      this.$emit("watermark:ok", { kind: "picture", spec });
    } else {
      this.$emit("watermark:ok", { kind: "none" });
    }
    this.hide();
  }

  #sizeValue(): number | "auto" {
    // An unmatched select reads "" (a stale value with no option) — 自动, not 0.
    const raw = this.sizeSel?.value ?? "";
    const n = Number(raw);
    return raw === "auto" || raw === "" || !Number.isFinite(n) ? "auto" : n;
  }

  /** 自动 reads as black to the stamp (the run color is plain hex). */
  #colorValue(): string {
    const raw = this.colorSel?.value ?? "auto";
    return raw === "auto" ? "000000" : raw;
  }

  #scaleValue(): number | "auto" {
    const raw = this.scaleSel?.value ?? "auto";
    const hit = SCALES.find(([key]) => key === raw);
    return raw === "auto" || !hit ? "auto" : (hit[1] as number);
  }

  #fillCombos(): void {
    if (this.fontSel && this.fontSel.options.length === 0) {
      // Word's font list leads with a blank entry — no face means inherit.
      const blank = document.createElement("option");
      blank.value = "";
      this.fontSel.append(blank);
      for (const name of FONT_NAMES) {
        const opt = document.createElement("option");
        opt.value = name;
        opt.textContent = name;
        this.fontSel.append(opt);
      }
    }
    if (this.sizeSel && this.sizeSel.options.length === 0) {
      const auto = document.createElement("option");
      auto.value = "auto";
      this.sizeSel.append(auto);
      for (const size of TEXT_SIZES) {
        const opt = document.createElement("option");
        opt.value = String(size);
        opt.textContent = String(size);
        this.sizeSel.append(opt);
      }
    }
    if (this.scaleSel && this.scaleSel.options.length === 0) {
      for (const [key, value] of SCALES) {
        const opt = document.createElement("option");
        opt.value = key;
        opt.textContent = value === "auto" ? "" : `${Math.round((value as number) * 100)}%`;
        this.scaleSel.append(opt);
      }
    }
    if (this.colorSel && this.colorSel.options.length === 0) {
      for (const [hex] of COLORS) {
        const opt = document.createElement("option");
        opt.value = hex;
        this.colorSel.append(opt);
      }
    }
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("watermarkDialog.title", this);
    if (this.noneLabel) this.noneLabel.textContent = t("watermarkDialog.none", this);
    if (this.pictureLabel) this.pictureLabel.textContent = t("watermarkDialog.picture", this);
    if (this.pickBtn) this.pickBtn.textContent = t("watermarkDialog.selectImage", this);
    if (this.scaleLabel) this.scaleLabel.textContent = t("watermarkDialog.scale", this);
    if (this.washoutLabel) this.washoutLabel.textContent = t("watermarkDialog.washout", this);
    if (this.textLabel) this.textLabel.textContent = t("watermarkDialog.text", this);
    if (this.textLabel2) this.textLabel2.textContent = t("watermarkDialog.textLabel", this);
    if (this.fontLabel) this.fontLabel.textContent = t("watermarkDialog.font", this);
    if (this.sizeLabel) this.sizeLabel.textContent = t("watermarkDialog.size", this);
    if (this.colorLabel) this.colorLabel.textContent = t("watermarkDialog.color", this);
    if (this.layoutLabel) this.layoutLabel.textContent = t("watermarkDialog.layout", this);
    if (this.transparentLabel)
      this.transparentLabel.textContent = t("watermarkDialog.semitransparent", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    if (this.scaleSel && this.scaleSel.options[0])
      this.scaleSel.options[0].textContent = t("watermarkDialog.auto", this);
    if (this.sizeSel && this.sizeSel.options[0])
      this.sizeSel.options[0].textContent = t("watermarkDialog.auto", this);
    if (this.colorSel)
      COLORS.forEach(([, key], i) => {
        const opt = this.colorSel?.options[i];
        if (opt && key) opt.textContent = t(`fontDialog.${key}`, this);
      });
    if (this.layoutSel) {
      this.layoutSel.options[0].textContent = t("watermarkDialog.diagonal", this);
      this.layoutSel.options[1].textContent = t("watermarkDialog.horizontal", this);
    }
  }
}

export default DocenWatermarkDialog;
