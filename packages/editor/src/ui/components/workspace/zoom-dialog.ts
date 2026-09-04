import {
  FASTElement,
  css,
  customElement,
  html,
  observable,
  ref,
  repeat,
} from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** A `fluent-text-input` widget plus its string value accessor (the value
 *  lives on the `value` property, like a native input). */
type FluentTextInput = HTMLElement & { value: string; disabled: boolean };

/** A named zoom preset — numeric strings are percents; the rest resolve from
 *  the stage geometry on the host (#zoomPreset). */
export type ZoomChoice = "200" | "100" | "75" | "page-width" | "text-width" | "fit-page" | number;

const PRESETS: Array<{ value: string; key: string }> = [
  { value: "200", key: "ribbon.opt.200" },
  { value: "100", key: "ribbon.opt.100" },
  { value: "75", key: "ribbon.opt.75" },
  { value: "page-width", key: "ribbon.opt.page-width" },
  { value: "text-width", key: "ribbon.opt.text-width" },
  { value: "fit-page", key: "ribbon.opt.fit-page" },
];

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(300px, 92vw);
  }
  .zoom-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  fluent-radio-group {
    align-items: flex-start;
    gap: 6px;
  }
  .percent-row {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  fluent-text-input {
    width: 90px;
  }
`;

const presetTemplate = html<(typeof PRESETS)[number], DocenZoomDialog>`<fluent-field
  label-position="after"
>
  <fluent-radio slot="input" value="${(p) => p.value}"></fluent-radio>
  <label slot="label"><span class="preset-label" data-key="${(p) => p.key}"></span></label>
</fluent-field>`;

const template = html<DocenZoomDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <!-- Radio changes are delegated here — change bubbles from every preset +
         percent radio through the group. -->
    <div class="zoom-body" @change="${(x) => x.syncPercentInput()}">
      <fluent-radio-group orientation="vertical" ${ref("presetGroup")}>
        ${repeat(() => PRESETS, presetTemplate)}
        <fluent-field label-position="after">
          <fluent-radio slot="input" value="percent" ${ref("percentRadio")}></fluent-radio>
          <label slot="label" class="percent-row">
            <span ${ref("percentLabel")}></span>
            <fluent-text-input
              ${ref("percentInput")}
              type="number"
              min="10"
              max="500"
              disabled
            ></fluent-text-input>
            <span>%</span>
          </label>
        </fluent-field>
      </fluent-radio-group>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyZoom()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-zoom-dialog>` — Word's Zoom dialog (View → Zoom, and the status-bar
 * percent click): preset radios (200/100/75/page width/text width/one page)
 * plus a free percent box (10–500). Opened with `show(zoom)` prefilled from
 * the current zoom; commits via `zoom:ok` with a preset name or a number.
 * Multi-page grids stay out until the stage lays pages out in grids.
 */
@customElement({ name: "docen-zoom-dialog", template, styles })
class DocenZoomDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  // The group's `value` mirrors the checked radio's `value` (a preset name or
  // "percent"); the percent radio + label + box ride inside the group.
  @observable presetGroup?: HTMLElement & { value: string | null };
  @observable percentRadio?: HTMLElement & { checked: boolean };
  @observable percentLabel?: HTMLElement;
  @observable percentInput?: FluentTextInput;
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

  /** Open prefilled from the current zoom: the matching preset radio when one
   *  matches, the percent radio + box otherwise. */
  show(zoom: number): void {
    const matched = PRESETS.some((p) => p.value === String(zoom));
    if (this.presetGroup) this.presetGroup.value = matched ? String(zoom) : "percent";
    if (this.percentInput) {
      // The percent box is live only while the percent radio is picked —
      // a preset match disables it.
      this.percentInput.disabled = matched;
      this.percentInput.value = String(zoom);
    }
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** The percent box is only live while its radio is picked. */
  syncPercentInput(): void {
    if (this.percentInput) this.percentInput.disabled = this.presetGroup?.value !== "percent";
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applyZoom(): void {
    const picked = this.presetGroup?.value ?? null;
    let value: ZoomChoice = 100;
    if (picked === "percent") {
      const n = Number(this.percentInput?.value);
      value = Number.isFinite(n) ? Math.max(10, Math.min(500, Math.round(n))) : 100;
    } else if (picked) {
      value = picked as ZoomChoice;
    }
    this.$emit("zoom:ok", value);
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("zoomDialog.title", this);
    for (const el of this.shadowRoot?.querySelectorAll<HTMLElement>(".preset-label") ?? [])
      el.textContent = t(el.dataset.key ?? "", this);
    if (this.percentLabel) this.percentLabel.textContent = t("zoomDialog.percent", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenZoomDialog;
