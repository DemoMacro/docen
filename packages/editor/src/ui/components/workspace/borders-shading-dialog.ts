import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import type { BorderSideState, BordersDialogPatch } from "../../../document/extensions/commands";
import { observeLang, t } from "../../i18n/localize";
import { listboxOf, opt, pick, pickedValue, type FluentDropdown } from "./fluent-combo";

/** OOXML border `size` is eighths of a point; the picker lists Word's standard
 *  pt ladder. */
const WIDTHS: Array<[number, string]> = [
  [2, "0.25"],
  [4, "0.5"],
  [6, "0.75"],
  [12, "1.5"],
  [24, "3"],
  [36, "4.5"],
];

/** ST_Border tokens the picker exposes (Word's common run, spelled-out). */
const LINE_STYLES = [
  "single",
  "double",
  "triple",
  "dashed",
  "dashSmallGap",
  "dotted",
  "dotDash",
  "wave",
  "thick",
] as const;

/** The palette row: Word's standard colors (auto = the text ink) — the keys
 *  are the `fontDialog.color*` i18n suffixes. */
const COLORS: Array<[string | null, string]> = [
  [null, "colorAuto"],
  ["000000", "colorBlack"],
  ["800000", "colorDarkRed"],
  ["008000", "colorGreen"],
  ["000080", "colorDarkBlue"],
  ["FF0000", "colorRed"],
  ["FF00FF", "colorMagenta"],
  ["FFFF00", "colorYellow"],
  ["00FFFF", "colorCyan"],
];

/** One tab's full style state — the two border tabs keep independent sides
 *  (each prefills from its own source) but share the widget layout. */
interface TabState {
  style: string;
  color: string | null;
  width: number;
  sides: Record<"top" | "bottom" | "left" | "right", BorderSideState | null>;
}

const emptySides = (): TabState["sides"] => ({ top: null, bottom: null, left: null, right: null });

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
    gap: 10px;
    font-size: 13px;
  }
  .tabs {
    font-size: 13px;
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
  fluent-dropdown {
    min-width: 0;
    flex: 1 1 auto;
  }
  fluent-dropdown input {
    width: 100%;
    box-sizing: border-box;
  }
  .presets {
    display: flex;
    gap: 6px;
  }
  .presets button {
    flex: 1;
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 4px;
    font: inherit;
    font-size: 11px;
    padding: 6px 2px;
    background: none;
    border: 1px solid var(--docen-color-divider, #e1e1e1);
    border-radius: 4px;
    cursor: pointer;
    color: inherit;
  }
  .presets button[aria-pressed="true"] {
    border-color: var(--docen-color-brand, #0078d4);
    background: color-mix(in srgb, var(--docen-color-brand, #0078d4) 8%, transparent);
  }
  .presets .box {
    width: 28px;
    height: 20px;
    border: 1px solid #6e6e6e;
  }
  .presets .box.shadow {
    border-bottom-width: 3px;
    border-right-width: 3px;
  }
  .presets .box.none {
    border-style: dotted;
  }
  .preview {
    display: grid;
    grid-template-columns: 16px 1fr 16px;
    grid-template-rows: 16px 1fr 16px;
    height: 90px;
    background: var(--docen-color-hover, rgba(0, 0, 0, 0.02));
  }
  .preview .edge {
    cursor: pointer;
    background: none;
    border: 0 solid #555;
    padding: 0;
  }
  .palette {
    display: flex;
    flex-wrap: wrap;
    gap: 4px;
  }
  .palette button {
    width: 24px;
    height: 24px;
    border: 1px solid var(--docen-color-divider, #e1e1e1);
    border-radius: 3px;
    cursor: pointer;
    padding: 0;
    font: inherit;
    font-size: 9px;
  }
  .palette button[aria-pressed="true"] {
    outline: 2px solid var(--docen-color-brand, #0078d4);
    outline-offset: 1px;
  }
  .heading {
    font-weight: 600;
  }
  .hint {
    margin: 0;
    font-size: 12px;
    color: var(--docen-color-text-3, #8a8a8a);
  }
  .hidden {
    display: none;
  }
`;

const template = html<DocenBordersShadingDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <fluent-tablist class="tabs" ${ref("tablist")} @change="${(x) => x.onTabChange()}">
        <fluent-tab id="bs-tab-border" ${ref("borderTabBtn")}></fluent-tab>
        <fluent-tab id="bs-tab-page" ${ref("pageTabBtn")}></fluent-tab>
        <fluent-tab id="bs-tab-shading" ${ref("shadingTabBtn")}></fluent-tab>
      </fluent-tablist>
      <div ${ref("borderPage")}>
        <div class="heading" ${ref("settingHeading")}></div>
        <div class="presets">
          <button @click="${(x) => x.applyPreset("none")}" ${ref("presetNone")}>
            <span class="box none"></span><span ${ref("presetNoneLabel")}></span>
          </button>
          <button @click="${(x) => x.applyPreset("box")}" ${ref("presetBox")}>
            <span class="box"></span><span ${ref("presetBoxLabel")}></span>
          </button>
          <button @click="${(x) => x.applyPreset("shadow")}" ${ref("presetShadow")}>
            <span class="box shadow"></span><span ${ref("presetShadowLabel")}></span>
          </button>
        </div>
        <div class="row">
          <div class="field">
            <label ${ref("styleLabel")}></label>
            <fluent-dropdown
              type="combobox"
              appearance="outline"
              ${ref("styleSel")}
              @change="${(x) => x.syncStyle()}"
            >
              <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
              <input
                slot="control"
                role="combobox"
                aria-haspopup="listbox"
                type="combobox"
                size="1"
                style="width:100%;box-sizing:border-box"
              />
            </fluent-dropdown>
          </div>
          <div class="field">
            <label ${ref("colorLabel")}></label>
            <fluent-dropdown
              type="combobox"
              appearance="outline"
              ${ref("colorSel")}
              @change="${(x) => x.syncColor()}"
            >
              <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
              <input
                slot="control"
                role="combobox"
                aria-haspopup="listbox"
                type="combobox"
                size="1"
                style="width:100%;box-sizing:border-box"
              />
            </fluent-dropdown>
          </div>
        </div>
        <div class="row">
          <div class="field">
            <label ${ref("widthLabel")}></label>
            <fluent-dropdown
              type="combobox"
              appearance="outline"
              ${ref("widthSel")}
              @change="${(x) => x.syncWidth()}"
            >
              <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
              <input
                slot="control"
                role="combobox"
                aria-haspopup="listbox"
                type="combobox"
                size="1"
                style="width:100%;box-sizing:border-box"
              />
            </fluent-dropdown>
          </div>
          <div class="field">
            <label ${ref("applyToLabel")}></label>
            <span ${ref("applyToValue")}></span>
          </div>
        </div>
        <div class="heading" ${ref("previewHeading")}></div>
        <div class="preview">
          <span></span>
          <button class="edge" ${ref("edgeTop")} @click="${(x) => x.toggleEdge("top")}"></button>
          <span></span>
          <button class="edge" ${ref("edgeLeft")} @click="${(x) => x.toggleEdge("left")}"></button>
          <span></span>
          <button
            class="edge"
            ${ref("edgeRight")}
            @click="${(x) => x.toggleEdge("right")}"
          ></button>
          <span></span>
          <button
            class="edge"
            ${ref("edgeBottom")}
            @click="${(x) => x.toggleEdge("bottom")}"
          ></button>
          <span></span>
        </div>
      </div>
      <p class="hint hidden" ${ref("pageHint")}></p>
      <div ${ref("shadingPage")} class="hidden">
        <div class="heading" ${ref("fillHeading")}></div>
        <div class="palette" ${ref("palette")}></div>
        <div class="row">
          <div class="field">
            <label ${ref("shadingApplyLabel")}></label>
            <span ${ref("shadingApplyValue")}></span>
          </div>
        </div>
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyBorders()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-borders-shading-dialog>` — Word's "Borders and Shading" dialog:
 * a Borders tab (paragraph borders with style/color/width + per-edge
 * toggles), a Page Border tab (the same widgets stamped onto the current
 * section's w:pgBorders), and a Shading tab (paragraph fill). The host
 * prefills via `show(tab, border, page)`; OK emits `borders-shading:ok`
 * with a {@link BordersDialogPatch} for the host to route. Rides on
 * `<docen-dialog>` for the modal shell.
 */
@customElement({ name: "docen-borders-shading-dialog", template, styles })
class DocenBordersShadingDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  // The Fluent tablist owns the selected state (activeid → the picked tab);
  // the tab refs stay for label writes only.
  @observable tablist?: HTMLElement & { activeid: string | number | null };
  @observable borderTabBtn?: HTMLElement;
  @observable pageTabBtn?: HTMLElement;
  @observable shadingTabBtn?: HTMLElement;
  @observable borderPage?: HTMLElement;
  @observable shadingPage?: HTMLElement;
  @observable settingHeading?: HTMLElement;
  @observable presetNone?: HTMLButtonElement;
  @observable presetBox?: HTMLButtonElement;
  @observable presetShadow?: HTMLButtonElement;
  @observable presetNoneLabel?: HTMLElement;
  @observable presetBoxLabel?: HTMLElement;
  @observable presetShadowLabel?: HTMLElement;
  @observable styleLabel?: HTMLElement;
  @observable styleSel?: FluentDropdown;
  @observable colorLabel?: HTMLElement;
  @observable colorSel?: FluentDropdown;
  @observable widthLabel?: HTMLElement;
  @observable widthSel?: FluentDropdown;
  @observable applyToLabel?: HTMLElement;
  @observable applyToValue?: HTMLElement;
  @observable previewHeading?: HTMLElement;
  @observable edgeTop?: HTMLButtonElement;
  @observable edgeLeft?: HTMLButtonElement;
  @observable edgeRight?: HTMLButtonElement;
  @observable edgeBottom?: HTMLButtonElement;
  @observable pageHint?: HTMLElement;
  @observable fillHeading?: HTMLElement;
  @observable palette?: HTMLElement;
  @observable shadingApplyLabel?: HTMLElement;
  @observable shadingApplyValue?: HTMLElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  #unobserveLang?: () => void;
  /** Which tab is up, and each tab's staged state (kept across switches). */
  #tab: "border" | "page" | "shading" = "border";
  #border: TabState = { style: "single", color: null, width: 6, sides: emptySides() };
  #page: TabState = { style: "single", color: null, width: 6, sides: emptySides() };
  #fill: string | null = null;

  connectedCallback(): void {
    super.connectedCallback();
    this.#fillCombos();
    this.#applyLabels();
    this.#unobserveLang = observeLang(() => this.#applyLabels());
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  /** Open on `tab`, prefilling the border tabs from the caret paragraph's
   *  attrs and the current section's w:pgBorders (either may be absent). */
  show(
    tab: "border" | "page" | "shading",
    border?: Record<string, unknown> | null,
    page?: Record<string, unknown> | null,
  ): void {
    this.#border = this.#readTab("single", 6, border);
    this.#page = this.#readTab("single", 4, page);
    this.#fill = null;
    this.showTab(tab);
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible tablist change handler — the tabs carry no click
   *  bindings of their own; user picks arrive here via the group's change. */
  onTabChange(): void {
    const id = String(this.tablist?.activeid ?? "");
    const tab = id.replace("bs-tab-", "") as "border" | "page" | "shading";
    // showTab also stamps activeid (programmatic opens); skip the echo.
    if (tab !== this.#tab) this.showTab(tab);
  }

  showTab(tab: "border" | "page" | "shading"): void {
    this.#tab = tab;
    // The page tab reuses the border widgets, plus a hint that the stamp
    // lands on the current section.
    this.borderPage?.classList.toggle("hidden", tab === "shading");
    this.shadingPage?.classList.toggle("hidden", tab !== "shading");
    this.pageHint?.classList.toggle("hidden", tab !== "page");
    if (this.tablist) this.tablist.activeid = `bs-tab-${tab}`;
    this.#loadTabState();
    this.#paintEdges();
    this.#paintPresets();
    if (this.applyToValue)
      this.applyToValue.textContent = t(
        tab === "page" ? "bordersShading.toSection" : "bordersShading.toParagraph",
        this,
      );
  }

  applyPreset(preset: "none" | "box" | "shadow"): void {
    const state = this.#tabState();
    if (preset === "none") state.sides = emptySides();
    else {
      const edge = (): BorderSideState => ({
        style: state.style,
        size: state.width,
        color: state.color,
      });
      state.sides = { top: edge(), bottom: edge(), left: edge(), right: edge() };
      // Word's shadow preset: the bottom/right rules run thick.
      if (preset === "shadow") {
        state.sides.bottom = { ...edge(), size: state.width * 3 };
        state.sides.right = { ...edge(), size: state.width * 3 };
      }
    }
    this.#loadTabState();
    this.#paintEdges();
    this.#paintPresets();
  }

  toggleEdge(side: "top" | "bottom" | "left" | "right"): void {
    const state = this.#tabState();
    state.sides[side] = state.sides[side]
      ? null
      : { style: state.style, size: state.width, color: state.color };
    this.#paintEdges();
    this.#paintPresets();
  }

  syncStyle(): void {
    this.#tabState().style = pickedValue(this.styleSel) ?? "single";
    this.#paintEdges();
  }

  syncColor(): void {
    const v = pickedValue(this.colorSel) ?? "";
    this.#tabState().color = v === "auto" || !v ? null : v;
    this.#paintEdges();
  }

  syncWidth(): void {
    this.#tabState().width = Number(pickedValue(this.widthSel) ?? 6);
    this.#paintEdges();
  }

  pickFill(color: string | null): void {
    this.#fill = color;
    for (const btn of this.palette?.querySelectorAll("button") ?? [])
      btn.setAttribute("aria-pressed", String((btn.dataset.color ?? "") === (color ?? "")));
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applyBorders(): void {
    const state = this.#tabState();
    const patch: BordersDialogPatch =
      this.#tab === "shading"
        ? { tab: "shading", fill: this.#fill }
        : { tab: this.#tab, sides: { ...state.sides } };
    this.$emit("borders-shading:ok", patch);
    this.hide();
  }

  #tabState(): TabState {
    return this.#tab === "page" ? this.#page : this.#border;
  }

  /** Normalize a source attrs object (paragraph `border` or section
   *  `pageBorders`) into a tab state; absent sides read as no border. */
  #readTab(
    fallbackStyle: string,
    fallbackWidth: number,
    src?: Record<string, unknown> | null,
  ): TabState {
    const state: TabState = {
      style: fallbackStyle,
      color: null,
      width: fallbackWidth,
      sides: emptySides(),
    };
    if (!src) return state;
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const edge = src[side] as Record<string, unknown> | null | undefined;
      if (!edge || edge.style === "nil" || edge.style === "none") continue;
      state.sides[side] = {
        style: typeof edge.style === "string" ? edge.style : fallbackStyle,
        size: typeof edge.size === "number" ? edge.size : fallbackWidth,
        color: typeof edge.color === "string" && edge.color !== "auto" ? edge.color : null,
      };
    }
    // The style widgets land on the first live edge so OK-without-edits keeps it.
    const live = state.sides.top ?? state.sides.bottom ?? state.sides.left ?? state.sides.right;
    if (live) {
      state.style = live.style;
      state.color = live.color;
      state.width = live.size;
    }
    return state;
  }

  /** Push the active tab's staged state into the shared widgets. */
  #loadTabState(): void {
    const state = this.#tabState();
    pick(this.styleSel, state.style);
    pick(this.colorSel, state.color ?? "auto");
    pick(this.widthSel, String(state.width));
  }

  /** Render each preview edge from the staged sides — the edge itself shows
   *  the live style; a slot with no edge reads as a grey hairline stub. */
  #paintEdges(): void {
    const state = this.#tabState();
    const cssStyle: Record<string, string> = {
      dashed: "dashed",
      dashSmallGap: "dashed",
      dotDash: "dashed",
      dotted: "dotted",
      double: "double",
      triple: "double",
    };
    const edges = {
      top: this.edgeTop,
      bottom: this.edgeBottom,
      left: this.edgeLeft,
      right: this.edgeRight,
    };
    for (const [side, btn] of Object.entries(edges)) {
      if (!btn) continue;
      const live = state.sides[side as "top" | "bottom" | "left" | "right"];
      if (!live) {
        btn.style.borderColor = "#d0d0d0";
        btn.style.borderStyle = "dashed";
        btn.style.borderWidth = "1px";
        continue;
      }
      btn.style.borderColor = state.color ? `#${state.color}` : "#555555";
      btn.style.borderStyle = cssStyle[live.style] ?? "solid";
      btn.style.borderWidth = `${Math.max(1, Math.round(live.size / 6))}px`;
    }
  }

  /** Highlight the preset matching the current sides (all none / plain box /
   *  thick bottom-right shadow). */
  #paintPresets(): void {
    const state = this.#tabState();
    const list = Object.values(state.sides);
    const none = list.every((s) => !s);
    const box = !none && list.every((s) => !!s);
    const marks: Array<[HTMLButtonElement | undefined, boolean]> = [
      [this.presetNone, none],
      [this.presetBox, box && state.sides.bottom?.size === state.width],
      [this.presetShadow, box && state.sides.bottom?.size !== state.width],
    ];
    for (const [btn, on] of marks) btn?.setAttribute("aria-pressed", String(on));
  }

  #fillCombos(): void {
    const styleBox = listboxOf(this.styleSel);
    if (styleBox && styleBox.children.length === 0)
      styleBox.replaceChildren(
        ...LINE_STYLES.map((style) => {
          const o = opt("", style);
          o.setAttribute("data-style", style);
          return o;
        }),
      );
    const widthBox = listboxOf(this.widthSel);
    if (widthBox && widthBox.children.length === 0)
      widthBox.replaceChildren(
        ...WIDTHS.map(([eighths, label]) => opt(`${label} pt`, String(eighths))),
      );
    const colorBox = listboxOf(this.colorSel);
    if (colorBox && colorBox.children.length === 0)
      colorBox.replaceChildren(...COLORS.map(([hex]) => opt("", hex ?? "auto")));
    if (this.palette && !this.palette.children.length) {
      const none = document.createElement("button");
      none.dataset.color = "";
      none.type = "button";
      none.textContent = "⃠";
      none.addEventListener("click", () => this.pickFill(null));
      this.palette.append(none);
      for (const [hex] of COLORS.filter(([h]) => h)) {
        const btn = document.createElement("button");
        btn.dataset.color = hex ?? "";
        btn.type = "button";
        btn.style.background = `#${hex}`;
        btn.addEventListener("click", () => this.pickFill(hex ?? null));
        this.palette.append(btn);
      }
    }
    this.#styleOptionLabels();
  }

  /** The line-style option labels resolve at refresh time (they are i18n
   *  keys); color names reuse the font dialog's palette keys. */
  #styleOptionLabels(): void {
    for (const o of listboxOf(this.styleSel)?.querySelectorAll("fluent-option") ?? [])
      o.textContent = t(`bordersShading.style-${o.getAttribute("data-style")}`, this);
    for (const o of listboxOf(this.colorSel)?.querySelectorAll("fluent-option") ?? []) {
      const hex = o.getAttribute("value") === "auto" ? null : o.getAttribute("value");
      const key = COLORS.find(([h]) => h === hex)?.[1] ?? "colorAuto";
      o.textContent = t(`fontDialog.${key}`, this);
    }
  }

  #applyLabels(): void {
    this.#fillCombos();
    if (this.dialogEl) this.dialogEl.heading = t("bordersShading.title", this);
    if (this.borderTabBtn) this.borderTabBtn.textContent = t("bordersShading.tabBorder", this);
    if (this.pageTabBtn) this.pageTabBtn.textContent = t("bordersShading.tabPage", this);
    if (this.shadingTabBtn) this.shadingTabBtn.textContent = t("bordersShading.tabShading", this);
    if (this.settingHeading) this.settingHeading.textContent = t("bordersShading.setting", this);
    if (this.presetNoneLabel) this.presetNoneLabel.textContent = t("bordersShading.setNone", this);
    if (this.presetBoxLabel) this.presetBoxLabel.textContent = t("bordersShading.setBox", this);
    if (this.presetShadowLabel)
      this.presetShadowLabel.textContent = t("bordersShading.setShadow", this);
    if (this.styleLabel) this.styleLabel.textContent = t("bordersShading.styleLine", this);
    if (this.colorLabel) this.colorLabel.textContent = t("bordersShading.colorB", this);
    if (this.widthLabel) this.widthLabel.textContent = t("bordersShading.widthB", this);
    if (this.previewHeading) this.previewHeading.textContent = t("bordersShading.preview", this);
    if (this.applyToLabel) this.applyToLabel.textContent = t("bordersShading.applyTo", this);
    if (this.pageHint) this.pageHint.textContent = t("bordersShading.pageHint", this);
    if (this.fillHeading) this.fillHeading.textContent = t("bordersShading.fill", this);
    if (this.shadingApplyLabel)
      this.shadingApplyLabel.textContent = t("bordersShading.applyTo", this);
    if (this.shadingApplyValue)
      this.shadingApplyValue.textContent = t("bordersShading.toParagraph", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    if (this.applyToValue)
      this.applyToValue.textContent = t(
        this.#tab === "page" ? "bordersShading.toSection" : "bordersShading.toParagraph",
        this,
      );
  }
}

export default DocenBordersShadingDialog;
