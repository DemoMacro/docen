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
import { COMMAND_HOST_STYLE, renderIcon } from "./command-helpers";

// ── Office 2013-2022 default theme palette ──
// 10 columns × 6 rows, matching Word/WPS's color grid. Each column is an OOXML
// schemeClr name (w:themeColor) with 6 row hexes: Lighter 80/60/40%, baseline,
// Darker 25/50%. Picking a theme swatch emits a {themeColor, themeTint|themeShade,
// val} object so the DOCX carries theme semantics (Word re-resolves on theme
// change); val is the resolved RGB for in-editor rendering + w:val fallback.
//
// themeTint = lighter% × 255 (offset); themeShade = (1 − darker%) × 255
// (multiplier). Per ECMA-376 (c-rex): L60% → "99", D50% → "80".
const THEME_TINTS = ["CC", "99", "66"] as const; // Lighter 80/60/40%
const THEME_SHADES = ["BF", "80"] as const; // Darker 25/50%
const THEME_PALETTE: readonly { themeColor: string; rows: readonly string[] }[] = [
  { themeColor: "background1", rows: ["FFFFFF", "FFFFFF", "FFFFFF", "FFFFFF", "D9D9D9", "BFBFBF"] },
  { themeColor: "text1", rows: ["D9D9D9", "A6A6A6", "7F7F7F", "000000", "000000", "000000"] },
  { themeColor: "background2", rows: ["F2F2F2", "D9D9D9", "BFBFBF", "E7E6E6", "A6A6A6", "808080"] },
  { themeColor: "text2", rows: ["D6DCE4", "ADB9CA", "8496B0", "44546A", "323E4F", "222A35"] },
  { themeColor: "accent1", rows: ["D9E1F2", "B4C6E7", "8EAADB", "4472C4", "2F5496", "1F3864"] },
  { themeColor: "accent2", rows: ["FCE4D6", "F8CBAD", "F4B183", "ED7D31", "C55A11", "843C0C"] },
  { themeColor: "accent3", rows: ["EDEDED", "DBDBDB", "C9C9C9", "A5A5A5", "7F7F7F", "595959"] },
  { themeColor: "accent4", rows: ["FFF2CC", "FFE699", "FFD966", "FFC000", "BF9000", "806000"] },
  { themeColor: "accent5", rows: ["DEEBF7", "BDD7EE", "9DC3E6", "5B9BD5", "2E75B6", "1F4E79"] },
  { themeColor: "accent6", rows: ["E2EFDA", "C6E0B4", "A9D08E", "70AD47", "538135", "375623"] },
];
const STANDARD_COLORS: readonly string[] = [
  "C00000",
  "FF0000",
  "FFC000",
  "FFFF00",
  "92D050",
  "00B050",
  "00B0F0",
  "0070C0",
  "002060",
  "7030A0",
];

/** A pickable color: a bare upper-hex string (standard/recent/custom) or a
 *  theme-semantic object whose val is the resolved RGB and themeColor/tint/shade
 *  ride along into the DOCX so Word keeps the color theme-bound. */
type ColorValue =
  | string
  | { themeColor: string; val: string; themeTint?: string; themeShade?: string };

// Per-command memory shared across the session — the primary action re-applies
// the last-used color, and the "Recent Colors" row tracks up to 10 picks.
// Keyed by `event` so font-color and shading each keep their own. Module-level
// matches Office's per-feature, app-lifetime color memory.
const MAX_RECENT = 10;
const lastColor = new Map<string, ColorValue>();
const recentColors = new Map<string, ColorValue[]>();

/** Dedup key for a ColorValue (hex or theme tuple) so the same pick doesn't
 *  fill Recent Colors with duplicates. */
function colorKey(value: ColorValue): string {
  return typeof value === "string"
    ? value
    : `${value.themeColor}:${value.themeTint ?? ""}:${value.themeShade ?? ""}`;
}

/** Resolved upper-hex RGB of a ColorValue (for the color stripe + swatch fill). */
function valOf(value: ColorValue | undefined): string | undefined {
  if (value == null) return undefined;
  return typeof value === "string" ? value : value.val;
}

function rememberColor(event: string, value: ColorValue): void {
  lastColor.set(event, value);
  const key = colorKey(value);
  const list = recentColors.get(event) ?? [];
  const at = list.findIndex((v) => colorKey(v) === key);
  if (at === 0) return; // already most-recent
  if (at > 0) list.splice(at, 1); // dedup before promoting to front
  list.unshift(value);
  if (list.length > MAX_RECENT) list.length = MAX_RECENT;
  recentColors.set(event, list);
}

function recentOf(event: string): ColorValue[] {
  return recentColors.get(event) ?? [];
}

// ── HSV ↔ RGB ↔ hex (for the custom color picker) ──

function hsvToRgb(h: number, s: number, v: number): [number, number, number] {
  s /= 100;
  v /= 100;
  const c = v * s;
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = v - c;
  let r = 0;
  let g = 0;
  let b = 0;
  if (h < 60) [r, g, b] = [c, x, 0];
  else if (h < 120) [r, g, b] = [x, c, 0];
  else if (h < 180) [r, g, b] = [0, c, x];
  else if (h < 240) [r, g, b] = [0, x, c];
  else if (h < 300) [r, g, b] = [x, 0, c];
  else [r, g, b] = [c, 0, x];
  return [Math.round((r + m) * 255), Math.round((g + m) * 255), Math.round((b + m) * 255)];
}

function rgbToHsv(r: number, g: number, b: number): [number, number, number] {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const d = max - min;
  let h = 0;
  if (d) {
    if (max === r) h = 60 * (((g - b) / d) % 6);
    else if (max === g) h = 60 * ((b - r) / d + 2);
    else h = 60 * ((r - g) / d + 4);
  }
  if (h < 0) h += 360;
  const s = max ? (d / max) * 100 : 0;
  const v = max * 100;
  return [h, s, v];
}

function rgbToHex(r: number, g: number, b: number): string {
  return [r, g, b]
    .map((n) => n.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}

function hexToRgb(hex: string): [number, number, number] | null {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
  if (!m) return null;
  const n = Number.parseInt(m[1], 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

// Per-instance CSS anchor name so each picker's popover aligns to its own
// primary button, not the viewport corner.
let seq = 0;

const styles = css`
  ${COMMAND_HOST_STYLE}
  /* Split layout: primary (icon + color stripe) hugging a narrow caret that
     opens the palette — Office's compact color split button. */
  fluent-button#target {
    min-height: 26px;
  }
  :host {
    display: inline-flex;
    align-items: stretch;
  }
  .rb-label {
    font-size: 12px;
  }
  /* The glyph stacks over a color stripe showing the last-used color —
     .rb-icon is repurposed as a flex column (COMMAND_HOST_STYLE makes it
     display:contents for the bare-icon case; our later rule wins). */
  fluent-button#target .rb-icon {
    display: flex;
    flex-direction: column;
    align-items: center;
  }
  .rb-color-bar {
    width: 16px;
    height: 3px;
    margin-top: 1px;
    border: 1px solid rgba(0, 0, 0, 0.3);
    background: #000;
  }
  button.cp-caret {
    appearance: none;
    -webkit-appearance: none;
    border: none;
    background: transparent;
    cursor: pointer;
    min-width: 12px;
    width: 12px;
    padding: 0;
    margin-inline-start: 1px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    border-radius: 2px;
    color: inherit;
  }
  button.cp-caret::after {
    content: "";
    display: block;
    border-left: 3px solid transparent;
    border-right: 3px solid transparent;
    border-top: 4px solid currentColor;
  }
  button.cp-caret:hover {
    background: var(--docen-color-hover, rgba(0, 0, 0, 0.06));
  }
  .color-grid {
    popover: auto;
    background: var(--docen-color-bg, #fff);
    border: 1px solid var(--docen-color-divider, #c7c7c7);
    border-radius: 4px;
    padding: 8px;
    box-shadow: 0 2px 12px rgba(0, 0, 0, 0.18);
    margin: 0;
    min-width: 188px;
    position-anchor: var(--cp-anchor);
    inset-block-start: anchor(bottom);
    inset-inline-start: anchor(self-start);
    inset-inline-end: auto;
  }
  .cp-section {
    font-size: 11px;
    color: #666;
    margin: 6px 2px 2px;
  }
  .cp-section:first-of-type {
    margin-top: 0;
  }
  .cp-none,
  .cp-more {
    display: block;
    width: 100%;
    box-sizing: border-box;
    padding: 4px 6px;
    margin: 0 0 4px;
    border: 1px solid var(--docen-color-divider, #c7c7c7);
    background: #fff;
    cursor: pointer;
    font-size: 12px;
    text-align: start;
    border-radius: 3px;
  }
  .cp-more {
    margin: 4px 0 0;
  }
  .cp-none:hover,
  .cp-more:hover {
    background: var(--docen-color-hover, rgba(0, 0, 0, 0.06));
  }
  .cp-swatches {
    display: grid;
    grid-template-columns: repeat(10, 16px);
    gap: 2px;
  }
  .cp-swatch {
    width: 16px;
    height: 16px;
    padding: 0;
    cursor: pointer;
    border: 1px solid rgba(0, 0, 0, 0.25);
    border-radius: 2px;
  }
  .cp-swatch:hover {
    outline: 1.5px solid #333;
    outline-offset: 0;
  }
  .cp-hidden {
    display: none;
  }
  /* Custom color view — opened by "More Colors", lives inside the same
     anchored popover so it never jumps to the viewport corner (the native
     <input type=color> picker did). */
  .cp-sv {
    width: 172px;
    height: 110px;
    position: relative;
    cursor: crosshair;
    border: 1px solid var(--docen-color-divider, #c7c7c7);
    border-radius: 3px;
    background-color: hsl(var(--cp-hue, 0), 100%, 50%);
    background-image:
      linear-gradient(to top, #000, transparent), linear-gradient(to right, #fff, transparent);
    touch-action: none;
  }
  .cp-sv-cursor {
    position: absolute;
    width: 10px;
    height: 10px;
    border: 1.5px solid #fff;
    border-radius: 50%;
    box-shadow: 0 0 0 1px rgba(0, 0, 0, 0.4);
    transform: translate(-50%, -50%);
    pointer-events: none;
  }
  .cp-hue {
    -webkit-appearance: none;
    appearance: none;
    width: 172px;
    height: 10px;
    margin: 6px 0 0;
    padding: 0;
    background: linear-gradient(to right, #f00, #ff0, #0f0, #0ff, #00f, #f0f, #f00);
    border: 1px solid var(--docen-color-divider, #c7c7c7);
    border-radius: 5px;
  }
  .cp-hue::-webkit-slider-thumb {
    -webkit-appearance: none;
    width: 12px;
    height: 14px;
    background: #fff;
    border: 1px solid #333;
    border-radius: 2px;
    cursor: pointer;
  }
  .cp-hue::-moz-range-thumb {
    width: 12px;
    height: 14px;
    background: #fff;
    border: 1px solid #333;
    border-radius: 2px;
    cursor: pointer;
  }
  .cp-cust-row {
    display: flex;
    align-items: center;
    gap: 6px;
    margin-top: 6px;
  }
  .cp-cust-swatch {
    width: 26px;
    height: 24px;
    flex-shrink: 0;
    border: 1px solid var(--docen-color-divider, #c7c7c7);
    border-radius: 3px;
  }
  .cp-hex {
    flex: 1;
    min-width: 0;
    padding: 3px 6px;
    font-size: 12px;
    border: 1px solid var(--docen-color-divider, #c7c7c7);
    border-radius: 3px;
    font-family: monospace;
    text-transform: uppercase;
  }
  .cp-cust-actions {
    display: flex;
    gap: 6px;
    margin-top: 6px;
  }
  .cp-back,
  .cp-apply {
    flex: 1;
    padding: 4px 6px;
    font-size: 12px;
    cursor: pointer;
    border: 1px solid var(--docen-color-divider, #c7c7c7);
    border-radius: 3px;
    background: #fff;
  }
  .cp-apply {
    background: var(--docen-color-accent, #0078d4);
    color: #fff;
    border-color: transparent;
  }
  .cp-back:hover {
    background: var(--docen-color-hover, rgba(0, 0, 0, 0.06));
  }
`;

const template = html<DocenColorPicker>`
  <fluent-button
    id="target"
    part="button"
    appearance="subtle"
    ?icon-only="${(x) => x.iconOnly}"
    ${ref("btn")}
  >
    <span class="rb-icon" ${ref("iconSlot")}></span>
    <span class="rb-label">${(x) => x.visibleLabel}</span>
  </fluent-button>
  <button type="button" class="cp-caret" part="caret" ${ref("caret")}></button>
  <div popover="auto" part="grid" class="color-grid" ${ref("grid")}>
    <div class="cp-picker" ${ref("picker")}>
      <button type="button" class="cp-none" part="none" ${ref("noneBtn")}></button>
      <div class="cp-section" data-i18n="theme-colors"></div>
      <div class="cp-swatches" part="theme" ${ref("themeEl")}></div>
      <div class="cp-section" data-i18n="standard-colors"></div>
      <div class="cp-swatches" part="standard" ${ref("standardEl")}></div>
      <div
        class="cp-section"
        data-i18n="recent-colors"
        part="recent-label"
        ${ref("recentLabel")}
      ></div>
      <div class="cp-swatches cp-hidden" part="recent" ${ref("recentEl")}></div>
      <button type="button" class="cp-more" part="more" ${ref("moreBtn")}></button>
    </div>
    <div class="cp-custom cp-hidden" ${ref("customEl")}>
      <div class="cp-sv" ${ref("sv")}>
        <div class="cp-sv-cursor"></div>
      </div>
      <input
        type="range"
        class="cp-hue"
        min="0"
        max="360"
        step="1"
        value="0"
        aria-label="Hue"
        ${ref("hue")}
      />
      <div class="cp-cust-row">
        <span class="cp-cust-swatch" ${ref("custSwatch")}></span>
        <input
          type="text"
          class="cp-hex"
          maxlength="7"
          spellcheck="false"
          aria-label="Hex color"
          ${ref("hex")}
        />
      </div>
      <div class="cp-cust-actions">
        <button type="button" class="cp-back" part="back" ${ref("backBtn")}></button>
        <button type="button" class="cp-apply" part="apply" ${ref("applyBtn")}></button>
      </div>
    </div>
  </div>
  <fluent-tooltip anchor="target" positioning="top" ${ref("tooltipEl")}>
    <span class="rb-tip">${(x) => x.tooltipText}</span>
  </fluent-tooltip>
`;

/**
 * `<docen-color-picker icon="font-color" event="font-color" default-color="000000">`
 * — an Office-style split color button: the primary action re-applies the
 * last-used color (shown as a stripe under the icon), and the caret opens a
 * palette popover anchored under the button. The palette mirrors Word/WPS:
 * No Color, theme colors (10×6 tints/shades), standard colors, recent colors,
 * and More Colors — which opens an in-popover HSV picker (not the native
 * `<input type=color>` popup, whose position the browser controls and strands
 * at the viewport corner). Selecting a color emits `command { event, value }`
 * where value is an upper-hex color or "none" (clear).
 */
@customElement({ name: "docen-color-picker", template, styles })
class DocenColorPicker extends FASTElement {
  @attr label?: string;
  @attr icon?: string;
  @attr event?: string;
  @attr tooltip?: string;
  @attr({ attribute: "default-color" }) defaultColor?: string;
  @attr({ attribute: "icon-only", mode: "boolean" }) iconOnly?: boolean;

  @observable btn?: HTMLElement;
  @observable caret?: HTMLElement;
  @observable grid?: HTMLElement;
  @observable picker?: HTMLElement;
  @observable customEl?: HTMLElement;
  @observable themeEl?: HTMLElement;
  @observable standardEl?: HTMLElement;
  @observable recentEl?: HTMLElement;
  @observable recentLabel?: HTMLElement;
  @observable noneBtn?: HTMLElement;
  @observable moreBtn?: HTMLElement;
  @observable sv?: HTMLElement;
  @observable hue?: HTMLInputElement;
  @observable hex?: HTMLInputElement;
  @observable custSwatch?: HTMLElement;
  @observable applyBtn?: HTMLElement;
  @observable backBtn?: HTMLElement;
  @observable iconSlot?: HTMLSpanElement;
  @observable tooltipEl?: HTMLElement;

  readonly anchorId = `--cp-${++seq}`;
  #bar?: HTMLElement;
  #focusCleanup?: () => void;
  #obsLang?: () => void;
  #endSvDrag?: () => void;
  #hsv: [number, number, number] = [0, 100, 100];
  #pendingHex = "000000";

  /** Command key (falls back to label) for the per-feature color memory. */
  get eventName(): string {
    return this.event || this.label || "";
  }
  /** Icon-only hides the visible label (it still feeds the tooltip). */
  get visibleLabel(): string {
    return this.iconOnly ? "" : (this.label ?? "");
  }
  get tooltipText(): string {
    return this.tooltip || this.label || "";
  }

  iconChanged(): void {
    this.#renderIcon();
  }
  iconOnlyChanged(): void {
    this.#syncIconSlot();
  }
  defaultColorChanged(): void {
    this.#refreshBar();
  }
  // label/tooltip are template-bound (.rb-label/.rb-tip) — no changed callback.

  connectedCallback(): void {
    super.connectedCallback();
    // Anchor the popover to THIS instance's primary button (same-shadow) —
    // anchoring the host crosses the shadow boundary and the browser strands
    // the popover at the viewport corner.
    if (this.btn) this.btn.style.anchorName = this.anchorId;
    if (this.grid) this.grid.style.setProperty("--cp-anchor", this.anchorId);
    if (this.tooltipEl) this.tooltipEl.style.positionAnchor = this.anchorId;
    // .rb-icon holds the glyph AND the color stripe beneath it.
    this.#bar = document.createElement("span");
    this.#bar.className = "rb-color-bar";
    this.#syncIconSlot();
    this.#renderIcon();
    this.#renderPalette();
    this.#refreshBar();
    // Keep the editor's text selection across swatch/button clicks, but let the
    // hex/hue inputs take focus (they need it to accept typing/dragging). Unlike
    // a blanket preventDefault, this guard excludes <input> targets.
    // composedPath()[0] is the real target — this listener is on the host, so
    // event.target is retargeted to the host across the shadow boundary, a naive
    // `.closest("input")` never matches, and every mousedown (including on the
    // hex field and the hue slider) got preventDefaulted, breaking both.
    const onMousedown = (event: Event): void => {
      if ((event.composedPath()[0] as HTMLElement | null)?.closest("input")) return;
      event.preventDefault();
    };
    this.addEventListener("mousedown", onMousedown, { capture: true });
    this.#focusCleanup = () =>
      this.removeEventListener("mousedown", onMousedown, { capture: true });
    // Primary click re-applies the last-used color (Office split behavior).
    this.btn?.addEventListener("click", (event) => {
      event.stopPropagation();
      this.#applyLast();
    });
    // Caret opens the palette (popover=auto handles light-dismiss on outside click).
    this.caret?.addEventListener("click", (event) => {
      event.stopPropagation();
      this.#open();
    });
    this.noneBtn?.addEventListener("click", () => {
      this.#hide();
      this.#emit("none");
    });
    this.moreBtn?.addEventListener("click", () => this.#showCustom());
    this.applyBtn?.addEventListener("click", () => {
      this.#showPicker();
      this.#hide();
      this.#pick(this.#pendingHex);
    });
    this.backBtn?.addEventListener("click", () => this.#showPicker());
    this.hue?.addEventListener("input", () => {
      this.#hsv[0] = Number(this.hue?.value) || 0;
      this.#syncCustom();
    });
    this.hex?.addEventListener("change", () => {
      const rgb = hexToRgb(this.hex?.value ?? "");
      if (rgb) this.#hsv = rgbToHsv(rgb[0], rgb[1], rgb[2]);
      this.#syncCustom();
    });
    this.sv?.addEventListener("pointerdown", (event) => this.#onSvPointerDown(event));
    this.#applyI18n();
    this.#obsLang = observeLang(() => this.#applyI18n());
  }

  disconnectedCallback(): void {
    // Release an in-flight SV drag: its document-level listeners would
    // otherwise keep firing #updateSvFromPointer (getBoundingClientRect on a
    // detached #sv) after the picker is gone, and never unbind without a
    // pointerup/pointercancel.
    this.#endSvDrag?.();
    this.#focusCleanup?.();
    this.#obsLang?.();
    super.disconnectedCallback();
  }

  /** Upper-hex default color for the primary action before any pick (font →
   *  black, shading → yellow); overridable via the `default-color` attribute. */
  #defaultHex(): string {
    const raw = this.defaultColor ?? "";
    return raw.replace(/^#/, "").toUpperCase() || "000000";
  }

  #applyLast(): void {
    const value = lastColor.get(this.eventName) ?? this.#defaultHex();
    rememberColor(this.eventName, value);
    this.#refreshBar();
    this.#emit(value);
  }

  #pick(value: ColorValue): void {
    rememberColor(this.eventName, value);
    this.#refreshBar();
    this.#emit(value);
  }

  #open(): void {
    const grid = this.grid as unknown as { showPopover?(): void } | null;
    if (!grid || this.grid?.matches(":popover-open")) return;
    this.#showPicker();
    this.#renderPalette();
    grid.showPopover?.();
  }

  #hide(): void {
    (this.grid as unknown as { hidePopover?(): void } | null)?.hidePopover?.();
  }

  #showPicker(): void {
    this.customEl?.classList.add("cp-hidden");
    this.picker?.classList.remove("cp-hidden");
  }

  #showCustom(): void {
    this.picker?.classList.add("cp-hidden");
    this.customEl?.classList.remove("cp-hidden");
    const hex = valOf(lastColor.get(this.eventName)) ?? this.#defaultHex();
    const rgb = hexToRgb(hex) ?? [0, 0, 0];
    this.#hsv = rgbToHsv(rgb[0], rgb[1], rgb[2]);
    this.#syncCustom();
  }

  /** Repaint the HSV picker from #hsv: SV pane hue + cursor, hue slider, hex
   *  field (unless the user is typing in it), preview swatch, pending value. */
  #syncCustom(): void {
    const [h, s, v] = this.#hsv;
    if (this.sv) this.sv.style.setProperty("--cp-hue", String(h));
    const cursor = this.sv?.querySelector(".cp-sv-cursor") as HTMLElement | null;
    if (cursor) {
      cursor.style.left = `${s}%`;
      cursor.style.top = `${100 - v}%`;
    }
    if (this.hue) this.hue.value = String(Math.round(h));
    const [r, g, b] = hsvToRgb(h, s, v);
    const hex = rgbToHex(r, g, b);
    if (this.hex && document.activeElement !== this.hex) this.hex.value = `#${hex}`;
    if (this.custSwatch) this.custSwatch.style.background = `#${hex}`;
    this.#pendingHex = hex;
  }

  #onSvPointerDown(event: PointerEvent): void {
    if (!this.sv) return;
    event.preventDefault();
    const move = (ev: PointerEvent): void => this.#updateSvFromPointer(ev);
    const end = (): void => {
      document.removeEventListener("pointermove", move);
      document.removeEventListener("pointerup", end);
      document.removeEventListener("pointercancel", end);
      this.#endSvDrag = undefined;
    };
    this.#endSvDrag = end;
    document.addEventListener("pointermove", move);
    document.addEventListener("pointerup", end);
    document.addEventListener("pointercancel", end);
    this.#updateSvFromPointer(event);
  }

  #updateSvFromPointer(event: PointerEvent): void {
    if (!this.sv) return;
    const rect = this.sv.getBoundingClientRect();
    const x = Math.max(0, Math.min(rect.width, event.clientX - rect.left));
    const y = Math.max(0, Math.min(rect.height, event.clientY - rect.top));
    this.#hsv[1] = (x / rect.width) * 100; // saturation
    this.#hsv[2] = (1 - y / rect.height) * 100; // value
    this.#syncCustom();
  }

  #refreshBar(): void {
    if (!this.#bar) return;
    const hex = valOf(lastColor.get(this.eventName)) ?? this.#defaultHex();
    this.#bar.style.background = `#${hex}`;
  }

  #renderPalette(): void {
    if (!this.themeEl || !this.standardEl || !this.recentEl || !this.recentLabel) return;
    this.themeEl.replaceChildren();
    this.standardEl.replaceChildren();
    this.recentEl.replaceChildren();
    // Theme grid is row-major to match Word's layout (6 rows × 10 columns).
    // Each swatch emits a theme-semantic object (themeColor + tint/shade + val).
    for (let row = 0; row < 6; row++) {
      for (const col of THEME_PALETTE) {
        const val = col.rows[row];
        const value: ColorValue =
          row < 3
            ? { themeColor: col.themeColor, val, themeTint: THEME_TINTS[row] }
            : row === 3
              ? { themeColor: col.themeColor, val }
              : { themeColor: col.themeColor, val, themeShade: THEME_SHADES[row - 4] };
        this.themeEl.append(this.#themeSwatch(value));
      }
    }
    for (const hex of STANDARD_COLORS) this.standardEl.append(this.#hexSwatch(hex));
    const recent = recentOf(this.eventName);
    for (const value of recent) {
      this.recentEl.append(
        typeof value === "string" ? this.#hexSwatch(value) : this.#themeSwatch(value),
      );
    }
    const empty = recent.length === 0;
    this.recentEl.classList.toggle("cp-hidden", empty);
    this.recentLabel.classList.toggle("cp-hidden", empty);
  }

  /** A bare-hex swatch (standard/recent/custom colors) — emits the hex string. */
  #hexSwatch(hex: string): HTMLButtonElement {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "cp-swatch";
    btn.style.backgroundColor = `#${hex}`;
    btn.title = `#${hex}`;
    btn.addEventListener("click", () => {
      this.#hide();
      this.#pick(hex);
    });
    return btn;
  }

  /** A theme swatch — fills with val (resolved RGB) but emits the theme-semantic
   *  object so the pick stays theme-bound in the DOCX. */
  #themeSwatch(value: {
    themeColor: string;
    val: string;
    themeTint?: string;
    themeShade?: string;
  }): HTMLButtonElement {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "cp-swatch";
    btn.style.backgroundColor = `#${value.val}`;
    btn.title = `#${value.val}`;
    btn.addEventListener("click", () => {
      this.#hide();
      this.#pick(value);
    });
    return btn;
  }

  #syncIconSlot(): void {
    if (this.iconSlot) this.iconSlot.slot = this.iconOnly ? "" : "start";
  }

  #renderIcon(): void {
    if (!this.iconSlot) return;
    renderIcon(this.iconSlot, this.icon ?? "");
    // renderIcon clears .rb-icon via replaceChildren, so re-append the stripe.
    if (this.#bar) this.iconSlot.append(this.#bar);
  }

  #applyI18n(): void {
    if (this.noneBtn) this.noneBtn.textContent = t("ribbon.opt.no-color", this);
    if (this.moreBtn) this.moreBtn.textContent = t("ribbon.opt.more-colors", this);
    if (this.applyBtn) this.applyBtn.textContent = t("ribbon.opt.color-ok", this);
    if (this.backBtn) this.backBtn.textContent = t("ribbon.opt.color-back", this);
    this.shadowRoot?.querySelectorAll<HTMLElement>("[data-i18n]").forEach((el) => {
      el.textContent = t("ribbon.opt." + (el.dataset.i18n ?? ""), this);
    });
    if (this.caret) this.caret.setAttribute("aria-label", t("ribbon.opt.more-colors", this));
  }

  #emit(value?: ColorValue): void {
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: { event: this.eventName, value, source: this },
      }),
    );
  }
}

export default DocenColorPicker;
