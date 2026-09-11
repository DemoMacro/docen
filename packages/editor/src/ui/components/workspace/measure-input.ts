import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** Spinner increments per unit id — a half unit for counts and points, the
 *  smallest readable tick per length unit. */
const UNIT_STEP: Record<string, number> = {
  char: 0.5,
  line: 0.5,
  pt: 0.5,
  mm: 0.5,
  cm: 0.1,
  in: 0.05,
};

/** Hold-to-repeat tuning (Word's spinner starts repeating after a beat). */
const HOLD_DELAY_MS = 400;
const HOLD_RATE_MS = 80;

/** Length units to centimeters — the dialogs' canonical patch unit, so a
 *  measure written in any unit resolves before it rides an existing patch. */
const UNIT_TO_CM: Record<string, number> = { cm: 1, mm: 0.1, in: 2.54, pt: 2.54 / 72 };

/** A value in the picked unit resolved to centimeters (unknown ids pass
 *  through — the unit lists only carry the table's keys). */
export function measureToCm(value: number, unit: string): number {
  return value * (UNIT_TO_CM[unit] ?? 1);
}

const styles = css`
  :host {
    display: flex;
    align-items: center;
    gap: 6px;
    min-width: 0;
    flex: 1 1 auto;
  }
  fluent-text-input {
    min-width: 0;
    /* basis 0 — the dropdown's natural basis is 100% and would crowd the
       number box out of the row. */
    flex: 1 1 0;
  }
  fluent-dropdown {
    flex: 0 0 88px;
    /* the dropdown's min-content is its listbox — without this the flex
       floor crowds the number box out of the row. */
    min-width: 0;
    width: 88px;
  }
  :host([no-unit]) fluent-dropdown {
    display: none;
  }
  .spin {
    display: flex;
    flex-direction: column;
    align-self: stretch;
    justify-content: center;
    /* the FAST input's 100% basis never grows, so the leftover space pools
       to its right — the auto margin walks the spinner to the box edge. */
    margin-inline-start: auto;
    border-inline-start: 1px solid var(--docen-color-divider);
  }
  .spin button {
    display: flex;
    align-items: center;
    justify-content: center;
    border: 0;
    padding: 0 5px;
    background: transparent;
    cursor: pointer;
    height: 15px;
    color: var(--docen-color-text-muted);
  }
  .spin button:hover {
    background: var(--docen-color-subtle-background-hover);
    color: var(--docen-color-text);
  }
  .spin svg {
    display: block;
  }
`;

const template = html<DocenMeasureInput>`
  <fluent-text-input ${ref("box")} @change="${(x) => x.onTextChange()}">
    <div class="spin" slot="end">
      <button
        type="button"
        tabindex="-1"
        aria-hidden="true"
        @pointerdown="${(x) => x.startHold(1)}"
        @pointerup="${(x) => x.stopHold()}"
        @pointerleave="${(x) => x.stopHold()}"
        @pointercancel="${(x) => x.stopHold()}"
      >
        <svg viewBox="0 0 8 5" width="8" height="5">
          <path fill="currentColor" d="M4 0 8 5H0Z" />
        </svg>
      </button>
      <button
        type="button"
        tabindex="-1"
        aria-hidden="true"
        @pointerdown="${(x) => x.startHold(-1)}"
        @pointerup="${(x) => x.stopHold()}"
        @pointerleave="${(x) => x.stopHold()}"
        @pointercancel="${(x) => x.stopHold()}"
      >
        <svg viewBox="0 0 8 5" width="8" height="5">
          <path fill="currentColor" d="M0 0h8L4 5Z" />
        </svg>
      </button>
    </div>
  </fluent-text-input>
  <fluent-dropdown type="combobox" appearance="outline" ${ref("unitDrop")}>
    <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
    <input slot="control" role="combobox" aria-haspopup="listbox" type="combobox" />
  </fluent-dropdown>
`;

/** A `fluent-text-input` widget plus its string value accessor (the value
 *  lives on the `value` property, like a native input). */
type FluentTextInput = HTMLElement & { value: string; disabled: boolean };

/** A `fluent-dropdown` combobox plus its picked value (null = none). */
type FluentDropdown = HTMLElement & { value: string | null; disabled: boolean };

/**
 * `<docen-measure-input>` — Word's measure box: a numeric input with up/down
 * spinner buttons in its end slot and a unit dropdown on the right. The host
 * reads `value` (the number, null when blank) and `unit` (a `units=` id);
 * the dropdown lists the ids given via `units`, labeled in the UI language.
 * A typed or pasted number with a trailing unit word splits into both slots
 * (the word matches an option's label or id); a bare number keeps the current
 * unit. `no-unit` drops the dropdown for a fixed-unit box; the spinner never
 * steps below zero. Repeats while held, with Word's initial delay.
 */
@customElement({ name: "docen-measure-input", template, styles })
export default class DocenMeasureInput extends FASTElement {
  @observable box?: FluentTextInput;
  @observable unitDrop?: FluentDropdown;
  /** Space-separated unit ids backing the dropdown options ("char pt mm…"). */
  @observable units = "";
  @observable noUnit = false;

  #unobserveLang?: () => void;
  #holdTimer?: number;
  #holdInterval?: number;

  connectedCallback(): void {
    super.connectedCallback();
    // Read the template's plain attributes directly — the FAST attr→property
    // sync doesn't fire for template-stamped instances here. The language
    // flip below re-renders the option labels; all paths are idempotent.
    this.units = this.getAttribute("units") ?? "";
    this.noUnit = this.hasAttribute("no-unit");
    this.#unobserveLang = observeLang(() => this.#renderUnitOptions());
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    this.stopHold();
    super.disconnectedCallback();
  }

  unitsChanged(): void {
    // Deferred: the dropdown resolves its listbox through a slotchange
    // microtask, so a value write before connection completes (the initial
    // attribute read in connectedCallback) finds no listbox and throws.
    queueMicrotask(() => this.#renderUnitOptions());
  }

  /** The parsed number (null when the box is blank or not a plain number). */
  get value(): number | null {
    const text = (this.box?.value ?? "").trim();
    const n = Number(text);
    return text === "" || !Number.isFinite(n) ? null : n;
  }

  set value(v: number | null | undefined) {
    if (this.box)
      this.box.value = v == null || !Number.isFinite(v) ? "" : String(Math.round(v * 100) / 100);
  }

  /** The picked unit id ("" for a no-unit box). Before the dropdown stamps
   *  its default (its listbox resolves late), the list's first unit is the
   *  effective pick — the list's order is the host's preference. */
  get unit(): string {
    if (this.noUnit) return "";
    return this.unitDrop?.value ?? this.units.split(/\s+/).filter(Boolean)[0] ?? "";
  }

  set unit(id: string) {
    if (this.unitDrop && id) this.unitDrop.value = id;
  }

  get disabled(): boolean {
    return this.box?.disabled ?? false;
  }

  set disabled(v: boolean) {
    if (this.box) this.box.disabled = v;
    if (this.unitDrop) this.unitDrop.disabled = v;
  }

  // Template-visible spinner handlers (#private names can't be referenced
  // from FAST bindings).
  startHold(dir: 1 | -1): void {
    this.#spinOnce(dir);
    this.#holdTimer = window.setTimeout(() => {
      this.#holdInterval = window.setInterval(() => this.#spinOnce(dir), HOLD_RATE_MS);
    }, HOLD_DELAY_MS);
  }

  stopHold(): void {
    window.clearTimeout(this.#holdTimer);
    window.clearInterval(this.#holdInterval);
    this.#holdTimer = undefined;
    this.#holdInterval = undefined;
  }

  /** The smart-measure parse: a number plus a trailing unit word fills both
   *  slots, matching the word against the options' labels or ids. Blank or
   *  bare-number text needs no split. */
  onTextChange(): void {
    const text = (this.box?.value ?? "").trim();
    if (text === "" || Number.isFinite(Number(text))) return;
    const m = /^(-?[\d.]+)\s*(.+)$/.exec(text);
    if (!m) return;
    const word = m[2].trim().toLowerCase();
    for (const opt of this.unitDrop?.querySelectorAll("fluent-option") ?? []) {
      const id = opt.getAttribute("value") ?? "";
      if (word === id || word === (opt.textContent ?? "").trim().toLowerCase()) {
        this.unit = id;
        this.value = Number(m[1]);
        return;
      }
    }
  }

  #spinOnce(dir: 1 | -1): void {
    const next = (this.value ?? 0) + dir * (UNIT_STEP[this.unit] ?? 0.5);
    this.value = Math.max(0, Math.round(next * 100) / 100);
  }

  /** Rebuilds the unit options from the `units` id list; the display words
   *  re-render with the UI language. A box with no pick lands on the first
   *  unit (the list's order is the host's preferred order). */
  #renderUnitOptions(): void {
    const drop = this.unitDrop;
    const listbox = drop?.querySelector("fluent-listbox");
    if (!listbox) return;
    // The pick rides the option's own `selected` attribute (FAST's
    // defaultSelected): the dropdown stamps and displays it once its listbox
    // resolves. The framework setter is never called from here — it reaches
    // for the control and listbox through FAST's update queue and slotchange,
    // and throws when those haven't resolved yet.
    const keep = drop?.value ?? null;
    listbox.textContent = "";
    const ids = this.units.split(/\s+/).filter(Boolean);
    for (const id of ids) {
      const opt = document.createElement("fluent-option");
      opt.setAttribute("value", id);
      opt.textContent = t(`unit.${id}`, this);
      if (keep ? id === keep : id === ids[0]) opt.setAttribute("selected", "");
      listbox.append(opt);
    }
  }
}
