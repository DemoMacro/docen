export interface RibbonOption {
  text: string;
  value?: string;
  disabled?: boolean;
}

// Per-instance CSS anchor name so each combobox's listbox popover is positioned
// against its own dropdown (CSS Anchor Positioning: `anchor-name` on the
// dropdown, matching `position-anchor` on the listbox).
let seq = 0;

const template = document.createElement("template");
template.innerHTML = `
  <style>
    :host { display: inline-flex; width: 120px; }
    :host([size="short"]) { width: 112px; }
    :host([size="long"]) { width: 200px; }
    fluent-dropdown { width: 100%; min-width: 0; }
    input { width: 100%; box-sizing: border-box; }
  </style>
  <fluent-dropdown type="combobox" appearance="outline" part="dropdown">
    <fluent-listbox part="listbox" popover="manual" tabindex="-1"></fluent-listbox>
    <input slot="control" role="combobox" aria-haspopup="listbox" type="combobox" part="input" size="1" style="width:100%;box-sizing:border-box" />
  </fluent-dropdown>`;

/**
 * `<docen-ribbon-combobox value="Calibri" event="font" size="short" items='[{...}]'>`
 * — a typeable selector (font name / size). Mirrors the documented
 * `<fluent-dropdown type="combobox">` shape: a `<fluent-listbox popover="manual">`
 * of `<fluent-option>`s plus an `<input slot="control">`, wired together with a
 * per-instance CSS anchor (`anchor-name` / `position-anchor`) so the popover
 * floats under this dropdown. `value` seeds the input; `change` emits `command`
 * with `{ event, value }`. All visuals/interaction (filtering, keyboard, list
 * positioning) are Fluent's.
 */
class DocenRibbonCombobox extends HTMLElement {
  static get observedAttributes(): string[] {
    return ["value", "items", "event", "size", "source"];
  }

  readonly #id = `rb-cb-${++seq}`;
  readonly #anchor = `--${this.#id}`;
  #dd?: HTMLElement;
  #lb?: HTMLElement;
  #input?: HTMLInputElement;

  attributeChangedCallback(name: string): void {
    if (!this.shadowRoot) return;
    if (name === "items") this.#renderItems();
    if (name === "value") this.#syncValue();
  }

  connectedCallback(): void {
    if (!this.shadowRoot) {
      this.attachShadow({ mode: "open" }).append(template.content.cloneNode(true));
    }
    this.#dd = this.shadowRoot!.querySelector("fluent-dropdown")!;
    this.#lb = this.shadowRoot!.querySelector("fluent-listbox")!;
    this.#input = this.shadowRoot!.querySelector("input")!;
    // Anchor the listbox popover to this dropdown (CSS Anchor Positioning).
    this.#lb.id = this.#id;
    this.#input.setAttribute("aria-controls", this.#id);
    this.#dd.style.anchorName = this.#anchor;
    this.#lb.style.positionAnchor = this.#anchor;
    this.#renderItems();
    this.#syncValue();
    this.#dd.addEventListener("change", () => this.#emit());
    if (this.getAttribute("source") === "local-fonts") void this.#loadLocalFonts();
  }

  /**
   * Enumerate locally-installed fonts via the Local Font Access API
   * (Chrome/Edge; requires user permission). On success, replaces the seeded
   * fallback list with the host's real font families (de-duplicated). No-op on
   * unsupported browsers or denied permission — the fallback list stays.
   */
  async #loadLocalFonts(): Promise<void> {
    const query = (
      window as unknown as {
        queryLocalFonts?: () => Promise<ReadonlyArray<{ family: string }> | undefined>;
      }
    ).queryLocalFonts;
    if (typeof query !== "function") return;
    try {
      const fonts = await query.call(window);
      if (!this.isConnected || !fonts?.length) return;
      const seen = new Set<string>();
      const families: string[] = [];
      for (const font of fonts) {
        if (font.family && !seen.has(font.family)) {
          seen.add(font.family);
          families.push(font.family);
        }
      }
      this.setAttribute("items", JSON.stringify(families.map((family) => ({ text: family }))));
    } catch {
      // Permission denied or enumeration failed — keep the fallback list.
    }
  }

  get items(): RibbonOption[] {
    try {
      return JSON.parse(this.getAttribute("items") ?? "[]") as RibbonOption[];
    } catch {
      return [];
    }
  }

  get event(): string {
    return this.getAttribute("event") ?? "";
  }

  get value(): string {
    // fluent renders its own control input (insertControl drops the seeded
    // one); read its current text rather than the detached original.
    const control = (this.#dd as unknown as { control?: HTMLInputElement } | undefined)?.control;
    return control?.value ?? this.#input?.value ?? "";
  }

  #renderItems(): void {
    if (!this.#lb) return;
    // short (font-size) options center; every option hides the selected
    // checkmark (data-center implies it, the font box uses data-no-checkmark).
    const center = this.getAttribute("size") === "short";
    this.#lb.replaceChildren();
    for (const item of this.items) {
      const option = document.createElement("fluent-option");
      if (center) option.setAttribute("data-center", "");
      else option.setAttribute("data-no-checkmark", "");
      option.textContent = item.text;
      if (item.value) option.setAttribute("value", item.value);
      if (item.disabled) option.setAttribute("disabled", "");
      this.#lb.append(option);
    }
  }

  #syncValue(): void {
    if (!this.#lb || !this.#dd) return;
    // Capture in a const so the `apply` closure below holds the narrowed type
    // (TS won't keep the `!this.#lb` narrowing across the function boundary).
    const lb = this.#lb;
    const hostValue = this.getAttribute("value") ?? "";
    // fluent-dropdown's connectedCallback enqueues insertControl(), which drops
    // the seeded <input> and renders its own (value bound to an internal
    // observable, initially ""). selectOption() both marks the option selected
    // and writes its displayValue to that control input — so the box shows the
    // default and opens with the right row highlighted. Defer until insertControl
    // (control) and slotchange (this.listbox) have both settled.
    const dd = this.#dd as unknown as {
      listbox?: unknown;
      control?: HTMLInputElement;
      selectOption(i: number): void;
    };
    let tries = 0;
    const apply = (): void => {
      if (!this.isConnected || tries++ > 10) return;
      if (!dd.listbox || !dd.control) {
        requestAnimationFrame(apply);
        return;
      }
      const opts = lb.querySelectorAll("fluent-option");
      let idx = -1;
      opts.forEach((opt, i) => {
        const ov = opt.getAttribute("value");
        const ot = (opt.textContent ?? "").trim();
        if (idx < 0 && hostValue !== "" && (ov === hostValue || ot === hostValue)) idx = i;
      });
      if (idx >= 0) {
        dd.selectOption(idx);
      } else {
        // Value is not among the options (e.g. an imported font absent from the
        // ribbon's seeded list, or a size typed by hand): clear the highlighted
        // row and surface the raw value as editable text so the box still
        // reports the current font/size at the caret.
        dd.selectOption(-1);
        if (dd.control) dd.control.value = hostValue;
      }
    };
    requestAnimationFrame(apply);
  }

  #emit(): void {
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: { event: this.event, value: this.value, source: this },
      }),
    );
  }
}

customElements.define("docen-ribbon-combobox", DocenRibbonCombobox);

export default DocenRibbonCombobox;
