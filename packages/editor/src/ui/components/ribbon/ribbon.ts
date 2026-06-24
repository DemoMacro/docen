import { observeLang, t } from "../../i18n";

// Per-instance CSS anchor name so the display-options menu right-aligns to its
// own trigger (and does not collide with another ribbon instance's).
let seq = 0;

const template = document.createElement("template");
template.innerHTML = `
  <style>
    :host {
      display: flex;
      flex-direction: column;
      position: relative;
      background: var(--docen-color-bg, #fff);
      border-bottom: 1px solid var(--docen-color-divider, #e2e2e2);
      font-family: "Segoe UI", "Segoe UI Web (West European)", system-ui, sans-serif;
      font-size: var(--docen-font-size-ribbon, 12px);
      color: var(--docen-color-text, #444);
    }
    .rb-tabs {
      display: flex;
      align-items: flex-end;
      background: var(--docen-color-tab-bg, #f0f0f0);
    }
    .rb-tabs-start { flex: 0 0 auto; }
    /* Office-style tab-row trailing actions (Comment/Edit/Share…): pushed to the
       inline-end. The strip aligns tabs to the bottom (indicator meets the
       panel), so override with align-self:center to vertically center actions. */
    .rb-tabs-end {
      align-self: center;
      margin-inline-start: auto;
      display: flex;
      align-items: center;
      gap: 2px;
      padding-inline-end: 6px;
    }
    .rb-panels {
      background: var(--docen-color-bg, #fff);
      padding: 4px 6px 2px;
      /* Reserve room on the inline-end for the Ribbon Display Options chevron
         (pinned to the bottom-right) so it never overlaps the last group. */
      padding-inline-end: 30px;
      box-sizing: border-box;
      /* A wide tab (Layout/References/Mailings) overflows on narrow widths —
         let the command row scroll horizontally instead of wrapping/clipping.
         The split-button menus open in the top layer (popover), so they are not
         clipped by this scroll container. */
      overflow-x: auto;
      overflow-y: hidden;
    }
    /* Panel visibility is owned here (the parent), not by the panel itself.
       Avoids a host display rule inside the panel shadow overriding it; the
       active panel is flagged with the data-active attribute from #sync. The
       panel height follows its content (no forced height) so the command row
       is never left floating above empty space. */
    ::slotted(docen-ribbon-panel) {
      display: flex;
      flex-direction: row;
      flex-wrap: nowrap;
      align-items: stretch;
    }
    ::slotted(docen-ribbon-panel:not([data-active])) {
      display: none;
    }

    /* Ribbon Display Options — a small downward chevron pinned to the ribbon's
       bottom-right (Office behavior). It lives at the :host level, not inside
       .rb-panels, so it stays put while the panel scrolls horizontally and
       remains reachable when the panel is collapsed (tabs-only / auto-hide).
       The icon-only fluent-menu-button renders its own chevron. */
    .rb-display-options {
      position: absolute;
      right: 0;
      bottom: 0;
      z-index: 2;
    }
    /* tabs-only: collapse the command panel (keep tabs); a tab click pops it
       open (data-expanded) until an outside click closes it (#onDocClick). */
    :host([data-ribbon-mode="tabs-only"]) .rb-panels { display: none; }
    :host([data-ribbon-mode="tabs-only"][data-expanded]) .rb-panels { display: block; }
    /* auto-hide = Full Screen (Office "Auto-hide Ribbon / Full Screen Mode"):
       collapse the whole ribbon (tabs + commands) to a 3px transparent sliver
       so it occupies no space but stays hoverable; hovering reveals the full
       ribbon as an overlay (does not push the document). Browser fullscreen +
       status-bar hide is driven by the host (document) on ribbon-mode-change;
       this component only owns its own collapse. */
    :host([data-ribbon-mode="auto-hide"]) {
      position: absolute;
      top: 0;
      inset-inline: 0;
      z-index: 5;
      height: 3px;
      overflow: hidden;
      transition: height 0.12s ease;
    }
    :host([data-ribbon-mode="auto-hide"]:hover) {
      height: auto;
      overflow: visible;
      background: var(--docen-color-bg, #fff);
      border-bottom: 1px solid var(--docen-color-divider, #e2e2e2);
      box-shadow: 0 4px 8px rgba(0, 0, 0, 0.12);
    }
    /* Match the split-button caret: a tight icon-only menu-button, borderless —
       not Fluent's default 32px padded icon-only button. Subtle = borderless;
       hover paints the usual fill. */
    .rb-display-options fluent-menu-button {
      min-width: 0;
      width: 22px;
      max-width: none;
      min-height: 22px;
      padding-inline: 0;
    }
  </style>
  <div class="rb-tabs" part="tabs">
    <div class="rb-tabs-start" part="tabs-start"><slot name="tabs"></slot></div>
    <div class="rb-tabs-end" part="tabs-end"><slot name="actions"></slot></div>
  </div>
  <div class="rb-panels" part="panels"><slot></slot></div>
  <div class="rb-display-options" part="display-options">
    <fluent-menu part="do-menu">
      <fluent-menu-button id="do-trigger" slot="trigger" icon-only appearance="subtle" part="do-button"></fluent-menu-button>
      <fluent-menu-list popover part="do-list" class="rb-do-list"></fluent-menu-list>
    </fluent-menu>
  </div>
`;

/**
 * `<docen-ribbon>` — Office-style command surface. Layout only (gray tab strip
 * + white panel area); the tab row and its indicator come from a consumer-
 * provided `<fluent-tablist slot="tabs">` with `docen-ribbon-tab` children
 * (fluent `Tab` under a docen name).
 *
 * Listens to the slotted tablist's `change` and shows the matching
 * `<docen-ribbon-panel value>`, re-emitting `change` with `{ value }`.
 *
 * Re-binds itself whenever its tablist/panel children change (a host re-stamps
 * the ribbon innerHTML on a locale switch); without that the new panels would
 * carry no `data-active` and all hide, leaving the ribbon blank.
 *
 * The bottom-right chevron opens the Ribbon Display Options menu (auto-hide /
 * tabs-only / always). `data-ribbon-mode` on the host drives panel visibility:
 * always (default) keeps the panel shown; tabs-only hides it until a tab is
 * clicked (then `data-expanded` holds it open until an outside click); auto-hide
 * hides it until the ribbon is hovered.
 */
class DocenRibbon extends HTMLElement {
  readonly #anchorId = `--rb-do-${++seq}`;
  #observer?: MutationObserver;
  #unobserveLang?: () => void;
  #boundTablist?: HTMLElement | null;
  #boundTabs = new WeakSet<HTMLElement>();

  static get observedAttributes(): string[] {
    return ["data-ribbon-mode"];
  }

  /** Sync the display-options checkmark when `data-ribbon-mode` changes
   *  externally — e.g. the host resets it after the browser leaves fullscreen
   *  on Esc. (Idempotent: `#setMode` also calls `#updateCheck`, so a change
   *  originating there just runs it twice.) */
  attributeChangedCallback(): void {
    if (!this.shadowRoot) return;
    this.#updateCheck();
  }

  connectedCallback(): void {
    if (!this.shadowRoot) {
      this.attachShadow({ mode: "open" }).append(template.content.cloneNode(true));
    }
    if (!this.#observer) {
      this.#observer = new MutationObserver(() => this.#setup());
      this.#observer.observe(this, { childList: true });
    }
    // Right-align the menu to the trigger (Office behavior: the button sits at
    // the ribbon's far right, so the menu opens from the right edge under it —
    // left-aligned it would overflow off-screen). Inline styles win over
    // Fluent's ::slotted([popover]) positioning rules.
    const trigger = this.shadowRoot!.querySelector("#do-trigger") as HTMLElement | null;
    const list = this.shadowRoot!.querySelector(".rb-do-list") as HTMLElement | null;
    if (trigger && list) {
      trigger.style.anchorName = this.#anchorId;
      list.style.positionAnchor = this.#anchorId;
      list.style.insetInlineEnd = "anchor(self-end)";
      list.style.insetInlineStart = "unset";
    }
    this.#renderDisplayOptions();
    // Re-render the display-options items when the page locale changes. The
    // host re-stamps the ribbon's tabs/panels on a language switch, but this
    // <docen-ribbon> element itself stays connected, so without this the items
    // would keep their initial-language text while the rest of the ribbon
    // updates — matching the navigation-pane/outline/properties-panel pattern.
    this.#unobserveLang = observeLang(() => this.#renderDisplayOptions());
    document.addEventListener("click", this.#onDocClick, true);
    queueMicrotask(() => this.#setup());
  }

  disconnectedCallback(): void {
    this.#observer?.disconnect();
    this.#observer = undefined;
    this.#boundTablist = undefined;
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    document.removeEventListener("click", this.#onDocClick, true);
  }

  /** Collapse a tabs-only panel when the click lands outside the ribbon. */
  #onDocClick = (event: MouseEvent): void => {
    if (this.getAttribute("data-ribbon-mode") !== "tabs-only") return;
    if (!this.hasAttribute("data-expanded")) return;
    if (this.contains(event.target as Node)) return;
    this.removeAttribute("data-expanded");
  };

  #setup(): void {
    const tablist = this.querySelector("fluent-tablist") as HTMLElement | null;
    if (!tablist) return;
    // Bind `change` once per tablist instance — a re-stamp swaps the element,
    // so the reference check re-binds only the new one (no duplicate listeners
    // accumulating across locale switches).
    if (tablist !== this.#boundTablist) {
      // Swallow the tablist's own `change` so callers only receive docen-ribbon's.
      tablist.addEventListener("change", (event) => {
        event.stopPropagation();
        this.#sync(tablist);
      });
      this.#boundTablist = tablist;
    }
    // `docen-ribbon-tab` is a renamed fluent Tab; in this Fluent build its
    // native click doesn't advance `activeid`, so drive it manually. Setting
    // activeid still lets the tablist move its indicator and emit `change`.
    const driver = tablist as unknown as { activeid: string };
    this.querySelectorAll<HTMLElement>("docen-ribbon-tab").forEach((tab) => {
      if (this.#boundTabs.has(tab)) return;
      this.#boundTabs.add(tab);
      tab.addEventListener("click", () => {
        if (tab.id) driver.activeid = tab.id;
        // tabs-only: a tab click pops the command panel open (Office behavior)
        // — it stays open until the user clicks outside the ribbon.
        if (this.getAttribute("data-ribbon-mode") === "tabs-only") {
          this.setAttribute("data-expanded", "");
        }
      });
    });
    this.#sync(tablist);
  }

  #sync(tablist: HTMLElement): void {
    const active = (tablist as unknown as { activeid?: string }).activeid ?? "";
    this.querySelectorAll<HTMLElement>("docen-ribbon-panel").forEach((panel) => {
      panel.toggleAttribute("data-active", panel.getAttribute("value") === active);
    });
    this.dispatchEvent(new CustomEvent("change", { detail: { value: active } }));
  }

  /** Build the three localized Ribbon Display Options items. */
  #renderDisplayOptions(): void {
    const list = this.shadowRoot!.querySelector(".rb-do-list")!;
    const MODES = [
      { mode: "auto-hide", key: "ribbon.opt.auto-hide" },
      { mode: "tabs-only", key: "ribbon.opt.tabs-only" },
      { mode: "always", key: "ribbon.opt.always" },
    ] as const;
    list.replaceChildren();
    for (const { mode, key } of MODES) {
      const item = document.createElement("fluent-menu-item");
      item.setAttribute("role", "menuitemradio");
      // role="menuitemradio" carries a checkmark; the registry spans this item's
      // content from col 2 (leaving col 1 for the checkmark) so a long label
      // stretches the popover to fit instead of being clipped at the col-2 track.
      item.setAttribute("data-mode", mode);
      item.textContent = t(key);
      item.addEventListener("change", () => this.#setMode(mode));
      list.append(item);
    }
    this.#updateCheck();
    this.shadowRoot!.querySelector("#do-trigger")?.setAttribute(
      "aria-label",
      t("ribbon.opt.ribbon-display"),
    );
  }

  #setMode(mode: string): void {
    if (mode === "always") this.removeAttribute("data-ribbon-mode");
    else this.setAttribute("data-ribbon-mode", mode);
    this.removeAttribute("data-expanded");
    this.#updateCheck();
    this.dispatchEvent(new CustomEvent("ribbon-mode-change", { detail: { mode } }));
  }

  #updateCheck(): void {
    const current = this.getAttribute("data-ribbon-mode") ?? "always";
    this.shadowRoot!.querySelectorAll<HTMLElement>(".rb-do-list fluent-menu-item").forEach(
      (item) => {
        // fluent-menu-item renders its own checkmark for role="menuitemradio"
        // when the `checked` attribute is set — no custom ::before needed.
        item.toggleAttribute("checked", item.getAttribute("data-mode") === current);
      },
    );
  }
}

customElements.define("docen-ribbon", DocenRibbon);

export default DocenRibbon;
