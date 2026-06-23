import { createIconSlot, renderIcon } from "./command-helpers";

const template = document.createElement("template");
template.innerHTML = `
  <style>
    :host {
      display: flex;
      flex-direction: column;
      gap: 2px;
      padding: 0 8px;
      position: relative;
      /* A group never shrinks below its content width — the Tables group holds
         a 72px large button and must not be compressed by its neighbors. */
      flex-shrink: 0;
      /* Fixed group height so every tab's ribbon is the same height. Without it
         a tab whose tallest group is a single large button (~88px) is shorter
         than one with a 3-row small-button grid (~109px), and the ribbon jumps
         when switching tabs. 112px fits the tallest group (Home); shorter ones
         stretch, and .rb-group-cmds { flex:1 } keeps their buttons centered. */
      min-height: 112px;
    }
    /* A vertical divider on every group's inline-end edge (including the last —
       Office groups are always separated, even when trailing space follows).
       Absolutely positioned so it spans the full group height without adding
       layout width. */
    [part="divider"] {
      position: absolute;
      inset-block: 0;
      inset-inline-end: 0;
      height: 100%;
      min-height: 0;
    }
    .rb-group-cmds {
      display: flex;
      flex-flow: row wrap;
      align-items: center;
      gap: 6px;
      flex: 1;
    }
    .rb-group-label {
      display: flex;
      align-items: center;
      gap: 4px;
      font-size: var(--docen-font-size-group-label, 10px);
      color: var(--docen-color-text-muted, #666);
      padding-top: 3px;
    }
    .rb-group-label-text { flex: 1; text-align: center; }
    .rb-launcher[hidden] { display: none; }
    /* Office dialog launcher: a small ⋯ at the group's inline-end. Fluent's
       :host([icon-only]) clamps the button to 32px; raise specificity past it. */
    :host([launcher]) .rb-launcher {
      flex: 0 0 auto;
      min-width: 22px;
      max-width: 22px;
      min-height: 22px;
      height: 22px;
      padding: 0;
    }
    /* Office dialog-box launcher: a small diagonal arrow (↘) at the group's
       inline-end. The icon is arrow_bidirectional_up_down tilted -45°. */
    .rb-launcher .rb-icon svg {
      display: block;
      fill: currentColor;
      width: 12px;
      height: 12px;
      transform: rotate(-45deg);
    }
  </style>
  <div class="rb-group-cmds" part="commands"><slot></slot></div>
  <div class="rb-group-label" part="label">
    <span class="rb-group-label-text" part="label-text"></span>
    <fluent-button class="rb-launcher" part="launcher" appearance="subtle" icon-only hidden></fluent-button>
  </div>
  <fluent-divider part="divider" orientation="vertical" role="separator" aria-orientation="vertical"></fluent-divider>`;

/**
 * `<docen-ribbon-group label="Clipboard" launcher="open-font-dialog">…</docen-ribbon-group>`
 * — a labeled command cluster inside a panel, visually separated by a right
 * hairline. The default slot holds command elements. A `launcher` attribute
 * (an event name) renders an Office-style **dialog box launcher** (⋯) at the
 * label row's inline-end; clicking it emits `command` with `{ event: launcher }`.
 * Layout only (no command styling beyond the launcher).
 */
class DocenRibbonGroup extends HTMLElement {
  static get observedAttributes(): string[] {
    return ["label", "launcher"];
  }

  #launcher?: HTMLElement;
  #icon?: HTMLSpanElement;

  attributeChangedCallback(name: string, _old: string, next: string): void {
    if (name === "label") this.#renderLabel(next);
    if (name === "launcher") this.#applyLauncher();
  }

  connectedCallback(): void {
    if (!this.shadowRoot) {
      this.attachShadow({ mode: "open" }).append(template.content.cloneNode(true));
    }
    this.#launcher = this.shadowRoot!.querySelector(".rb-launcher")!;
    this.#icon = createIconSlot("");
    this.#launcher.prepend(this.#icon);
    renderIcon(this.#icon, "launcher");
    this.#renderLabel(this.getAttribute("label") ?? "");
    this.#applyLauncher();
    this.#launcher.addEventListener("click", () => this.#emit());
  }

  get event(): string {
    return this.getAttribute("launcher") ?? "";
  }

  #applyLauncher(): void {
    if (!this.#launcher) return;
    const has = this.hasAttribute("launcher");
    this.#launcher.toggleAttribute("hidden", !has);
    if (has) this.#launcher.setAttribute("title", this.event);
  }

  #renderLabel(label: string): void {
    this.shadowRoot
      ?.querySelector(".rb-group-label-text")
      ?.replaceChildren(...(label ? [document.createTextNode(label)] : []));
  }

  #emit(): void {
    if (!this.hasAttribute("launcher")) return;
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: { event: this.event, source: this },
      }),
    );
  }
}

customElements.define("docen-ribbon-group", DocenRibbonGroup);

export default DocenRibbonGroup;
