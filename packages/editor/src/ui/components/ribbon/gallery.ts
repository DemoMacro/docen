import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

import { COMMAND_HOST_STYLE, renderIcon } from "./command-helpers";

/** One gallery entry — icon thumbnail over a short label (the compound
 *  button shape); `value` rides the emitted command detail. */
export interface RibbonGalleryItem {
  icon: string;
  text: string;
  value?: string;
  disabled?: boolean;
}

// Per-instance CSS anchor name so the drop-down gallery anchors to this
// strip, not the viewport corner.
let seq = 0;

const styles = css`
  ${COMMAND_HOST_STYLE}
  :host {
    display: inline-flex;
    align-items: stretch;
  }
  /* Word's Table Styles gallery strip: compact icon-over-label entries —
     the compound-button shape, ~56px to align with a large split's primary
     (a full 70px large button reads as a tall empty block). */
  .rb-gallery-strip {
    display: flex;
    gap: 2px;
  }
  .rb-gallery-item {
    appearance: none;
    -webkit-appearance: none;
    border: 1px solid transparent;
    background: transparent;
    box-sizing: border-box;
    width: 68px;
    padding: 3px 2px;
    margin: 0;
    cursor: pointer;
    border-radius: 4px;
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 2px;
    color: inherit;
    font: inherit;
  }
  .rb-gallery-item:hover {
    border-color: var(--docen-color-divider, #c7c7c7);
    background: var(--docen-color-hover, rgba(0, 0, 0, 0.04));
  }
  .rb-gallery-item[disabled] {
    opacity: 0.4;
    cursor: default;
  }
  .rb-gicon svg {
    display: block;
    width: 27px;
    height: 27px;
  }
  .rb-glabel {
    font-size: 10px;
    line-height: 1.15;
    text-align: center;
    overflow-wrap: break-word;
  }
  /* The More bar — a narrow full-height caret (Word's gallery expand). */
  button.rb-gallery-more {
    appearance: none;
    -webkit-appearance: none;
    border: none;
    background: transparent;
    cursor: pointer;
    min-width: 16px;
    width: 16px;
    padding: 0;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    border-radius: 2px;
    color: inherit;
  }
  button.rb-gallery-more::after {
    content: "";
    display: block;
    border-left: 3px solid transparent;
    border-right: 3px solid transparent;
    border-top: 4px solid currentColor;
  }
  button.rb-gallery-more:hover {
    background: var(--docen-color-hover, rgba(0, 0, 0, 0.06));
  }
  /* The expanded gallery — a grid of the same entries, anchored below the
     strip (Word's More gallery). No display here: the UA's
     [popover]:not(:popover-open) { display:none } must win until showPopover,
     an author display would keep it permanently visible. The inner grid div
     carries the layout instead. */
  .rb-gallery-pop {
    margin: 0;
    padding: 4px;
    background: var(--docen-color-bg, #fff);
    border: 1px solid var(--docen-color-divider, #c7c7c7);
    border-radius: 4px;
    box-shadow: 0 2px 12px rgba(0, 0, 0, 0.18);
    position-anchor: var(--rbg-anchor);
    inset-block-start: anchor(bottom);
    inset-inline-start: anchor(self-start);
    inset-inline-end: auto;
  }
  .rb-gallery-grid {
    display: grid;
    grid-template-columns: repeat(var(--rbg-columns, 3), 68px);
    gap: 2px;
  }
`;

const template = html<DocenRibbonGallery>`
  <div class="rb-gallery-strip" ${ref("strip")}></div>
  <button
    type="button"
    class="rb-gallery-more"
    part="more"
    aria-haspopup="true"
    aria-label="More"
    ?disabled="${(x) => x.disabled}"
    ${ref("more")}
  ></button>
  <div popover="auto" part="pop" class="rb-gallery-pop" ${ref("pop")}>
    <div class="rb-gallery-grid" ${ref("grid")}></div>
  </div>
`;

/**
 * `<docen-ribbon-gallery event="table-style" items='[{"icon","text","value"}]'>`
 * — Word's ribbon gallery control: the first `visible-count` entries as
 * icon-over-label thumbnails in a strip, then a narrow More bar whose
 * drop-down shows every entry in the same compound shape as a grid. Clicking
 * an entry (strip or drop-down) emits `command { event, value }`.
 */
@customElement({ name: "docen-ribbon-gallery", template, styles })
class DocenRibbonGallery extends FASTElement {
  @attr event?: string;
  @attr items?: string;
  @attr({ attribute: "visible-count" }) visibleCount?: string;
  @attr({ mode: "boolean" }) disabled?: boolean;

  @observable strip?: HTMLElement;
  @observable more?: HTMLElement;
  @observable pop?: HTMLElement;
  @observable grid?: HTMLElement;

  readonly anchorId = `--rbg-${++seq}`;

  get eventName(): string {
    return this.event || "";
  }
  get parsedItems(): RibbonGalleryItem[] {
    try {
      return JSON.parse(this.items ?? "[]") as RibbonGalleryItem[];
    } catch {
      return [];
    }
  }
  /** Strip entries (Word shows ~4); 4 is the default when unset/invalid. */
  get visible(): number {
    const n = Number(this.visibleCount);
    return Number.isFinite(n) && n > 0 ? n : 4;
  }

  itemsChanged(): void {
    this.#render();
  }
  visibleCountChanged(): void {
    this.#render();
  }

  connectedCallback(): void {
    super.connectedCallback();
    // Anchor the drop-down to the strip (same-shadow) — anchoring the host
    // crosses the shadow boundary and strands the popover at the corner.
    if (this.strip) this.strip.style.anchorName = this.anchorId;
    if (this.pop) this.pop.style.setProperty("--rbg-anchor", this.anchorId);
    this.#render();
    this.more?.addEventListener("click", this.onMoreClick);
  }

  disconnectedCallback(): void {
    this.more?.removeEventListener("click", this.onMoreClick);
    super.disconnectedCallback();
  }

  private readonly onMoreClick = (event: Event): void => {
    event.stopPropagation();
    if (this.disabled || !this.pop || this.pop.matches(":popover-open")) return;
    (this.pop as unknown as { showPopover?(): void }).showPopover?.();
  };

  #render(): void {
    if (!this.strip || !this.grid) return;
    const items = this.parsedItems;
    this.strip.replaceChildren(...items.slice(0, this.visible).map((item) => this.#entry(item)));
    this.grid.replaceChildren(...items.map((item) => this.#entry(item)));
  }

  #entry(item: RibbonGalleryItem): HTMLButtonElement {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "rb-gallery-item";
    if (item.disabled) btn.setAttribute("disabled", "");
    const icon = document.createElement("span");
    icon.className = "rb-gicon";
    renderIcon(icon, item.icon);
    const label = document.createElement("span");
    label.className = "rb-glabel";
    label.textContent = item.text;
    btn.append(icon, label);
    btn.addEventListener("click", () => {
      if (item.disabled) return;
      (this.pop as unknown as { hidePopover?(): void }).hidePopover?.();
      this.dispatchEvent(
        new CustomEvent("command", {
          bubbles: true,
          composed: true,
          detail: { event: this.eventName, value: item.value, source: this },
        }),
      );
    });
    return btn;
  }
}

export default DocenRibbonGallery;
