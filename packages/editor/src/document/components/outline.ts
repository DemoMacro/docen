import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

import { observeLang, t } from "../../ui/i18n/localize";

/** A single outline (TOC) entry. `children` nest recursively. */
export interface OutlineItem {
  /** Stable id emitted on `outline:select`. */
  readonly id: string;
  /** Heading text (business string — translated by the editor package). */
  readonly title: string;
  /** Outline level (1-based); informational. */
  readonly level?: number;
  /** Target page number, emitted with `outline:select` for jump-to-page. */
  readonly page?: number;
  /** Nested entries, rendered as child tree items. */
  readonly children?: readonly OutlineItem[];
}

const styles = css`
  :host {
    display: block;
    font-size: 12px;
  }
  fluent-tree {
    padding: 8px;
    box-sizing: border-box;
  }
  /* fluent-tree-item's ::part(content) is a flex item with min-width:auto, so
     a long label stretches it past the pane (visible overflow). min-width:0
     lifts the no-shrink floor and the width cap bounds it to the pane; the
     label inside then truncates instead of overflowing. */
  fluent-tree-item::part(content) {
    min-width: 0;
    max-width: 100%;
    overflow: hidden;
  }
  /* The label sits as a flex item inside ::part(content); min-width:0 lets it
     shrink below its text, and the full text is reachable via the item title. */
  .outline-label {
    flex: 1 1 auto;
    min-width: 0;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .empty {
    font-size: 12px;
    color: #666;
    padding: 12px 8px;
    margin: 0;
  }
`;

const template = html<DocenOutline>`<div part="body" ${ref("body")}></div>`;

/**
 * `<docen-outline items='[{id, title, children:[…]}]'>` — an Office-style
 * outline/TOC pane backed by `<fluent-tree>`. Entries recurse via
 * `slot="item"` children; clicking one emits `outline:select` with
 * `{ id, page }`. Entry titles are business strings (the editor package
 * builds the outline and translates them); only the empty state is localized
 * here. Items default to expanded.
 */
@customElement({ name: "docen-outline", template, styles })
class DocenOutline extends FASTElement {
  @attr items?: string;
  @observable body?: HTMLElement;
  #unsubscribe?: () => void;

  itemsChanged(): void {
    this.#render();
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.#render();
    this.#unsubscribe = observeLang(() => {
      const empty = this.body?.querySelector(".empty");
      if (empty) empty.textContent = t("outline.empty", this);
    });
  }

  disconnectedCallback(): void {
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  // The JSON `items` attribute parsed to entries; bad JSON → empty (no render).
  get parsedItems(): OutlineItem[] {
    try {
      return JSON.parse(this.items ?? "[]") as OutlineItem[];
    } catch {
      return [];
    }
  }

  #render(): void {
    const body = this.body;
    if (!body) return;
    body.replaceChildren();
    const items = this.parsedItems;
    if (items.length === 0) {
      const empty = document.createElement("p");
      empty.className = "empty";
      empty.textContent = t("outline.empty", this);
      body.append(empty);
      return;
    }
    const tree = document.createElement("fluent-tree");
    for (const item of items) {
      tree.append(this.#renderItem(item));
    }
    // Event delegation: one listener on the tree resolves to the innermost
    // clicked item (closest), so a click emits exactly one select.
    tree.addEventListener("click", (event: Event) => {
      const target = event.target as HTMLElement;
      const itemEl = target.closest("fluent-tree-item") as HTMLElement | null;
      const id = itemEl?.dataset.id;
      if (id != null) this.#emit(id, itemEl!.dataset.page);
    });
    body.append(tree);
  }

  #renderItem(item: OutlineItem): HTMLElement {
    const el = document.createElement("fluent-tree-item");
    el.dataset.id = item.id;
    // title surfaces the full heading text on hover — essential when the label
    // is truncated, and it overrides the ancestor task-pane's "Navigation"
    // tooltip so hovering an entry shows its own heading.
    el.title = item.title;
    if (item.page != null) el.dataset.page = String(item.page);
    const label = document.createElement("span");
    label.className = "outline-label";
    label.textContent = item.title;
    el.append(label);
    if (item.children?.length) {
      el.setAttribute("expanded", "");
      for (const child of item.children) {
        const childEl = this.#renderItem(child);
        childEl.slot = "item";
        el.append(childEl);
      }
    }
    return el;
  }

  #emit(id: string, page?: string): void {
    this.dispatchEvent(
      new CustomEvent("outline:select", {
        bubbles: true,
        composed: true,
        detail: { id, page: page != null ? Number(page) : undefined, source: this },
      }),
    );
  }
}

export default DocenOutline;
