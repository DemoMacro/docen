import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import type { RibbonControlOrLayout, RibbonTab } from "../../addin/types";
import { observeLang, t } from "../../i18n/localize";
import { ribbonIcon } from "../ribbon/icons";

/** A flattened, searchable command entry. */
interface CommandItem {
  label: string;
  event: string;
  value?: string;
  /** Owning tab label — shown as a source hint, Office-style. */
  tab: string;
  icon?: string;
}

const MAX_RESULTS = 12;

/** Flatten `Tab > Group > Control(+Layout)` into a flat command list.
 *  Menu/split items expand to their own entries (each carries its own event or
 *  inherits the control's). Deduped by event+value — the same command may have
 *  several ribbon entries. `separator` and disabled nodes are skipped.
 *
 *  Labels and tab names in the ribbon schema are i18n KEYS (ribbon.tab.* /
 *  ribbon.cmd.* / ribbon.opt.*); they are translated for `scope`'s locale so
 *  the list shows localized text. Re-run on lang change (see observeLang). */
function flattenCommands(tabs: readonly RibbonTab[], scope: Element | null): CommandItem[] {
  const out: CommandItem[] = [];
  const seen = new Set<string>();
  const push = (item: CommandItem): void => {
    if (!item.label || !item.event) return;
    const key = `${item.event}|${item.value ?? ""}`;
    if (seen.has(key)) return;
    seen.add(key);
    out.push(item);
  };
  const walk = (nodes: readonly RibbonControlOrLayout[], tab: string): void => {
    const tabLabel = t(tab, scope);
    for (const node of nodes) {
      if (node.type === "layout") {
        walk(node.controls, tab);
        continue;
      }
      if (node.type === "separator") continue;
      if (node.label && node.event && !node.disabled) {
        push({ label: t(node.label, scope), event: node.event, tab: tabLabel, icon: node.icon });
      }
      if ((node.type === "menu" || node.type === "split") && node.items) {
        for (const it of node.items) {
          if (it.disabled) continue;
          const ev = it.event ?? node.event;
          if (it.text && ev) {
            push({
              label: t(it.text, scope),
              event: ev,
              value: it.value,
              tab: tabLabel,
              icon: node.icon,
            });
          }
        }
      }
    }
  };
  for (const tab of tabs) {
    for (const group of tab.groups) walk(group.controls, tab.label);
  }
  return out;
}

/** Case-insensitive substring filter on label + tab. Empty query returns the
 *  first MAX_RESULTS (Office shows top commands on focus, before any typing). */
function filterCommands(all: readonly CommandItem[], query: string): CommandItem[] {
  const q = query.trim().toLowerCase();
  if (!q) return all.slice(0, MAX_RESULTS);
  return all
    .filter((c) => c.label.toLowerCase().includes(q) || c.tab.toLowerCase().includes(q))
    .slice(0, MAX_RESULTS);
}

const styles = css`
  :host {
    display: block;
    position: relative;
    width: 100%;
  }
  fluent-text-input {
    width: 100%;
    box-sizing: border-box;
  }
  .popover {
    position: fixed;
    background: var(--docen-color-bg, #fff);
    border: 1px solid var(--docen-color-divider, #e2e2e2);
    border-radius: 4px;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.14);
    max-height: 320px;
    overflow: auto;
    padding: 4px;
    box-sizing: border-box;
    z-index: 1000;
    font-size: var(--docen-font-size-ribbon, 12px);
    font-family: "Segoe UI", system-ui, sans-serif;
    color: var(--docen-color-text, #444);
  }
  .popover[hidden] {
    display: none;
  }
  .item {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 6px 8px;
    border-radius: 3px;
    cursor: pointer;
  }
  .item:hover,
  .item[aria-selected="true"] {
    background: var(--docen-color-hover, #f0f0f0);
  }
  .item .icon {
    display: contents;
  }
  .item .icon svg {
    display: block;
    width: 16px;
    height: 16px;
    fill: currentColor;
  }
  .item .label {
    flex: 1 1 auto;
    min-width: 0;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .item .tab {
    flex: 0 0 auto;
    color: var(--docen-color-text-muted, #888);
    font-size: 11px;
  }
  .empty {
    padding: 12px 8px;
    color: var(--docen-color-text-muted, #888);
    text-align: center;
  }
`;

const template = html<DocenCommandSearch>`
  <fluent-text-input
    ${ref("input")}
    placeholder="${(x) => x.placeholderText}"
    autocomplete="off"
    role="searchbox"
  ></fluent-text-input>
  <div class="popover" ${ref("listEl")} hidden></div>
`;

/**
 * `<docen-command-search>` — Office's "Tell me what you want to do" box: type a
 * command name, pick from the filtered list, and the chosen command dispatches
 * exactly as if its ribbon button were clicked (same `command` CustomEvent the
 * host's `#onCommand` already routes). The host pushes the full ribbon schema
 * (built-in tabs + addin contributions) via {@link setTabs}; the component
 * flattens it into a searchable index.
 *
 * UI: `fluent-text-input` (reused) + a self-built popover list. Not a
 * `fluent-dropdown type="combobox"` — that component's semantics are "pick a
 * value to keep", whereas a command palette fires-and-clears on select.
 */
@customElement({ name: "docen-command-search", template, styles })
class DocenCommandSearch extends FASTElement {
  @observable input?: HTMLElement;
  @observable listEl?: HTMLDivElement;
  @observable placeholderText = "";

  #commands: CommandItem[] = [];
  #results: CommandItem[] = [];
  #activeIndex = -1;
  #unsubscribe?: () => void;

  /** Host pushes the merged ribbon schema (built-in + addin tabs), with the
   *  i18n scope the labels resolve against — the workspace, so `<docen-workspace
   *  lang>` drives the locale (same scope the ribbon uses; passing it explicitly
   *  avoids relying on `this.closest('docen-workspace')` which is null for a
   *  transient moment while the host re-stamps the title-bar). The host re-calls
   *  this on lang change, so the list re-translates with the new locale.
   *  Re-runs the active filter so a schema change mid-typing stays consistent. */
  setTabs(tabs: readonly RibbonTab[], scope: Element | null = this): void {
    this.#commands = flattenCommands(tabs, scope);
    if (this.#isOpen()) this.#runFilter(this.#query());
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.placeholderText = t("header.search", this);
    this.addEventListener("input", this.#onInput);
    this.addEventListener("keydown", this.#onKeyDown);
    this.addEventListener("focusin", this.#onFocusIn);
    this.addEventListener("focusout", this.#onFocusOut);
    // mousedown-default on the popover so clicking an item does NOT blur the
    // input first (otherwise the commit click would race the close-on-blur).
    this.listEl?.addEventListener("mousedown", this.#onPopoverMouseDown);
    this.listEl?.addEventListener("click", this.#onPopoverClick);
    this.#unsubscribe = observeLang(() => {
      this.placeholderText = t("header.search", this);
    });
  }

  disconnectedCallback(): void {
    this.removeEventListener("input", this.#onInput);
    this.removeEventListener("keydown", this.#onKeyDown);
    this.removeEventListener("focusin", this.#onFocusIn);
    this.removeEventListener("focusout", this.#onFocusOut);
    this.listEl?.removeEventListener("mousedown", this.#onPopoverMouseDown);
    this.listEl?.removeEventListener("click", this.#onPopoverClick);
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  /** Office's Alt+Q lands here. Forwards to the inner <input>. */
  override focus(): void {
    this.#innerInput()?.focus();
  }

  // ── helpers ───────────────────────────────────────────────────────────────

  /** fluent-text-input hosts an <input> in its shadow tree. */
  #innerInput(): HTMLInputElement | null {
    const sr = (this.input as (HTMLElement & { shadowRoot?: ShadowRoot }) | undefined)?.shadowRoot;
    return sr?.querySelector("input") ?? null;
  }

  #query(): string {
    return this.#innerInput()?.value ?? "";
  }

  #isOpen(): boolean {
    return !!this.listEl && !this.listEl.hidden;
  }

  // ── event handlers ─────────────────────────────────────────────────────────

  #onInput = (): void => {
    this.#runFilter(this.#query());
  };

  #onKeyDown = (event: KeyboardEvent): void => {
    if (event.key === "ArrowDown") {
      event.preventDefault();
      if (!this.#isOpen()) this.#runFilter(this.#query());
      this.#move(1);
    } else if (event.key === "ArrowUp") {
      event.preventDefault();
      if (!this.#isOpen()) this.#runFilter(this.#query());
      this.#move(-1);
    } else if (event.key === "Enter") {
      const cmd = this.#results[this.#activeIndex];
      if (cmd) {
        event.preventDefault();
        this.#commit(cmd);
      }
    } else if (event.key === "Escape") {
      if (this.#isOpen()) event.preventDefault();
      this.#close();
    }
  };

  #onFocusIn = (): void => {
    if (!this.#isOpen()) this.#runFilter(this.#query());
  };

  #onFocusOut = (event: FocusEvent): void => {
    // Stay open while focus moves within this host (e.g. onto the popover).
    if (event.relatedTarget && this.contains(event.relatedTarget as Node)) return;
    this.#close();
  };

  #onPopoverMouseDown = (event: MouseEvent): void => {
    // Prevent the input from blurring when an item is clicked — the click then
    // fires normally and commits.
    event.preventDefault();
  };

  #onPopoverClick = (event: MouseEvent): void => {
    const item = (event.target as HTMLElement | null)?.closest<HTMLElement>("[data-idx]");
    if (!item) return;
    const idx = Number(item.dataset.idx ?? -1);
    const cmd = this.#results[idx];
    if (cmd) this.#commit(cmd);
  };

  // ── list state ──────────────────────────────────────────────────────────────

  #move(delta: number): void {
    if (this.#results.length === 0) return;
    const n = this.#results.length;
    let i = this.#activeIndex + delta;
    if (i < 0) i = n - 1;
    if (i >= n) i = 0;
    this.#activeIndex = i;
    this.#renderActive();
    this.#scrollActive();
  }

  #runFilter(query: string): void {
    this.#results = filterCommands(this.#commands, query);
    this.#activeIndex = this.#results.length > 0 ? 0 : -1;
    this.#renderList();
    this.#open();
  }

  #renderList(): void {
    const popover = this.listEl;
    if (!popover) return;
    popover.replaceChildren();
    if (this.#results.length === 0) {
      const empty = document.createElement("div");
      empty.className = "empty";
      empty.textContent = t("header.search.no-results", this);
      popover.append(empty);
      return;
    }
    this.#results.forEach((cmd, idx) => {
      const item = document.createElement("div");
      item.className = "item";
      item.dataset.idx = String(idx);
      item.setAttribute("role", "option");
      if (idx === this.#activeIndex) item.setAttribute("aria-selected", "true");
      if (cmd.icon) {
        const slot = document.createElement("span");
        slot.className = "icon";
        slot.innerHTML = ribbonIcon(cmd.icon) ?? "";
        item.append(slot);
      }
      const label = document.createElement("span");
      label.className = "label";
      label.textContent = cmd.label;
      item.append(label);
      const tab = document.createElement("span");
      tab.className = "tab";
      tab.textContent = cmd.tab;
      item.append(tab);
      popover.append(item);
    });
  }

  #renderActive(): void {
    this.listEl?.querySelectorAll(".item").forEach((el, i) => {
      el.setAttribute("aria-selected", i === this.#activeIndex ? "true" : "false");
    });
  }

  #scrollActive(): void {
    const el = this.listEl?.querySelector<HTMLElement>(`.item[data-idx="${this.#activeIndex}"]`);
    el?.scrollIntoView({ block: "nearest" });
  }

  // ── popover geometry ─────────────────────────────────────────────────────────

  #open(): void {
    if (!this.listEl) return;
    this.listEl.hidden = false;
    this.#positionPopover();
  }

  #close(): void {
    if (this.listEl) this.listEl.hidden = true;
  }

  #positionPopover(): void {
    if (!this.listEl || !this.input) return;
    const rect = this.input.getBoundingClientRect();
    const popover = this.listEl;
    popover.style.top = `${rect.bottom + 2}px`;
    popover.style.left = `${rect.left}px`;
    popover.style.width = `${Math.max(rect.width, 280)}px`;
  }

  // ── commit ───────────────────────────────────────────────────────────────────

  #commit(cmd: CommandItem): void {
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: { event: cmd.event, value: cmd.value },
      }),
    );
    // Office dismisses the box after a command runs.
    const input = this.#innerInput();
    if (input) input.value = "";
    this.#close();
  }
}

export default DocenCommandSearch;
