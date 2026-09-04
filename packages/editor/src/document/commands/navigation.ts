import type { Editor } from "@docen/docx/core";
import {
  findNext,
  findPrev,
  getMatchHighlights,
  replaceAll,
  replaceNext,
  setSearchState,
  SearchQuery,
} from "prosemirror-search";

import { t } from "../../ui";
import type { OutlineItem } from "../components/outline";

/** Build a nested OutlineItem tree from the flat outline anchor list: each
 *  heading nests under the nearest preceding heading with a smaller level. */
function buildOutlineTree(
  anchors: readonly { id: string; textContent: string; originalLevel: number }[],
): OutlineItem[] {
  type Node = { id: string; title: string; level: number; children?: Node[] };
  const roots: Node[] = [];
  const stack: Node[] = [];
  for (const a of anchors) {
    const node: Node = { id: a.id, title: a.textContent, level: a.originalLevel };
    while (stack.length && stack[stack.length - 1].level >= a.originalLevel) stack.pop();
    const parent = stack[stack.length - 1];
    if (parent) (parent.children ??= []).push(node);
    else roots.push(node);
    stack.push(node);
  }
  return roots as OutlineItem[];
}

/** The navigation commands' view of the host — resolved per call so the
 *  controller can be built before a document opens. */
export interface NavigationHost {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The story bridge — jumps scroll the target into view. */
  bridge(): { scrollIntoView(pos: number): void } | undefined;
  /** The host element — the shadow-DOM root for the panes, the Results slot,
   *  and the Find & Replace dialog; also the i18n language source. */
  element(): HTMLElement;
  /** The host's shared selection setter (the viewless editor's dual d.ts
   *  identity bridges there). */
  setTextSelection(from: number, to?: number): void;
}

/**
 * The navigation domain, split out of the host element: the outline pane
 * (anchors → tree → click-to-jump), the nav-pane search (live highlights,
 * the Results list, next/prev), and the Find & Replace dialog driving
 * prosemirror-search.
 */
export class NavigationCommands {
  constructor(private readonly host: NavigationHost) {}

  /** Outline anchors from the latest projection (an outline click resolves to
   *  a position through this list). */
  #anchors: readonly { id: string; pos: number; textContent: string; originalLevel: number }[] = [];

  /** Fingerprint of what the outline pane shows — skips fluent-tree rebuilds
   *  (and their flicker) when nothing visible changed. */
  #outlineSig = "";

  /** The debounced Results-list rebuild (one pending at a time). */
  #searchTimer?: ReturnType<typeof setTimeout>;

  /** Tear down pending timers (the host's disconnectedCallback). */
  dispose(): void {
    clearTimeout(this.#searchTimer);
  }

  /** Outline.onUpdate → <docen-outline>. Cache the anchors (so an
   *  outline click resolves to a position) and rebuild the nested tree. */
  renderOutline(
    anchors: readonly { id: string; pos: number; textContent: string; originalLevel: number }[],
  ): void {
    this.#anchors = anchors;
    const outline = this.host.element().shadowRoot?.querySelector("docen-outline");
    if (!outline) return;
    // Fingerprint only what the pane shows (id/level/title). `pos` moves on
    // every re-render but never changes the outline, so excluding it avoids
    // rebuilding — and flickering — the fluent tree each pass. Built from
    // per-anchor arrays rather than the serialized tree, so object key order
    // is irrelevant (no dependency on buildOutlineTree's literal field order,
    // unlike a plain JSON.stringify(tree) comparison).
    const sig = anchors
      .map((a) => JSON.stringify([a.id, a.originalLevel, a.textContent]))
      .join("\n");
    if (this.#outlineSig === sig) return;
    this.#outlineSig = sig;
    outline.setAttribute("items", JSON.stringify(buildOutlineTree(anchors)));
  }

  /** Outline click → select the heading at its position and scroll it into view. */
  readonly onOutlineSelect = (event: CustomEvent<{ id?: string }>): void => {
    const id = event.detail?.id;
    const bridge = this.host.bridge();
    if (!id || !bridge) return;
    const anchor = this.#anchors.find((a) => a.id === id);
    if (!anchor) return;
    this.host.setTextSelection(anchor.pos);
    bridge.scrollIntoView(anchor.pos);
  };

  /** navigation:search → set the active query; matches highlight live. */
  readonly onSearch = (event: CustomEvent<{ query?: string }>): void => {
    const editor = this.host.editor();
    if (!editor) return;
    const query = new SearchQuery({ search: event.detail?.query ?? "", caseSensitive: false });
    editor.view.dispatch(setSearchState(editor.state.tr, query));
    // Debounce the result-list rebuild (O(matches) DOM nodes per keystroke); the
    // query already dispatched above, so find-next reads the live search state.
    clearTimeout(this.#searchTimer);
    this.#searchTimer = setTimeout(() => this.#updateSearchResults(), 120);
  };

  /** navigation:find → jump to the next/previous match (prosemirror-search). */
  readonly onFind = (event: CustomEvent<{ direction: "next" | "prev" }>): void => {
    const editor = this.host.editor();
    if (!editor) return;
    (event.detail.direction === "prev" ? findPrev : findNext)(editor.state, editor.view.dispatch);
    this.host.bridge()?.scrollIntoView(editor.state.selection.from);
  };

  /** Stamp the Results slot with the live match list — each hit rendered with
   *  surrounding context and a data-from/to for click-to-jump (Word's Results
   *  pane lists every match with context, not just a count). */
  #updateSearchResults(): void {
    const editor = this.host.editor();
    const slot = this.host.element().shadowRoot?.querySelector(".search-results");
    if (!slot) return;
    const decos = editor ? getMatchHighlights(editor.state).find() : [];
    slot.replaceChildren();
    const header = document.createElement("div");
    header.className = "result-count";
    header.textContent =
      decos.length > 0
        ? `${decos.length} ${t("search.matches", this.host.element())}`
        : t("search.noResults", this.host.element());
    slot.append(header);
    if (!editor || decos.length === 0) return;
    const doc = editor.state.doc;
    const RADIUS = 24;
    for (const deco of decos) {
      const { from, to } = deco as { from: number; to: number };
      const before = doc.textBetween(Math.max(0, from - RADIUS), from, " ");
      const after = doc.textBetween(to, Math.min(doc.content.size, to + RADIUS), " ");
      const item = document.createElement("button");
      item.type = "button";
      item.className = "result-item";
      item.dataset.from = String(from);
      item.dataset.to = String(to);
      if (before) {
        const span = document.createElement("span");
        span.textContent = `…${before}`;
        item.append(span);
      }
      const hit = document.createElement("mark");
      hit.textContent = doc.textBetween(from, to, " ");
      item.append(hit);
      if (after) {
        const span = document.createElement("span");
        span.textContent = `${after}…`;
        item.append(span);
      }
      slot.append(item);
    }
  }

  /** Click a Results entry → select that match range and scroll it into view. */
  readonly onSearchResultClick = (event: Event): void => {
    const bridge = this.host.bridge();
    if (!bridge) return;
    const item = (event.target as HTMLElement | null)?.closest(".result-item");
    if (!(item instanceof HTMLElement)) return;
    const from = Number(item.dataset.from);
    const to = Number(item.dataset.to);
    if (!Number.isFinite(from) || !Number.isFinite(to)) return;
    this.host.setTextSelection(from, to);
    bridge.scrollIntoView(from);
  };

  /** Ctrl+F → open the nav pane and focus its search box (Word behavior). */
  openSearch(): void {
    const root = this.host.element().shadowRoot;
    const taskPane = root?.querySelector('docen-task-pane[position="start"]') as
      | (HTMLElement & { open: boolean })
      | null;
    if (taskPane) taskPane.open = true;
    const input = root
      ?.querySelector("docen-navigation-pane")
      ?.shadowRoot?.querySelector("[part='search-input']") as
      | (HTMLElement & { select: () => void })
      | null;
    input?.focus();
    input?.select?.();
  }

  /** Ctrl+H / ribbon Replace → open the Find & Replace dialog. */
  openFindReplace(): void {
    const dialog = this.host.element().shadowRoot?.querySelector("docen-find-replace-dialog") as
      | (HTMLElement & { show: () => void })
      | null;
    dialog?.show();
  }

  /** find-replace:action → drive prosemirror-search (query highlights, find-next,
   *  replace-next = replace + advance, replace-all). Each action re-stamps the
   *  query so Find/Replace/options are always current. */
  readonly onFindReplace = (
    event: CustomEvent<{
      action: string;
      find: string;
      replace: string;
      caseSensitive: boolean;
      wholeWord: boolean;
    }>,
  ): void => {
    const editor = this.host.editor();
    if (!editor) return;
    const { action, find, replace, caseSensitive, wholeWord } = event.detail ?? {};
    const query = new SearchQuery({ search: find, replace, caseSensitive, wholeWord });
    editor.view.dispatch(setSearchState(editor.state.tr, query));
    if (action === "find-next") findNext(editor.state, editor.view.dispatch);
    else if (action === "replace-next") replaceNext(editor.state, editor.view.dispatch);
    else if (action === "replace-all") replaceAll(editor.state, editor.view.dispatch);
    if (action === "find-next" || action === "replace-next") {
      this.host.bridge()?.scrollIntoView(editor.state.selection.from);
    }
  };
}
