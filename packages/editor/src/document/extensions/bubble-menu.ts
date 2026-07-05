import type { Editor } from "@docen/docx/core";
import { FASTElement, css, customElement, html, observable, repeat } from "@microsoft/fast-element";
import {
  BubbleMenu,
  BubbleMenuPlugin,
  type BubbleMenuOptions,
} from "@tiptap/extension-bubble-menu";
import { NodeSelection, PluginKey } from "prosemirror-state";

import type { BubbleButton } from "../../ui";
import { renderIcon } from "../../ui/components/ribbon/command-helpers";
import { observeLang, t } from "../../ui/i18n/localize";

/** The built-in bubble-menu buttons (Bold/Italic/Underline/Strike/Link/Clear).
 *  Returned by a function (mirrors `ribbonTabs()`) so the host combines them
 *  with addin-contributed buttons at boot — `[...defaultBubbleButtons(),
 *  ...host.mergedBubbleMenu()]`. `label` is an i18n key resolved at render
 *  time, so a locale switch re-renders without the host re-configuring. */
export function defaultBubbleButtons(): readonly BubbleButton[] {
  return [
    { event: "bold", icon: "bold", label: "ribbon.cmd.bold", activeMark: "bold" },
    { event: "italic", icon: "italic", label: "ribbon.cmd.italic", activeMark: "italic" },
    {
      event: "underline",
      icon: "underline",
      label: "ribbon.cmd.underline",
      activeMark: "underline",
    },
    { event: "strike", icon: "strike", label: "ribbon.cmd.strike", activeMark: "strike" },
    { event: "link", icon: "link", label: "ribbon.cmd.link", activeMark: "link" },
    { event: "clear-format", icon: "clear-format", label: "ribbon.cmd.clear-format" },
  ];
}

const styles = css`
  :host {
    display: flex;
    /* BubbleMenu's then-callback sets width:max-content inline, but the FIRST
       computePosition (run synchronously inside show()) reads width BEFORE
       that callback fires. At width:auto the flex parent (view.dom.parentElement)
       stretches the bar to ~its own width, so computePosition centers against
       that stretched width and the bar lands off-canvas until the next update.
       Setting it here makes the first read already see the real content width. */
    width: max-content;
    gap: 2px;
    padding: 2px;
    background: var(--docen-color-bg, #fff);
    border: 1px solid var(--docen-color-divider, #e2e2e2);
    border-radius: 4px;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.14);
    font-family: "Segoe UI", system-ui, sans-serif;
  }
  /* BubbleMenu.show() flips inline visibility/opacity the instant it appends
     this element, but updatePosition() is async (computePosition returns a
     Promise) — between append and the resolve, the bar would flash at the
     append target's default top-left spot. data-placed is set in onUpdate
     (fired inside the resolve) and cleared in onShow, so the first frame of
     each appearance stays hidden. !important beats the inline opacity set by
     show(). */
  :host(:not([data-placed])) {
    opacity: 0 !important;
  }
  fluent-button {
    min-height: 26px;
    min-width: 26px;
    padding: 0;
  }
  fluent-button[aria-pressed="true"] {
    background: var(--docen-color-hover, #e5e5e5);
    color: var(--docen-color-accent, #1064be);
  }
  .rb-icon {
    display: contents;
  }
  .rb-icon svg {
    display: block;
    fill: currentColor;
    width: 16px;
    height: 16px;
  }
`;

const template = html<DocenBubbleMenuBar>`
  ${repeat(
    (x: DocenBubbleMenuBar) => x.commands,
    html<BubbleButton>`
      <fluent-button
        id="${(cmd: BubbleButton) => cmd.event}"
        appearance="subtle"
        data-cmd="${(cmd: BubbleButton) => cmd.event}"
      >
        <span class="rb-icon"></span>
      </fluent-button>
      <fluent-tooltip
        anchor="${(cmd: BubbleButton) => cmd.event}"
        positioning="top"
      ></fluent-tooltip>
    `,
  )}
`;

/**
 * `<docen-bubble-menu>` — the toolbar element Tiptap's BubbleMenu positions
 * above the selection. Each button dispatches a `command` CustomEvent exactly
 * like a ribbon button, so the host's `#onCommand` routes it to
 * `editor.chain().focus()[event].run()` (no editor reference held here — the
 * element is created before the editor exists). Listeners are delegated to the
 * host so a `commands` re-render (locale switch) doesn't lose bindings, and
 * icons are re-injected in `commandsChanged`.
 */
@customElement({ name: "docen-bubble-menu", template, styles })
class DocenBubbleMenuBar extends FASTElement {
  /** The button list, set by the extension from `options.commands` (the host's
   *  merge of `defaultBubbleButtons()` + `mergedBubbleMenu()`). Empty until the
   *  extension assigns it in `addProseMirrorPlugins` — the bar is hidden by
   *  `:host(:not([data-placed]))` until then, so the empty state never shows. */
  @observable commands: readonly BubbleButton[] = [];

  /** Set by the extension in addProseMirrorPlugins — the bar needs editor
   *  access to mirror editor.isActive into aria-pressed, including on its
   *  first connect (which happens inside BubbleMenu.show(), before
   *  onTransaction has run with a connected bar). */
  editor?: Editor;

  #unsubscribe?: () => void;

  /** Dispatch a `command` event the host's #onCommand routes to the editor. */
  dispatch(event: string): void {
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: { event },
      }),
    );
  }

  /** Mirror editor.isActive into each toggle button's aria-pressed. Safe to
   *  call when disconnected (the buttons live in shadowRoot either way); a
   *  no-op when the template hasn't rendered yet (querySelector returns null). */
  syncActive(): void {
    if (!this.editor) return;
    for (const cmd of this.commands) {
      if (!cmd.activeMark) continue;
      const btn = this.shadowRoot?.querySelector(`fluent-button[data-cmd="${cmd.event}"]`);
      btn?.setAttribute("aria-pressed", String(this.editor.isActive(cmd.activeMark)));
    }
  }

  connectedCallback(): void {
    super.connectedCallback();
    // Delegate on shadowRoot (not host) so event.target stays the original
    // fluent-button — a host-level listener would see a retargeted target (the
    // host itself) and closest('fluent-button') would miss. mousedown uses
    // capture to beat the editor's blur (same pattern as ribbon-button).
    this.shadowRoot?.addEventListener("mousedown", onMousedown, { capture: true });
    this.shadowRoot?.addEventListener("click", this.#onClick);
    queueMicrotask(() => {
      this.#injectIcons();
      this.#applyI18n();
      // Sync after the repeat template has rendered the buttons — the show()
      // path can connect the bar before any onTransaction has run with it.
      this.syncActive();
    });
    this.#unsubscribe = observeLang(() => this.#applyI18n());
  }

  disconnectedCallback(): void {
    this.shadowRoot?.removeEventListener("mousedown", onMousedown, { capture: true });
    this.shadowRoot?.removeEventListener("click", this.#onClick);
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  commandsChanged(): void {
    queueMicrotask(() => {
      this.#injectIcons();
      this.#applyI18n();
    });
  }

  #onClick = (event: Event): void => {
    const btn = (event.target as HTMLElement | null)?.closest<HTMLElement>(
      "fluent-button[data-cmd]",
    );
    const cmd = btn?.dataset.cmd;
    if (cmd) this.dispatch(cmd);
  };

  /** Inject the Office icon svg into each button's slot. renderIcon uses
   *  replaceChildren, so re-runs (e.g. after a locale-driven template
   *  re-render) don't accumulate svgs. The svg slightly widens each button, so
   *  after injecting we dispatch an `updatePosition` meta: BubbleMenuView's
   *  transactionHandler (reachable because the extension shares BUBBLE_KEY with
   *  the plugin) re-runs computePosition against the final width. The first
   *  computePosition (run synchronously in show(), before this microtask) may
   *  measure the pre-icon width; this dispatch corrects it before paint. */
  #injectIcons(): void {
    const buttons = this.shadowRoot?.querySelectorAll("fluent-button[data-cmd]") ?? [];
    buttons.forEach((btn, i) => {
      const cmd = this.commands[i];
      if (!cmd) return;
      const slot = btn.querySelector<HTMLElement>(".rb-icon");
      if (slot) renderIcon(slot, cmd.icon);
    });
    const view = this.editor?.view;
    view?.dispatch(view.state.tr.setMeta(BUBBLE_KEY, "updatePosition"));
  }

  /** Translate each button's aria-label + tooltip for the active locale.
   *  Imperative — FASTElement's `repeat` skips rebinding when `commands` is
   *  re-assigned with the same element references, so a locale switch via
   *  observeLang wouldn't otherwise re-run a `t()` binding. Resolving from
   *  `this` lets `<docen-workspace lang>` override `<html lang>` (same pattern
   *  as the ribbon color-picker). */
  #applyI18n(): void {
    for (const cmd of this.commands) {
      const btn = this.shadowRoot?.querySelector(`fluent-button[data-cmd="${cmd.event}"]`);
      const tip = this.shadowRoot?.querySelector(`fluent-tooltip[anchor="${cmd.event}"]`);
      const label = cmd.label ? t(cmd.label, this) : "";
      if (btn) {
        if (label) btn.setAttribute("aria-label", label);
        else btn.removeAttribute("aria-label");
      }
      if (tip && label) tip.textContent = label;
    }
  }
}

/** mousedown preventDefault (capture) keeps the contenteditable selection
 *  alive while the user clicks a bubble button — same trick as the ribbon. */
const onMousedown = (event: Event): void => event.preventDefault();

/** The BubbleMenu plugin key. Passed to the extension as `pluginKey` (a
 *  PluginKey instance, not a string) so BubbleMenuView holds this same instance
 *  and `transactionHandler` reads meta set with it — `tr.setMeta(BUBBLE_KEY,
 *  "updatePosition")` reaches the handler and re-runs computePosition
 *  immediately (no updateDelay debounce). Two `new PluginKey("docenBubbleMenu")`
 *  would resolve to different keys ("…$" vs "…$1" — ProseMirror appends a
 *  counter in createKey), so the instance must be shared between the extension
 *  config and any code dispatching meta. */
const BUBBLE_KEY = new PluginKey("docenBubbleMenu");

// Created lazily in addProseMirrorPlugins, not at module top-level. The
// @customElement decorator runs as part of the class declaration, but a
// module-top `const x = createElement(name)` can resolve before that
// declaration executes (module evaluation order vs. class TDZ), yielding an
// HTMLUnknownElement that misses syncActive — and the extension would hold
// that wrong reference. By the time the editor loads the extension the class
// is guaranteed registered, so createElement returns a real DocenBubbleMenuBar.
let bubbleBar: DocenBubbleMenuBar | null = null;
const ensureBubbleBar = (): DocenBubbleMenuBar => {
  if (!bubbleBar) {
    bubbleBar = document.createElement("docen-bubble-menu") as DocenBubbleMenuBar;
  }
  return bubbleBar;
};

/** The singleton bubble bar element — null until the extension first creates
 *  it in `addProseMirrorPlugins`. Exported so the host can re-merge buttons at
 *  runtime (`addinsChanged`) without rebuilding the BubbleMenu plugin: the
 *  bar's `commands` is `@observable`, so a re-assignment re-renders the row
 *  and re-injects icons. Symmetric to ribbon's runtime re-render. */
export function getBubbleBar(): DocenBubbleMenuBar | null {
  return bubbleBar;
}

/** Extension options — {@link BubbleMenuOptions} (plugin config) plus `commands`,
 *  the merged button list the host assembles from `defaultBubbleButtons()` +
 *  `mergedBubbleMenu()` at boot. */
type DocenBubbleMenuOptions = BubbleMenuOptions & {
  commands: readonly BubbleButton[];
};

/** A floating format toolbar on the selection — Tiptap BubbleMenu wrapped so
 *  each button dispatches a `command` event the host already routes, and
 *  aria-pressed syncs with editor.isActive. The host configures `commands`
 *  (built-in defaults via `defaultBubbleButtons()` + addin contributions via
 *  `mergeBubbleMenu`), symmetric to the ribbon's `ribbonTabs()` +
 *  `mergeRibbonSchema`. */
export const DocenBubbleMenu = BubbleMenu.extend<DocenBubbleMenuOptions>({
  addOptions(): DocenBubbleMenuOptions {
    return {
      element: null,
      pluginKey: BUBBLE_KEY,
      shouldShow: ({ editor, view, from, to }) => {
        if (!editor.isEditable || from === to) return false;
        if (view.state.selection instanceof NodeSelection) return false;
        return true;
      },
      options: {
        placement: "top",
        strategy: "fixed",
        offset: 8,
        // show() clears the placed flag (re-entering the hidden-until-positioned
        // state); updatePosition's then-callback sets it once the bar is placed.
        onShow: () => bubbleBar?.removeAttribute("data-placed"),
        onUpdate: () => bubbleBar?.setAttribute("data-placed", ""),
      },
      commands: [] as readonly BubbleButton[],
    };
  },
  addProseMirrorPlugins() {
    const bar = ensureBubbleBar();
    bar.editor = this.editor;
    bar.commands = this.options.commands;
    // Pre-warm: connect the bar now so fluent-button upgrades + icon injection
    // finish before the first show() runs computePosition (otherwise it reads
    // the pre-icon width and centers the bar off-canvas). Stays hidden via
    // :host(:not([data-placed])); show()'s later appendChild is a no-op move.
    if (!bar.isConnected) {
      const container =
        (typeof this.options.appendTo === "function"
          ? this.options.appendTo()
          : this.options.appendTo) ?? this.editor.view.dom.parentElement;
      container?.appendChild(bar);
    }
    // Call BubbleMenuPlugin directly (the parent's addProseMirrorPlugins just
    // guards on options.element and forwards here) — this.parent()'s type
    // isn't declared on this hook, and we need to set bar.editor anyway.
    return [
      BubbleMenuPlugin({
        pluginKey: this.options.pluginKey,
        editor: this.editor,
        element: bar,
        updateDelay: this.options.updateDelay,
        options: this.options.options,
        appendTo: this.options.appendTo,
        shouldShow: this.options.shouldShow,
      }),
    ];
  },
  onTransaction() {
    bubbleBar?.syncActive();
  },
});
