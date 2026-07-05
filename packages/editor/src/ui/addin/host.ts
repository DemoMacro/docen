import { FASTElement } from "@microsoft/fast-element";

import { registerLocalization } from "../i18n";
import type { DocenAddin, DocenHost, MiniToolbarButton, RibbonTab } from "./types";

/**
 * Pure merge of every addin's ribbon contributions into the tab order.
 *
 * Tabs appear in addin-registration order, then contribution order. A
 * contribution targeting an existing tab id appends its groups to that tab
 * (`label` is ignored once the tab exists); a fresh id creates a new tab.
 * Exported separately so the merge is testable without a live HTMLElement.
 */
export function mergeRibbonSchema<THost extends DocenHost>(
  addins: readonly DocenAddin<THost>[],
): RibbonTab[] {
  const tabs: RibbonTab[] = [];
  const index = new Map<string, RibbonTab>();
  for (const addin of addins) {
    for (const contribution of addin.ribbon ?? []) {
      let tab = index.get(contribution.tab);
      if (!tab) {
        tab = {
          id: contribution.tab,
          label: contribution.label ?? contribution.tab,
          groups: [],
        };
        index.set(contribution.tab, tab);
        tabs.push(tab);
      }
      tab.groups.push(...contribution.groups);
    }
  }
  return tabs;
}

/**
 * Pure merge of every addin's mini-toolbar buttons into a flat list, in
 * addin-registration then contribution order. Unlike ribbon (tab/group
 * structure) the mini toolbar is a single flat row, so this is a flatten — an
 * addin that wants its own variant of a default button just contributes
 * another entry (both appear). Exported separately so the merge is testable
 * without a live HTMLElement.
 */
export function mergeMiniToolbar<THost extends DocenHost>(
  addins: readonly DocenAddin<THost>[],
): MiniToolbarButton[] {
  const buttons: MiniToolbarButton[] = [];
  for (const addin of addins) {
    if (addin.miniToolbar) buttons.push(...addin.miniToolbar);
  }
  return buttons;
}

/**
 * Base class for docen editor hosts that load add-ins.
 *
 * Subclasses (e.g. `<docen-document>`'s `DocumentHost`) override `editor`,
 * `getContent`, and `setContent`; this base owns the addin registry, the ribbon
 * schema merge, and command routing. Injecting the merged ribbon/panes into the
 * shadow DOM happens in a subclass override of {@link addinsChanged}, invoked
 * whenever the addin set changes.
 *
 * `DocenAddin<this>` lets each addin's commands/panes receive the concrete host
 * subtype, so a `DocumentAddin` command can call `DocumentHost`-specific methods.
 */
export class AddinHost<TEditor = unknown> extends FASTElement implements DocenHost<TEditor> {
  #addins: DocenAddin<this>[] = [];

  /** Registered add-ins, in registration order. */
  get addins(): readonly DocenAddin<this>[] {
    return this.#addins;
  }

  // ── DocenHost surface (subclasses override the editor/content accessors) ──

  get element(): HTMLElement {
    return this;
  }

  /** The underlying editor (Tiptap `Editor` for `<docen-document>`). `undefined`
   *  until the editor mounts, or always for an editor-agnostic host. */
  get editor(): TEditor | undefined {
    return undefined;
  }

  /** The host's UI locale (Office.js `Office.context.displayLanguage`).
   *  Mirrors {@link resolveLang}: `lang` attribute > `<html lang>` > "en".
   *  `this.lang` is `HTMLElement.lang` (reflects the `lang` attribute), so a
   *  consumer setting `<docen-document lang="…">` — or the host forwarding the
   *  status-bar toggle to its own `lang` — takes effect here immediately. */
  get displayLanguage(): string {
    return this.lang || document.documentElement.lang || "en";
  }

  getContent(): unknown {
    return undefined;
  }

  setContent(_content: unknown): void {
    // subclasses override
  }

  // ── Addin registry ──────────────────────────────────────────────────────

  /** Register an add-in (idempotent on `addin.id`). Triggers {@link addinsChanged}. */
  addAddin(addin: DocenAddin<this>): void {
    if (this.#addins.some((existing) => existing.id === addin.id)) return;
    // Register the addin's localization manifest before its UI renders —
    // defaultLanguageTag becomes the fallback, additionalLanguages merge into
    // the global table so its label keys resolve on the next t() call.
    if (addin.localizationInfo) registerLocalization(addin.localizationInfo);
    this.#addins = [...this.#addins, addin];
    this.addinsChanged();
  }

  /** Remove an add-in by id. No-op if absent. Triggers {@link addinsChanged}. */
  removeAddin(id: string): void {
    if (!this.#addins.some((existing) => existing.id === id)) return;
    this.#addins = this.#addins.filter((existing) => existing.id !== id);
    this.addinsChanged();
  }

  // ── Contribution accessors ──────────────────────────────────────────────

  /** The merged ribbon schema — every addin ribbon contribution, in order. */
  mergedRibbonSchema(): RibbonTab[] {
    return mergeRibbonSchema(this.#addins);
  }

  /** The merged mini-toolbar buttons — every addin mini-toolbar contribution,
   *  in order. The host combines these with built-in defaults at editor boot
   *  (mirror of `ribbonTabs` + `mergedRibbonSchema` for the ribbon). */
  mergedMiniToolbar(): MiniToolbarButton[] {
    return mergeMiniToolbar(this.#addins);
  }

  /** Route `type` to the first registered addin that declares it. Returns
   *  whether a handler ran. */
  dispatchCommand(type: string, value?: string): boolean {
    for (const addin of this.#addins) {
      const handler = addin.commands?.[type];
      if (handler) {
        handler(this, value);
        return true;
      }
    }
    return false;
  }

  /** Subclass hook — re-render the ribbon/panes after the addin set changes.
   *  No-op by default. */
  protected addinsChanged(): void {
    // subclasses override
  }
}
