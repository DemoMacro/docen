import type { DocenTranslation } from "../i18n";

/**
 * Office.js-style add-in model for docen editor hosts.
 *
 * A {@link DocenHost} (e.g. `<docen-document>`) owns the editing surface; every
 * surrounding UI surface — ribbon tabs, task panes, commands — is *contributed*
 * by a {@link DocenAddin}. The host merges contributions and routes commands,
 * so a host ships with a default addin (editor essentials) and users plug in
 * more (citations, mail-merge, …) without touching host internals.
 *
 * Names mirror the Office.js manifest: ribbon `Tab > Group > Control > Action`,
 * `ExtensionPoint`-style contributions, `<Host>` = the host application.
 * (learn.microsoft.com — Office Add-ins manifest overview)
 */

// ── Ribbon schema (Tab > Group > Control) ────────────────────────────────────

/** Ribbon control size — `large` is a tall labelled button (Word's large
 *  control); `small` is an icon-only button stacked in a group row/column. */
export type RibbonControlSize = "small" | "large";

/** A single menu/combobox option. `value` is the data carried by the command;
 *  `event` lets an option dispatch a different command than its control's. */
export interface RibbonMenuItem {
  /** Display text — an i18n key resolved via `t()` at render time (a plain
   *  string also works; `t()` returns it unchanged if no key matches). */
  text: string;
  value?: string;
  event?: string;
  checked?: boolean;
  disabled?: boolean;
}

/** Fields shared by every ribbon control. `event` is the kebab-case command
 *  name the control dispatches on activation. */
export interface RibbonControlBase {
  id?: string;
  /** Docen icon key (see the RIBBON_ICONS map). */
  icon?: string;
  /** Label — an i18n key resolved via `t()` at render time (a plain string
   *  also works). */
  label?: string;
  event?: string;
  disabled?: boolean;
  size?: RibbonControlSize;
  /** Render the icon only (label surfaces in a tooltip). */
  iconOnly?: boolean;
}

export interface RibbonButton extends RibbonControlBase {
  type: "button";
}

export interface RibbonMenu extends RibbonControlBase {
  type: "menu";
  items: RibbonMenuItem[];
}

/** A button with a default action plus a dropdown of alternatives — Word's
 *  split button (e.g. Paste / Paste Special). */
export interface RibbonSplit extends RibbonControlBase {
  type: "split";
  items: RibbonMenuItem[];
}

export interface RibbonCombobox extends RibbonControlBase {
  type: "combobox";
  value?: string;
  items?: RibbonMenuItem[];
  /** Special source: backfill options from the Local Font Access API. */
  source?: "local-fonts";
  /** `short` = the narrow font-size combobox; otherwise full width. */
  comboboxSize?: "short" | "normal";
}

/** A swatch popover + "More Colors" picker (font color / paragraph shading). */
export interface RibbonColorPicker extends RibbonControlBase {
  type: "color-picker";
  defaultColor?: string;
}

/** A vertical separator between controls in a row. */
export interface RibbonSeparator {
  type: "separator";
}

export type RibbonControl =
  | RibbonButton
  | RibbonMenu
  | RibbonSplit
  | RibbonCombobox
  | RibbonColorPicker
  | RibbonSeparator;

/** A layout wrapper. Office stacks a large button beside rows/columns of small
 *  buttons; docen expresses that with explicit column/row/grid groups (the
 *  rb-col / rb-row / rb-grid classes the ribbon container renders). */
export interface RibbonLayout {
  type: "layout";
  layout: "column" | "row" | "grid";
  controls: readonly RibbonControlOrLayout[];
}

/** A control or a layout wrapper — a group's `controls` is a tree of these. */
export type RibbonControlOrLayout = RibbonControl | RibbonLayout;

export interface RibbonGroup {
  id: string;
  /** Group heading — an i18n key resolved via `t()` at render time. */
  label: string;
  /** Optional launcher id (opens a dialog or pane). */
  launcher?: string;
  controls: readonly RibbonControlOrLayout[];
}

export interface RibbonTab {
  id: string;
  /** Tab heading — an i18n key resolved via `t()` at render time. */
  label: string;
  groups: RibbonGroup[];
}

// ── Contributions (what an addin gives the host) ─────────────────────────────

/** Ribbon contribution — Office.js `ExtensionPoint > CustomTab|OfficeTab`.
 *  Target an existing tab id (home/insert/…) to append groups, or a fresh id to
 *  create a new tab. */
export interface RibbonContribution {
  /** Target tab id. Existing id → append groups; new id → create a tab. */
  tab: string;
  /** Tab heading (i18n key) for a newly created tab — ignored when targeting
   *  an existing tab. */
  label?: string;
  groups: RibbonGroup[];
}

/** Mini-toolbar button — a flat action on the selection floating toolbar (the
 *  Word "mini toolbar"). `event` is the command name the host routes to
 *  `editor.commands.<event>`; `label` is an i18n key resolved at render time
 *  (e.g. "ribbon.cmd.bold"); `activeMark` drives aria-pressed (omit for
 *  non-toggle actions like clear-format). */
export interface MiniToolbarButton {
  id?: string;
  /** Docen icon key (see the RIBBON_ICONS map). */
  icon: string;
  /** Command name (editor.commands.<event>). */
  event: string;
  /** Tooltip i18n key (resolved at render time), e.g. "ribbon.cmd.bold". */
  label?: string;
  /** Mark name for the pressed state; omit for non-toggle actions. */
  activeMark?: string;
}

/** Task pane contribution — Office.js `ShowTaskpane` action. `start` is the
 *  navigation side (left in LTR); `end` is the format/properties side (right). */
export interface TaskPaneContribution<THost extends DocenHost = DocenHost> {
  id: string;
  title: string;
  position: "start" | "end";
  icon?: string;
  /** Build the pane's content element. Omitted = the host fills the pane itself
   *  (e.g. the built-in outline/search navigation). */
  render?: (host: THost) => HTMLElement | null;
  defaultOpen?: boolean;
}

// ── Host + Addin contracts ───────────────────────────────────────────────────

/** A docen editor host — the `<Host>` in Office.js terms (Word/Excel/PowerPoint
 *  → docen Document/Presentation/Workbook). Owns the editing surface; addins
 *  contribute the surrounding UI. */
export interface DocenHost<TEditor = unknown> {
  readonly element: HTMLElement;
  readonly editor: TEditor | undefined;
  getContent(): unknown;
  setContent(content: unknown): void;
  /** Route `type` to the first addin that declares it. Returns whether handled. */
  dispatchCommand(type: string, value?: string): boolean;
}

/** An Office.js-style addin. Every field is optional except `id` — an addin may
 *  contribute only ribbon tabs, only a task pane, or only commands. */
export interface DocenAddin<THost extends DocenHost = DocenHost> {
  readonly id: string;
  readonly name?: string;
  /** Per-locale translation tables this addin contributes. The host registers
   *  them on `addAddin` (merged into the global table, so built-in keys and
   *  other addins' keys coexist). Pure data, so the `addins` JSON attribute
   *  carries them too. Office.js parallel: manifest `localizationInfo` +
   *  per-locale JSON. */
  readonly translations?: readonly DocenTranslation[];
  readonly ribbon?: readonly RibbonContribution[];
  /** Mini-toolbar buttons (the selection floating toolbar — the Word "mini
   *  toolbar"). Office.js keeps it internal; docen opens it as an addin
   *  surface, symmetric to `ribbon`. Merged with built-in defaults
   *  (`defaultMiniToolbarButtons()` + `mergeMiniToolbar`); runtime `addAddin`
   *  re-merges immediately (the bar's buttons are `@observable`), so
   *  contributions appear without re-mounting. */
  readonly miniToolbar?: readonly MiniToolbarButton[];
  readonly taskPanes?: readonly TaskPaneContribution<THost>[];
  readonly commands?: Readonly<Record<string, (host: THost, value?: string) => void>>;
}
