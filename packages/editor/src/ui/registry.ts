import {
  Accordion,
  AccordionDefinition,
  AccordionItem,
  AccordionItemDefinition,
  Button,
  ButtonStyles,
  ButtonTemplate,
  Checkbox,
  CheckboxDefinition,
  CompoundButton,
  CompoundButtonDefinition,
  Dialog,
  DialogBody,
  DialogBodyDefinition,
  DialogDefinition,
  Divider,
  DividerDefinition,
  Dropdown,
  DropdownStyles,
  DropdownTemplate,
  DropdownOption,
  DropdownOptionStyles,
  DropdownOptionTemplate,
  Drawer,
  DrawerBody,
  DrawerBodyDefinition,
  DrawerStyles,
  DrawerTemplate,
  Field,
  FieldDefinition,
  FluentDesignSystem,
  Menu,
  MenuButton,
  MenuButtonStyles,
  MenuButtonTemplate,
  MenuItem,
  MenuItemStyles,
  MenuItemTemplate,
  MenuList,
  MenuListDefinition,
  MenuStyles,
  MenuTemplate,
  ProgressBar,
  ProgressBarStyles,
  ProgressBarTemplate,
  Listbox,
  ListboxDefinition,
  Switch,
  SwitchDefinition,
  Tab,
  TabStyles,
  TabTemplate,
  Tablist,
  TablistDefinition,
  Radio,
  RadioDefinition,
  RadioGroup,
  RadioGroupDefinition,
  TextInput,
  TextInputDefinition,
  TextArea,
  TextAreaDefinition,
  Tree,
  TreeDefinition,
  TreeItem,
  TreeItemDefinition,
  ToggleButton,
  ToggleButtonStyles,
  ToggleButtonTemplate,
  Tooltip,
  TooltipDefinition,
} from "@fluentui/web-components";
import {
  FASTElementDefinition,
  css,
  type Constructable,
  type PartialFASTElementDefinition,
} from "@microsoft/fast-element";

import "./components"; // side-effect: registers docen-* custom elements
import { injectOfficeTokens } from "./tokens";

let registered = false;

/**
 * fluent-tab pins its `--textContent` placeholder — a hidden bold clone of the
 * tab's text that holds the tab's width so bolding the active tab doesn't
 * reflow its neighbors — once, in `connectedCallback`. Rewriting the tab's
 * light-DOM text later (e.g. switching locale) leaves the placeholder on the
 * old text, so a tab whose text shrank (English → Chinese) keeps the old,
 * wider width and trails a gap. Re-sync `--textContent` whenever the text
 * changes; set it inline so it wins over the style Fluent injects on connect.
 */
class TabWithSyncedText extends Tab {
  #observer?: MutationObserver;

  override connectedCallback(): void {
    super.connectedCallback();
    this.#syncPlaceholder();
    this.#observer = new MutationObserver(() => this.#syncPlaceholder());
    this.#observer.observe(this, { childList: true, characterData: true, subtree: true });
  }

  override disconnectedCallback(): void {
    this.#observer?.disconnect();
    this.#observer = undefined;
    super.disconnectedCallback();
  }

  #syncPlaceholder(): void {
    // `content: var(--textContent)` needs a CSS <string> (quoted); escape any
    // backslash/quote in the text so the value stays a single quoted string.
    const text = (this.textContent ?? "").replace(/\\/g, "\\\\").replace(/'/g, "\\'");
    this.style.setProperty("--textContent", `'${text}'`);
  }
}

/**
 * Composes a Fluent component (or a docen override) and defines it in the
 * Fluent registry. fast-element 3.0 made `FASTElementDefinition.compose()`
 * async (it returns a `Promise`), so each registration is awaited. Fluent
 * exports its `*Definition` objects as untyped `PartialFASTElementDefinition`
 * (no type parameter), and the template generic is invariant, so the partial
 * is cast to the type's parameter before composing.
 */
async function defineElement<T extends Constructable<HTMLElement>>(
  type: T,
  def: PartialFASTElementDefinition,
): Promise<void> {
  (await FASTElementDefinition.compose(type, def as PartialFASTElementDefinition<T>)).define(
    FluentDesignSystem.registry,
  );
}

/**
 * Register the Fluent UI components docen wraps, plus a docen-named alias of
 * `fluent-tab` (`docen-ribbon-tab`). Call once at app bootstrap.
 *
 * `docen-ribbon-tab` is the fluent `Tab` implementation (same template/styles,
 * native indicator) registered under a docen name — not a hand-rolled tab.
 */
export async function registerComponents(): Promise<void> {
  if (registered) return;
  registered = true;
  // Fluent UI's Tablist calls `rootNode.getElementById(...)` in setTabs/
  // changeTab, where `rootNode = this.getRootNode()`. While the tablist is still
  // connecting, getRootNode() resolves to the tablist element ITSELF (an
  // Element), and Element has no getElementById → "rootNode.getElementById is
  // not a function". Patch Element.prototype with a querySelector-based lookup
  // (idempotent; Document/ShadowRoot/DocumentFragment already have their own).
  const elProto = Element.prototype as unknown as {
    getElementById?: (id: string) => Element | null;
  };
  if (typeof elProto.getElementById !== "function") {
    elProto.getElementById = function (this: Element, id: string): Element | null {
      return this.querySelector(`#${CSS.escape(id)}`);
    };
  }
  injectOfficeTokens();
  // fluent-button: trim the default 12px horizontal padding + 96px min-width
  // and add a column-gap so an icon (start slot) and its label have breathing
  // room. Pinned to subtle/transparent (the ribbon appearances); `.control`-like
  // internals expose no `part`, so the override is merged at registration.
  await defineElement(Button, {
    name: "fluent-button",
    template: ButtonTemplate,
    styles: css`
      ${ButtonStyles}
      ${css`
        :host {
          column-gap: 6px;
        }
      `}
      :host([appearance="subtle"]),
      :host([appearance="transparent"]) {
        padding-inline: 6px;
        min-width: 0;
      }
    `,
  });
  await defineElement(Checkbox, CheckboxDefinition);
  await defineElement(CompoundButton, CompoundButtonDefinition);
  await defineElement(Divider, DividerDefinition);
  // fluent-dropdown: the built-in `.control { min-width: 160px }` forces every
  // combobox at least 160px wide. The control exposes no `part`, so the override
  // is merged into the composed styles at registration (css`` can interpolate
  // the original ElementStyles) — lets the combobox shrink to a font-size width.
  await defineElement(Dropdown, {
    name: "fluent-dropdown",
    template: DropdownTemplate,
    styles: css`
      ${DropdownStyles} .control {
        min-width: 0;
      }
    `,
  });
  // fluent-option: the short (font-size) combobox centers its options; every
  // combobox hides the selected checkmark. The checkmark lives in the option's
  // own shadow (no part), so the overrides are composed at registration and
  // pinned to data-center (center + hide) / data-no-checkmark (hide only).
  await defineElement(DropdownOption, {
    name: "fluent-option",
    template: DropdownOptionTemplate,
    styles: css`
      ${DropdownOptionStyles}
      ${css`
        :host([data-center]) {
          display: flex;
          justify-content: center;
          align-items: center;
          padding-inline: 0;
          width: 100%;
        }
      `}
      :host([data-center]) slot[name="checked-indicator"],
      :host([data-center]) slot[name="start"],
      :host([data-center]) .description {
        display: none;
      }
      :host([data-center]) .content {
        text-align: center;
      }
      :host([data-no-checkmark]) slot[name="checked-indicator"] {
        display: none;
      }
    `,
  });
  // fluent-menu-button: the end-slot caret keeps a --icon-spacing inline-start
  // margin meant to separate it from a label. icon-only (the split caret) has
  // no label, so under justify-content:center that lone margin pushes the
  // glyph off-center — drop it for icon-only so the caret sits centered.
  await defineElement(MenuButton, {
    name: "fluent-menu-button",
    template: MenuButtonTemplate,
    styles: css`
      ${MenuButtonStyles}
      :host([icon-only]) ::slotted([slot="end"]),
      :host([icon-only]) [slot="end"] {
        margin-inline-start: 0;
      }
    `,
  });
  // fluent-menu: a large split stacks its primary over a caret bar
  // (data-vertical). The built-in split is horizontal; the column override is
  // pinned to data-vertical so ordinary menus/splits are untouched. Fluent also
  // paints a split divider on the primary's inline-end edge regardless of
  // appearance — clear it so a subtle split is fully flat until hover.
  await defineElement(Menu, {
    name: "fluent-menu",
    template: MenuTemplate,
    styles: css`
      ${MenuStyles}
      ${css`
        :host([data-vertical][split]) {
          flex-direction: column;
        }
      `}
      /* Fluent paints a split divider on the primary's inline-end (and clears
         the trigger's inline-start) regardless of appearance; clear both so a
         subtle split is fully flat by default. */
      :host([split]) ::slotted([slot="primary-action"]) {
        border-inline-end-color: transparent;
      }
      :host([split]) ::slotted([slot="trigger"]) {
        border-inline-start-color: transparent;
      }
      /* Hovering the split (primary or caret) lights both up together — the
         :hover is on the menu, so the pointer over either child counts. */
      :host([split]:hover) ::slotted([slot="primary-action"]),
      :host([split]:hover) ::slotted([slot="trigger"]) {
        border-color: var(--colorNeutralStroke1);
      }
      /* Vertical split: the caret is below — restore its inline-start (left)
         (Fluent's split clears it for the horizontal layout) and drop the
         primary's block-end so the caret's block-start is the single divider. */
      :host([data-vertical][split]:hover) ::slotted([slot="trigger"]) {
        border-inline-start: var(--strokeWidthThin) solid var(--colorNeutralStroke1);
      }
      :host([data-vertical][split]:hover) ::slotted([slot="primary-action"]) {
        border-block-end-color: transparent;
      }
    `,
  });
  // fluent-menu-item overrides on top of Fluent's 7-track subgrid (10px rim /
  // indicator / start / content / end / submenu-glyph / 10px rim — the checkmark
  // paints in the indicator track, col 2). A plain-text item (no start icon ->
  // data-indent="0") pins its content to the 1fr content track (col 4), whose
  // width the popover's max-inline-size clamps — a long label (e.g. "Add Space
  // Before Paragraph") overflows and is clipped. Span the full row so the label
  // stretches the popover to fit. `data-indent="1"` marks a plain member of a
  // pick list (Word's Editing/Viewing drop-down): it too gives up the checkmark
  // track so both labels share one left edge. A radio item's label starts after
  // the checkmark (col 3) — col 2 would collide with the checkmark there and
  // grid auto-placement would push the label to a SECOND row (the checked item's
  // icon and text stacked). Items with a start icon keep Fluent's default
  // (icon-then-text columns).
  await defineElement(MenuItem, {
    name: "fluent-menu-item",
    template: MenuItemTemplate,
    styles: css`
      ${MenuItemStyles}
      :host([data-indent="0"]) .content {
        grid-column: 1 / -1;
      }
      :host([data-indent="1"]) .content,
      :host([role="menuitemradio"]) .content,
      :host([role="menuitemcheckbox"]) .content {
        grid-column: 3 / -1;
      }
    `,
  });
  await defineElement(MenuList, MenuListDefinition);
  // fluent-drawer (type=inline = side-by-side, pushes content — Office sidebar
  // style) + fluent-tree/fluent-tree-item for the outline panel.
  // fluent-drawer: inline drawers are content-pushing sidebars. The built-in
  // exit animates the <dialog> with translateX, but the host collapses at once
  // (width:fit-content follows display), so the dialog slides out floating
  // past the already-collapsed host — reads as the pane swelling then
  // vanishing. Inline panes should open/close instantly, so drop the
  // transition for inline.
  await defineElement(Drawer, {
    name: "fluent-drawer",
    template: DrawerTemplate,
    styles: css`
      ${DrawerStyles}
      ${css`
        :host([type="inline"]) dialog {
          transition: none;
        }
      `}
    `,
  });
  await defineElement(DrawerBody, DrawerBodyDefinition);
  // fluent-dialog (type=modal = native <dialog> showModal → backdrop + ESC) +
  // fluent-dialog-body (title / content / action regions) for <docen-dialog>.
  await defineElement(Dialog, DialogDefinition);
  await defineElement(DialogBody, DialogBodyDefinition);
  await defineElement(Tree, TreeDefinition);
  await defineElement(TreeItem, TreeItemDefinition);
  // fluent-accordion / fluent-accordion-item: collapsible sections for the
  // properties panel (radio groups, label+input rows).
  await defineElement(Accordion, AccordionDefinition);
  await defineElement(AccordionItem, AccordionItemDefinition);
  // fluent-radio / fluent-radio-group: single-choice rows in the properties panel.
  await defineElement(Radio, RadioDefinition);
  await defineElement(RadioGroup, RadioGroupDefinition);
  // fluent-field: wraps each radio with its label (label-position="after" =
  // radio on the start, text on the end) — the radio itself renders no text.
  await defineElement(Field, FieldDefinition);
  await defineElement(Switch, SwitchDefinition);
  // fluent-tab: use TabWithSyncedText so the width placeholder re-syncs when
  // the tab's text changes (see TabWithSyncedText) — tabs shrink to their
  // text, including after a locale switch Fluent's one-shot set misses.
  await defineElement(TabWithSyncedText, {
    name: "fluent-tab",
    template: TabTemplate,
    styles: TabStyles,
  });
  await defineElement(Listbox, ListboxDefinition);
  // fluent-progress-bar: the document area's load veil (Office shows
  // "Opening <file>" over the page area while a large document opens).
  // Fluent's stock indeterminate loop animates inset-inline-start — a layout
  // property the main thread owns, so the bar freezes solid while a
  // synchronous parse blocks the thread. Restate the same sweep as a
  // transform animation (compositor-driven, keeps moving through the
  // freeze): the 33%-wide indicator travels -33% → 100% of the track =
  // -100% → ~303% of its own width.
  await defineElement(ProgressBar, {
    name: "fluent-progress-bar",
    template: ProgressBarTemplate,
    styles: css`
      ${ProgressBarStyles}
      :host(:not([value])) .indicator {
        animation-name: docen-indeterminate;
        will-change: transform;
      }
      @keyframes docen-indeterminate {
        from {
          transform: translateX(-100%);
        }
        to {
          transform: translateX(303%);
        }
      }
    `,
  });
  await defineElement(Tablist, TablistDefinition);
  await defineElement(TextInput, TextInputDefinition);
  // fluent-textarea (no text-area hyphenation — Fluent v3's registered tag):
  // the comments pane's compose box (Word's comment sidebar uses a multiline
  // input with Ctrl+Enter to post).
  await defineElement(TextArea, TextAreaDefinition);
  // fluent-toggle-button: same padding/min-width/gap tightening as the button.
  await defineElement(ToggleButton, {
    name: "fluent-toggle-button",
    template: ToggleButtonTemplate,
    styles: css`
      ${ToggleButtonStyles}: host {
        column-gap: 6px;
      }
      :host([appearance="subtle"]),
      :host([appearance="transparent"]) {
        padding-inline: 6px;
        min-width: 0;
      }
    `,
  });
  await defineElement(Tooltip, TooltipDefinition);
  // docen alias of fluent Tab — fluent's Tab implementation (same template/
  // styles, native indicator) under a docen name. Mirrors how fluent builds
  // TabDefinition.
  await defineElement(TabWithSyncedText, {
    name: "docen-ribbon-tab",
    template: TabTemplate,
    styles: TabStyles,
  });
}
