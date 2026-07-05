import { ribbonIcon } from "./icons";

/**
 * Shared host style + icon injection for ribbon commands. Layout/icon wiring
 * only — every command affordance (color, hover, pressed, border, focus,
 * keyboard, a11y) comes from the wrapped Fluent elements.
 */

/** Host flex + icon sizing (currentColor so the Fluent theme paints the svg). */
export const COMMAND_HOST_STYLE = `
  :host { display: inline-flex; }
  .rb-icon { display: contents; }
  .rb-icon svg { display: block; fill: currentColor; width: 16px; height: 16px; }
  :host([icon-only]) .rb-label { display: none; }
  /* Scope to fluent-button only. ::part(content) is a wildcard and would
     also hit fluent-menu-item / fluent-option (both expose part="content"),
     forcing their text into flex-center and overflowing it sideways out of
     the item — the dropdown options of an icon-only split looked off while a
     large split's (no [icon-only]) stayed correct, exactly because of this. */
  :host([icon-only]) fluent-button::part(content) {
    display: flex; justify-content: center; align-items: center;
  }
`;

// Cache a parsed <template> per icon string so repeated renders clone instead
// of re-running the HTML parser on the same static SVG markup (a ribbon mount
// parses ~100 icons; cloning a cached template skips the parser entirely).
const iconTemplates = new Map<string, HTMLTemplateElement>();

/** Inject the named Office icon svg into a slot (empty when unknown). */
export function renderIcon(slot: HTMLElement, name: string): void {
  const svg = ribbonIcon(name);
  if (!svg) {
    slot.replaceChildren();
    return;
  }
  let template = iconTemplates.get(svg);
  if (!template) {
    template = document.createElement("template");
    template.innerHTML = svg;
    iconTemplates.set(svg, template);
  }
  slot.replaceChildren(template.content.cloneNode(true));
}

/** A popover-capable element (fluent-tooltip) — showPopover/hidePopover are the
 *  native Popover API on HTMLElement; typed optional so this file compiles even
 *  where the TS DOM lib hasn't declared them yet. */
type PopoverElement = HTMLElement & { showPopover?: () => void; hidePopover?: () => void };

/**
 * Keep a ribbon command's tooltip from dismissing its own menu.
 *
 * fluent-tooltip's showTooltip schedules showPopover() on a ~250ms timer; if the
 * click lands inside that window the menu opens first and the tooltip's delayed
 * showPopover fires after it. The tooltip and the menu-list are both auto
 * popovers, so the re-shown tooltip light-dismisses the menu via auto-popover
 * mutual exclusion — the "menu appears then instantly vanishes" flicker
 * (intermittent: only when hover precedes the click by <250ms).
 *
 * While the menu is open, no-op the tooltip's showPopover (blocks the pending
 * delayed show) and hidePopover it on open (clears one already shown, bypassing
 * fluent-tooltip's :hover guard, which otherwise keeps it up). Returns a
 * disposer to call in disconnectedCallback.
 */
export function suppressTooltipWhileMenuOpen(
  tooltip: PopoverElement | undefined,
  menuList: HTMLElement | undefined,
): () => void {
  if (!tooltip || !menuList) return () => {};
  let menuOpen = false;
  const origShow = tooltip.showPopover?.bind(tooltip);
  if (origShow) {
    tooltip.showPopover = () => {
      if (menuOpen) return;
      origShow();
    };
  }
  const onToggle = (event: Event): void => {
    menuOpen = (event as ToggleEvent).newState === "open";
    if (menuOpen) tooltip.hidePopover?.();
  };
  menuList.addEventListener("toggle", onToggle as EventListener);
  return () => {
    menuList.removeEventListener("toggle", onToggle as EventListener);
    // Drop the instance override so tooltip.showPopover resolves back to the
    // native HTMLElement prototype method.
    delete (tooltip as { showPopover?: () => void }).showPopover;
  };
}

/** A menu item's renderable shape (the shared subset of every ribbon/context
 *  menu item type — command routing fields stay on the caller's own type). */
interface MenuItemLike {
  text: string;
  checked?: boolean;
  disabled?: boolean;
}

/** Append `items` as `<fluent-menu-item>`s into `list`, replacing its children.
 *  A `change` on any item routes to `onSelect(item)`. `checked` items render as
 *  `role="menuitemradio"` with Fluent's own checkmark; plain items stay
 *  `menuitem`. Every item gets `data-indent="0"` so a plain-text label spans the
 *  full row — without it Fluent pins the content to the fixed-width indicator
 *  track and clips long labels (see the registry's fluent-menu-item override). */
export function appendMenuItems<T extends MenuItemLike>(
  list: HTMLElement,
  items: readonly T[],
  onSelect: (item: T) => void,
): void {
  list.replaceChildren();
  for (const item of items) {
    const menuItem = document.createElement("fluent-menu-item");
    if (item.checked) {
      menuItem.setAttribute("role", "menuitemradio");
      menuItem.setAttribute("checked", "");
    } else {
      menuItem.setAttribute("role", "menuitem");
    }
    menuItem.setAttribute("data-indent", "0");
    menuItem.textContent = item.text;
    if (item.disabled) menuItem.setAttribute("disabled", "");
    menuItem.addEventListener("change", () => onSelect(item));
    list.append(menuItem);
  }
}
