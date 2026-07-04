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
