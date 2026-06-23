import { ribbonIcon } from "./icons";

/**
 * Shared DOM + icon plumbing for ribbon commands. Layout/icon wiring only —
 * every command affordance (color, hover, pressed, border, focus, keyboard,
 * a11y) comes from the wrapped Fluent elements.
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

/** A `<fluent-tooltip>` anchored to `#target` (the wrapped button). */
export const TOOLTIP_PART = `<fluent-tooltip anchor="target" positioning="top"><span class="rb-tip"></span></fluent-tooltip>`;

/**
 * Create a `.rb-icon` span. With a `slotName` it goes into Fluent's `start`
 * slot (icon + label); with an empty name it lands in the default slot (the
 * centered `content` part) so icon-only commands center their glyph.
 */
export function createIconSlot(slotName = "start"): HTMLSpanElement {
  const span = document.createElement("span");
  if (slotName) span.setAttribute("slot", slotName);
  span.className = "rb-icon";
  return span;
}

/** Inject the named Office icon svg into a slot (empty when unknown). */
export function renderIcon(slot: HTMLElement, name: string): void {
  slot.innerHTML = ribbonIcon(name) ?? "";
}

/** Set a slot's text (or clear it when empty). */
export function renderLabel(slot: HTMLElement, text: string): void {
  slot.replaceChildren(...(text ? [document.createTextNode(text)] : []));
}

/** Attributes a docen-* command forwards to its wrapped Fluent element. */
const FORWARDED_ATTRS = ["appearance", "size", "shape"] as const;

/**
 * Mirror `appearance`/`size`/`shape` from the host onto the wrapped Fluent
 * element, so callers drive the underlying button's look from the docen-*
 * attribute. `defaults` supplies the value used when the host omits an
 * attribute (a small ribbon button defaults to `subtle`; a split primary
 * defaults to nothing — Fluent's own look). Returns a disconnect function.
 */
export function forwardAttributes(
  host: HTMLElement,
  target: HTMLElement,
  defaults: Readonly<Record<string, string>> = {},
  exclude: readonly string[] = [],
): () => void {
  const attrs = FORWARDED_ATTRS.filter((name) => !exclude.includes(name));
  const sync = (name: string): void => {
    if (host.hasAttribute(name)) {
      target.setAttribute(name, host.getAttribute(name) ?? "");
    } else if (name in defaults) {
      target.setAttribute(name, defaults[name]);
    } else {
      target.removeAttribute(name);
    }
  };
  attrs.forEach(sync);
  const observer = new MutationObserver((records) => {
    for (const record of records) {
      if (record.attributeName) sync(record.attributeName);
    }
  });
  observer.observe(host, { attributes: true, attributeFilter: [...attrs] });
  return () => observer.disconnect();
}

/**
 * Stop a ribbon command host from stealing focus on mousedown. Clicking a
 * toolbar button would otherwise blur the contenteditable, dropping ProseMirror's
 * text selection before the command runs — `editor.chain().focus()` can restore
 * the caret but not a dragged selection range. preventDefault on mousedown
 * (capture phase, on the host) keeps the editor focused so the selection
 * survives; this mirrors Tiptap BubbleMenu's mousedown handler.
 *
 * Use only on hosts whose command is a pure click (button, toggle button). Do
 * NOT use on a typeable affordance (the font/size combobox `<input>` needs focus
 * to accept typed input). Returns a disconnect function.
 */
export function preventFocusLoss(host: HTMLElement): () => void {
  const onMousedown = (event: Event): void => event.preventDefault();
  host.addEventListener("mousedown", onMousedown, { capture: true });
  return () => host.removeEventListener("mousedown", onMousedown, { capture: true });
}
