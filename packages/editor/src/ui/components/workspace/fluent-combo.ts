/** Shared helpers for the dialogs' `<fluent-dropdown>` comboboxes — the
 *  programmatic-prefill semantics the native `<select>` gave for free. */

/** A `fluent-dropdown` combobox plus its picked value (null = none). */
export type FluentDropdown = HTMLElement & { value: string | null };

/** One `fluent-option` — the attr form mirrors the static template options
 *  (an absent value attr would fall back to the option's text). */
export function opt(text: string, value: string): HTMLElement {
  const el = document.createElement("fluent-option");
  el.textContent = text;
  el.setAttribute("value", value);
  return el;
}

/** The listbox a dropdown's options live in. */
export function listboxOf(sel: FluentDropdown | undefined | null): HTMLElement | null {
  return sel?.querySelector("fluent-listbox") ?? null;
}

/** Programmatically pick the option carrying `value`. The FAST `value`
 *  setter ignores options it hasn't indexed yet (freshly appended ones), so
 *  set `selected` on the option itself. The control input is only mirrored
 *  through selectOption() (the user-click path), which the listbox's
 *  idle-callback indexing races with — write the input here, exactly what a
 *  user click does. */
export function pick(sel: FluentDropdown | undefined | null, value: string): void {
  const listbox = listboxOf(sel);
  if (!listbox) return;
  const option = [
    ...listbox.querySelectorAll<HTMLElement & { selected?: boolean }>("fluent-option"),
  ].find((o) => o.getAttribute("value") === value);
  if (!option) return;
  option.selected = true;
  const input = sel?.querySelector('input[slot="control"]') as HTMLInputElement | null;
  if (input) input.value = option.textContent ?? "";
}

/** The dropdown's picked value. A user pick syncs the FAST `value`
 *  property; a programmatic prefill may leave it "" (the control input
 *  still shows the text) — fall back to the input so prefill-then-OK
 *  round-trips. */
export function pickedValue(sel: FluentDropdown | undefined | null): string | null {
  if (!sel) return null;
  const input = sel.querySelector('input[slot="control"]') as HTMLInputElement | null;
  return sel.value || input?.value.trim() || null;
}

/** A standalone `fluent-listbox` (the dialogs' multi-row `.list` pickers) —
 *  single mode unless `multiple` is set; reads selection off `selectedOptions`. */
export type FluentListbox = HTMLElement & { selectedOptions?: Element[] };

/** The listbox's picked value (null = none). A standalone list has no control
 *  input — the selected options are the only readback. */
export function listValue(listbox: FluentListbox | undefined | null): string | null {
  return listbox?.selectedOptions?.[0]?.getAttribute("value") ?? null;
}
