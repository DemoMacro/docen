import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import {
  CROSS_REFERENCE_CONTENTS,
  crossReferenceContentAvailable,
  crossReferenceContentKey,
  crossReferenceTypeKey,
  type CrossRefContent,
  type CrossRefKind,
  type CrossReferenceTarget,
} from "../../../document/cross-reference";
import { observeLang, t } from "../../i18n/localize";
import {
  listboxOf,
  opt,
  pick,
  pickedValue,
  type FluentDropdown,
  type FluentListbox,
} from "./fluent-combo";

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(460px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .row {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .row > label {
    min-width: 92px;
  }
  .row fluent-dropdown {
    flex: 1;
    min-width: 0;
  }
  .row fluent-dropdown input {
    width: 100%;
    box-sizing: border-box;
  }
  .preview {
    border: 1px solid var(--colorNeutralStroke2, #e0e0e0);
    border-radius: 4px;
    background: var(--colorNeutralBackground1, #fff);
    padding: 6px 10px;
    min-height: 20px;
  }
  .which {
    opacity: 0.75;
  }
  fluent-listbox.list {
    width: 100%;
    min-height: 168px;
    border: 1px solid var(--colorNeutralStroke1, #d1d1d1);
    border-radius: 4px;
    padding: 2px;
    font-size: 13px;
    box-shadow: none;
  }
`;

/** One Reference type list entry: a kind, or one caption label's bucket. */
interface TypeEntry {
  id: string;
  kind: CrossRefKind;
  label?: string;
}

const template = html<DocenCrossReferenceDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="row">
        <label ${ref("typeLabel")}></label>
        <fluent-dropdown
          type="combobox"
          appearance="outline"
          ${ref("typeSel")}
          @change="${(x) => x.syncType()}"
        >
          <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
          <input
            slot="control"
            role="combobox"
            aria-haspopup="listbox"
            type="combobox"
            size="1"
            style="width:100%;box-sizing:border-box"
          />
        </fluent-dropdown>
      </div>
      <div class="row">
        <label ${ref("contentLabel")}></label>
        <fluent-dropdown
          type="combobox"
          appearance="outline"
          ${ref("contentSel")}
          @change="${(x) => x.syncPreviewSoon()}"
        >
          <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
          <input
            slot="control"
            role="combobox"
            aria-haspopup="listbox"
            type="combobox"
            size="1"
            style="width:100%;box-sizing:border-box"
          />
        </fluent-dropdown>
      </div>
      <div class="preview" ${ref("previewEl")}></div>
      <span class="which" ${ref("whichLabel")}></span>
      <fluent-listbox
        class="list"
        ${ref("listSel")}
        @click="${(x) => x.syncPreviewSoon()}"
        @change="${(x) => x.syncPreviewSoon()}"
      ></fluent-listbox>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyCrossReference()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-cross-reference-dialog>` — Word's Cross-reference dialog (交叉引用)
 * with its full type × content matrix: numbered items, headings, bookmarks,
 * footnotes, endnotes, and caption labels (Figure/Table/…); the equation type
 * defers to the equation workstream and shows disabled. The "Insert reference
 * to" list follows the picked type (the word follows too: "Heading text" vs
 * "Bookmark text"), the preview box shows what the reference will read, and
 * options with nothing to point at stay disabled (a caption without text, an
 * unnumbered heading). Opened with `show(targets, caretPos)`; commits via
 * `cross-ref:ok` `{ key, content }` or cancels.
 */
@customElement({ name: "docen-cross-reference-dialog", template, styles })
class DocenCrossReferenceDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable typeLabel?: HTMLElement;
  @observable typeSel?: FluentDropdown;
  @observable contentLabel?: HTMLElement;
  @observable contentSel?: FluentDropdown;
  @observable previewEl?: HTMLElement;
  @observable whichLabel?: HTMLElement;
  @observable listSel?: FluentListbox;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  /** The candidates the host passed to show() — filtered by the type. */
  #targets: CrossReferenceTarget[] = [];
  /** The caret the dialog opened at (the above/below preview's field side). */
  #caretPos?: number;
  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#unobserveLang = observeLang(() => this.#rebuild());
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  show(targets: CrossReferenceTarget[], caretPos?: number): void {
    this.#targets = targets;
    this.#caretPos = caretPos;
    this.#buildTypes();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible type change handler (FAST bindings can't reach
   *  `#`-private methods). */
  #syncType(): void {
    const type = this.#selectedType();
    this.#renderContents(type);
    this.#renderList(type);
    this.#applyLabels(type);
    this.syncPreview();
  }

  syncType(): void {
    this.#syncType();
  }

  /** The picked content / target changed — refresh the disabled options and
   *  the live preview. */
  syncPreview(): void {
    const type = this.#selectedType();
    const target = this.#selectedTarget();
    this.#syncContentDisabled(type, target);
    if (!this.previewEl) return;
    const content = this.#picked(this.contentSel) as CrossRefContent | null;
    this.previewEl.textContent =
      target && content && !this.#contentDisabled(target, content)
        ? this.#previewOf(target, content)
        : "";
  }

  /** The list's click selection lands after its own click handler runs — the
   *  template's preview refresh waits a microtask so it reads the new pick. */
  syncPreviewSoon(): void {
    queueMicrotask(() => this.syncPreview());
  }

  applyCrossReference(): void {
    const type = this.#selectedType();
    const target = this.#selectedTarget();
    const content = this.#picked(this.contentSel) as CrossRefContent | null;
    if (!type || !target || !content) return;
    if (!CROSS_REFERENCE_CONTENTS[type.kind].includes(content)) return;
    if (this.#contentDisabled(target, content)) return;
    this.$emit("cross-ref:ok", { key: target.key, content });
    this.hide();
  }

  /** Rebuild the type options from the candidates — Word's fixed order with
   *  the disabled equation entry, then one bucket per caption label — and
   *  pick the first usable type. */
  #buildTypes(): void {
    const entries = this.#typeEntries();
    const listbox = listboxOf(this.typeSel);
    listbox?.replaceChildren(
      ...entries.map((entry) => {
        const option = opt(this.#typeLabel(entry), entry.id);
        // Equation defers to the equation workstream (G4): shown, disabled.
        if (entry.kind === "equation") option.setAttribute("disabled", "");
        return option;
      }),
    );
    const first = this.#options(this.typeSel).find((option) => !option.hasAttribute("disabled"));
    if (first) this.#pickOption(this.typeSel, first.getAttribute("value") ?? "");
    this.#syncType();
  }

  /** Language change — rebuild the labels keeping the picks when possible. */
  #rebuild(): void {
    const type = this.#picked(this.typeSel);
    const target = this.#picked(this.listSel);
    const content = this.#picked(this.contentSel);
    this.#buildTypes();
    if (
      type &&
      this.#options(this.typeSel).some((option) => option.getAttribute("value") === type)
    ) {
      this.#pickOption(this.typeSel, type);
      this.#syncType();
    }
    if (
      target &&
      this.#options(this.listSel).some((option) => option.getAttribute("value") === target)
    )
      this.#pickOption(this.listSel, target);
    if (
      content &&
      this.#options(this.contentSel).some(
        (option) => option.getAttribute("value") === content && !option.hasAttribute("disabled"),
      )
    )
      this.#pickOption(this.contentSel, content);
    this.syncPreview();
  }

  #typeEntries(): TypeEntry[] {
    const entries: TypeEntry[] = [];
    for (const kind of [
      "numbered",
      "heading",
      "bookmark",
      "footnote",
      "endnote",
    ] as CrossRefKind[]) {
      if (this.#targets.some((target) => target.kind === kind)) entries.push({ id: kind, kind });
    }
    entries.push({ id: "equation", kind: "equation" });
    const labels: string[] = [];
    for (const target of this.#targets) {
      if (target.kind !== "caption") continue;
      const label = target.label ?? "";
      if (!labels.includes(label)) labels.push(label);
    }
    for (const label of labels) entries.push({ id: `caption:${label}`, kind: "caption", label });
    return entries;
  }

  #typeLabel(entry: TypeEntry): string {
    const key = crossReferenceTypeKey(entry.kind, entry.label);
    return key ? t(key, this) : (entry.label ?? "");
  }

  #selectedType(): TypeEntry | null {
    const entries = this.#typeEntries();
    const picked = this.#picked(this.typeSel);
    return entries.find((entry) => entry.id === picked) ?? null;
  }

  #selectedTarget(): CrossReferenceTarget | null {
    const key = this.#picked(this.listSel);
    return this.#targets.find((target) => target.key === key) ?? null;
  }

  #renderContents(type: TypeEntry | null): void {
    const listbox = listboxOf(this.contentSel);
    if (!listbox) return;
    const contents = type ? CROSS_REFERENCE_CONTENTS[type.kind] : [];
    listbox.replaceChildren(
      ...contents.map((content) =>
        opt(t(crossReferenceContentKey(type!.kind, content), this), content),
      ),
    );
  }

  #renderList(type: TypeEntry | null): void {
    if (!this.listSel) return;
    const targets = type
      ? this.#targets.filter((target) =>
          type.kind === "caption"
            ? target.kind === "caption" && (target.label ?? "") === (type.label ?? "")
            : target.kind === type.kind,
        )
      : [];
    this.listSel.replaceChildren(...targets.map((target) => opt(target.listText, target.key)));
    // A native select shows its first option preselected — a standalone
    // listbox starts with nothing picked, so pick it here.
    const first = this.listSel.querySelector("fluent-option");
    if (first) this.#pickOption(this.listSel, first.getAttribute("value") ?? "");
  }

  /** The content options follow the selected target: an unnumbered heading
   *  has no heading number, a caption without text no caption text. A disabled
   *  pick moves to the first usable option (never a silent no-op insert). */
  #syncContentDisabled(type: TypeEntry | null, target: CrossReferenceTarget | null): void {
    const contents = type ? CROSS_REFERENCE_CONTENTS[type.kind] : [];
    const options = this.#options(this.contentSel);
    contents.forEach((content, i) => {
      const option = options[i];
      if (option) option.toggleAttribute("disabled", this.#contentDisabled(target, content));
    });
    const picked = this.#picked(this.contentSel);
    const pickedOption = options.find((option) => option.getAttribute("value") === picked);
    if (!pickedOption || pickedOption.hasAttribute("disabled")) {
      const first = options.find((option) => !option.hasAttribute("disabled"));
      if (first) this.#pickOption(this.contentSel, first.getAttribute("value") ?? "");
    }
  }

  #contentDisabled(target: CrossReferenceTarget | null, content: CrossRefContent): boolean {
    return !target || !crossReferenceContentAvailable(target, content);
  }

  /** What the reference reads right now — the dialog's "Refer to" preview. */
  #previewOf(target: CrossReferenceTarget, content: CrossRefContent): string {
    switch (content) {
      case "text":
      case "label":
        return target.text;
      case "entire":
        return target.captionFull ?? target.text;
      case "captionText":
        return target.captionText ?? "";
      case "number":
        return target.number ?? "";
      case "page":
        return String(target.page ?? 1);
      default:
        return this.#caretPos != null && target.pos < this.#caretPos
          ? t("crossRef.above", this)
          : t("crossRef.below", this);
    }
  }

  #applyLabels(type?: TypeEntry | null): void {
    if (this.dialogEl) this.dialogEl.heading = t("crossRef.title", this);
    if (this.typeLabel) this.typeLabel.textContent = t("crossRef.type", this);
    if (this.contentLabel) this.contentLabel.textContent = t("crossRef.content", this);
    if (this.okBtn) this.okBtn.textContent = t("crossRef.insert", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    if (this.whichLabel) {
      const kind = type?.kind;
      const key =
        kind === "heading"
          ? "crossRef.whichHeading"
          : kind === "numbered"
            ? "crossRef.whichNumbered"
            : kind === "footnote"
              ? "crossRef.whichFootnote"
              : kind === "endnote"
                ? "crossRef.whichEndnote"
                : kind === "bookmark"
                  ? "crossRef.whichBookmark"
                  : "crossRef.whichCaption";
      this.whichLabel.textContent = t(key, this);
    }
  }

  /** The dropdown's/listbox's `fluent-option`s in order (the control input is
   *  not an option). */
  #options(sel: FluentDropdown | FluentListbox | undefined): HTMLElement[] {
    const scope = sel?.querySelector?.("fluent-listbox") ?? sel;
    return [...(scope?.querySelectorAll<HTMLElement>("fluent-option") ?? [])];
  }

  /** Pick an option by value, clearing the rest first: `pick()` only sets the
   *  one, and FAST's `selectedOptions` lags freshly rendered children (its
   *  slotchange indexing is async — the field dialog documents the same). */
  #pickOption(sel: FluentDropdown | FluentListbox | undefined, value: string): void {
    for (const option of this.#options(sel))
      (option as HTMLElement & { selected?: boolean }).selected =
        option.getAttribute("value") === value;
    pick(sel as FluentDropdown, value);
  }

  /** The picked option value — read straight off the options' `selected`
   *  property (programmatic picks and user clicks both set it), then the
   *  indexed selectedOptions, then the dropdown's value/input mirror. */
  #picked(sel: FluentDropdown | FluentListbox | undefined): string | null {
    const selected = this.#options(sel).find(
      (option) => (option as HTMLElement & { selected?: boolean }).selected,
    );
    if (selected?.getAttribute("value")) return selected.getAttribute("value");
    const own = (sel as FluentListbox | undefined)?.selectedOptions?.[0]?.getAttribute("value");
    if (own) return own;
    const listbox = listboxOf(sel as FluentDropdown) as FluentListbox | null;
    const indexed = listbox?.selectedOptions?.[0]?.getAttribute("value");
    if (indexed) return indexed;
    return pickedValue(sel as FluentDropdown);
  }
}

export default DocenCrossReferenceDialog;
