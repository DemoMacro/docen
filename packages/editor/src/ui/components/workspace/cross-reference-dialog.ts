import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(440px, 92vw);
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
  .row select {
    flex: 1;
    min-width: 0;
  }
  .which {
    opacity: 0.75;
  }
  select.list {
    min-height: 168px;
    border: 1px solid var(--colorNeutralStroke1, #d1d1d1);
    border-radius: 4px;
    padding: 2px;
    font-size: 13px;
  }
`;

/** One referenceable document target (a _Ref caption bookmark or a user
 *  bookmark) as the dialog's candidate list shows it. */
export interface CrossReferenceTarget {
  /** The bookmark name — the REF field's payload. */
  name: string;
  /** The bookmark's inner text (a caption's "图 1", a bookmark's words). */
  text: string;
  kind: "caption" | "bookmark";
}

const template = html<DocenCrossReferenceDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="row">
        <label ${ref("typeLabel")}></label>
        <select ${ref("typeSel")} @change="${(x) => x.syncType()}">
          <option value="caption"></option>
          <option value="bookmark"></option>
        </select>
      </div>
      <div class="row">
        <label ${ref("contentLabel")}></label>
        <select ${ref("contentSel")}>
          <option value="text"></option>
          <option value="page"></option>
        </select>
      </div>
      <span class="which" ${ref("whichLabel")}></span>
      <select class="list" ${ref("listSel")}></select>
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
 * `<docen-cross-reference-dialog>` — Word's Cross-reference dialog (交叉引用),
 * scoped to the referenceable targets the document carries: caption `_Ref`
 * bookmarks and user bookmarks. The reference lands as a cached REF field
 * ("label and number" / "bookmark text") or a PAGEREF (page) — content choice
 * switches with the type. Opened with `show(targets)`; commits via
 * `cross-ref:ok` `{ name, content: "text" | "page" }` or cancels.
 */
@customElement({ name: "docen-cross-reference-dialog", template, styles })
class DocenCrossReferenceDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable typeLabel?: HTMLElement;
  @observable typeSel?: HTMLSelectElement;
  @observable contentLabel?: HTMLElement;
  @observable contentSel?: HTMLSelectElement;
  @observable whichLabel?: HTMLElement;
  @observable listSel?: HTMLSelectElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  /** The candidates the host passed to show() — filtered by the type. */
  #targets: CrossReferenceTarget[] = [];
  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#unobserveLang = observeLang(() => this.#applyLabels());
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  show(targets: CrossReferenceTarget[]): void {
    this.#targets = targets;
    if (this.typeSel) this.typeSel.value = "caption";
    this.#syncType();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible type change handler (FAST bindings can't reach
   *  `#`-private methods). */
  #syncType(): void {
    const kind = this.typeSel?.value === "bookmark" ? "bookmark" : "caption";
    // The content option's word follows the type, like Word's dialog
    // ("Label and number" for captions, "Bookmark text" for bookmarks).
    if (this.contentSel) {
      const text = this.contentSel.options[0];
      if (text)
        text.textContent = t(kind === "bookmark" ? "crossRef.refText" : "crossRef.refLabel", this);
      this.contentSel.value = "text";
    }
    if (this.whichLabel)
      this.whichLabel.textContent = t(
        kind === "bookmark" ? "crossRef.whichBookmark" : "crossRef.whichCaption",
        this,
      );
    this.#renderList();
  }

  syncType(): void {
    this.#syncType();
  }

  applyCrossReference(): void {
    const name = this.listSel?.value;
    if (!name) return;
    this.$emit("cross-ref:ok", {
      name,
      content: this.contentSel?.value === "page" ? "page" : "text",
    });
    this.hide();
  }

  #renderList(): void {
    if (!this.listSel) return;
    const kind = this.typeSel?.value === "bookmark" ? "bookmark" : "caption";
    this.listSel.replaceChildren(
      ...this.#targets
        .filter((target) => target.kind === kind)
        .map((target) => new Option(target.text || target.name, target.name)),
    );
    if (this.listSel.options.length > 0) this.listSel.selectedIndex = 0;
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("crossRef.title", this);
    if (this.typeLabel) this.typeLabel.textContent = t("crossRef.type", this);
    if (this.contentLabel) this.contentLabel.textContent = t("crossRef.content", this);
    if (this.okBtn) this.okBtn.textContent = t("crossRef.insert", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    if (this.typeSel) {
      const [caption, bookmark] = this.typeSel.options;
      if (caption) caption.textContent = t("crossRef.caption", this);
      if (bookmark) bookmark.textContent = t("crossRef.bookmark", this);
    }
    if (this.contentSel) {
      const [text, page] = this.contentSel.options;
      if (text)
        text.textContent = t(
          this.typeSel?.value === "bookmark" ? "crossRef.refText" : "crossRef.refLabel",
          this,
        );
      if (page) page.textContent = t("crossRef.refPage", this);
    }
  }
}

export default DocenCrossReferenceDialog;
