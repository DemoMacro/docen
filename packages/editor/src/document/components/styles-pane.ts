import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../ui/i18n/localize";

/** One list row: a style's id/name plus the run formatting its preview
 *  renders with (absent fields inherit — the pane renders them so). */
export interface StylePaneEntry {
  id: string;
  name: string;
  preview?: {
    font?: string;
    /** Points. */
    size?: number;
    bold?: boolean;
    italic?: boolean;
    /** Hex without "#". */
    color?: string;
    underline?: boolean;
  };
}

/** The Styles list the host pushes after each transaction. */
export interface StylesPaneState {
  entries: StylePaneEntry[];
  /** The caret paragraph's style id — the highlighted row. */
  currentId: string;
}

/** The Style Inspector view: the selection's style stack plus the direct
 *  formatting, pre-formatted by the host (it owns the i18n). */
export interface StylesInspectorData {
  paragraphStyle: string;
  characterStyle: string | null;
  direct: string[];
}

const styles = css`
  :host {
    display: flex;
    flex-direction: column;
    height: 100%;
    font-size: 12px;
  }
  .list {
    flex: 1 1 auto;
    overflow-y: auto;
    padding: 4px 0;
  }
  .entry {
    display: block;
    width: 100%;
    border: none;
    background: transparent;
    font: inherit;
    text-align: start;
    padding: 6px 10px;
    cursor: pointer;
    border-radius: 2px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .entry:hover {
    background: var(--docen-color-hover, rgba(0, 0, 0, 0.06));
  }
  .entry[data-current] {
    background: var(--docen-color-selected, rgba(0, 120, 212, 0.12));
  }
  .empty {
    color: #666;
    padding: 12px 10px;
  }
  .inspector {
    flex: 1 1 auto;
    overflow-y: auto;
    padding: 8px 10px;
    display: flex;
    flex-direction: column;
    gap: 10px;
  }
  .inspector h3 {
    font-size: 12px;
    margin: 0 0 2px;
    color: #666;
    font-weight: 600;
  }
  .inspector .value {
    font-size: 13px;
  }
  .inspector ul {
    margin: 0;
    padding-inline-start: 18px;
  }
  .inspector .none {
    color: #666;
  }
  .footer {
    flex: 0 0 auto;
    display: flex;
    gap: 4px;
    padding: 6px;
    border-block-start: 1px solid var(--docen-color-divider, #e1e1e1);
  }
  .footer fluent-button {
    flex: 1 1 0;
    min-width: 0;
  }
`;

const template = html<DocenStylesPane>`
  <div class="list" ${ref("listEl")} part="list"></div>
  <div class="footer">
    <fluent-button ${ref("modifyBtn")} @click="${(x) => x.emitModify()}"></fluent-button>
    <fluent-button ${ref("inspectorBtn")} @click="${(x) => x.toggleInspector()}"></fluent-button>
  </div>
`;

/**
 * `<docen-styles-pane>` — the Word Styles task pane: every paragraph style in
 * the document, each row previewed in its own formatting; clicking applies it
 * to the selection (Word's single-click behavior). The footer holds Modify
 * Style (opens `<docen-modify-style-dialog>` for the selected/current style)
 * and the Style Inspector toggle — the pane's second view, showing the
 * selection's style stack plus its direct formatting, pre-formatted by the
 * host. The host pushes data: `renderStyles(state)` after each transaction,
 * `renderInspector(data)` when the inspector view is showing, and
 * `setCurrent(id)` for a highlight-only refresh. Applies/emits:
 * `style-apply` (id), `modify-style` (id), `view` ("styles" | "inspector").
 */
@customElement({ name: "docen-styles-pane", template, styles })
class DocenStylesPane extends FASTElement {
  @observable listEl?: HTMLElement;
  @observable modifyBtn?: HTMLElement;
  @observable inspectorBtn?: HTMLElement;

  /** The view the footer's toggle switches to. The host listens for `view`
   *  and pushes the matching data. */
  mode: "styles" | "inspector" = "styles";
  /** The last-clicked style — the Modify button's target (defaults to the
   *  caret's current style). */
  #selectedId = "";

  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#unobserveLang = observeLang(() => this.#applyLabels());
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  renderStyles(state: StylesPaneState): void {
    this.mode = "styles";
    if (!this.listEl) return;
    this.#selectedId = state.currentId;
    const list = this.listEl;
    list.replaceChildren();
    for (const entry of state.entries) {
      const row = document.createElement("button");
      row.type = "button";
      row.className = "entry";
      row.dataset.styleId = entry.id;
      row.toggleAttribute("data-current", entry.id === state.currentId);
      const preview = entry.preview ?? {};
      row.textContent = entry.name;
      row.style.fontFamily = preview.font || "inherit";
      if (preview.size) row.style.fontSize = `${preview.size}pt`;
      row.style.fontWeight = preview.bold ? "600" : "inherit";
      row.style.fontStyle = preview.italic ? "italic" : "inherit";
      if (preview.color) row.style.color = `#${preview.color}`;
      row.style.textDecoration = preview.underline ? "underline" : "none";
      row.addEventListener("click", () => {
        this.#selectedId = entry.id;
        this.$emit("style-apply", entry.id);
      });
      list.append(row);
    }
    if (state.entries.length === 0) {
      const empty = document.createElement("p");
      empty.className = "empty";
      list.append(empty);
    }
    this.#applyLabels();
  }

  renderInspector(data: StylesInspectorData): void {
    this.mode = "inspector";
    if (!this.listEl) return;
    const root = this.listEl;
    root.replaceChildren();
    const wrap = document.createElement("div");
    wrap.className = "inspector";
    const styleHeading = document.createElement("h3");
    const styleValue = document.createElement("div");
    styleValue.className = "value";
    styleHeading.textContent = t("stylesPane.paragraphStyle", this);
    styleValue.textContent = data.paragraphStyle;
    wrap.append(styleHeading, styleValue);
    if (data.characterStyle) {
      const charHeading = document.createElement("h3");
      charHeading.textContent = t("stylesPane.characterStyle", this);
      const charValue = document.createElement("div");
      charValue.className = "value";
      charValue.textContent = data.characterStyle;
      wrap.append(charHeading, charValue);
    }
    const directHeading = document.createElement("h3");
    directHeading.textContent = t("stylesPane.direct", this);
    wrap.append(directHeading);
    if (data.direct.length > 0) {
      const ul = document.createElement("ul");
      for (const line of data.direct) {
        const li = document.createElement("li");
        li.textContent = line;
        ul.append(li);
      }
      wrap.append(ul);
    } else {
      const none = document.createElement("div");
      none.className = "none";
      none.textContent = t("stylesPane.directNone", this);
      wrap.append(none);
    }
    root.append(wrap);
    this.#applyLabels();
  }

  /** Highlight-only refresh after a transaction (the entries themselves are
   *  unchanged — the styles model is stable across edits). */
  setCurrent(id: string): void {
    if (!this.listEl) return;
    for (const row of this.listEl.querySelectorAll<HTMLElement>(".entry"))
      row.toggleAttribute("data-current", row.dataset.styleId === id);
  }

  /** Template-bound (public — the FAST template cannot reach a #private
   *  member): the footer buttons. */
  emitModify(): void {
    this.$emit("modify-style", this.#selectedId);
  }

  toggleInspector(): void {
    this.mode = this.mode === "styles" ? "inspector" : "styles";
    this.$emit("view", this.mode);
  }

  #applyLabels(): void {
    if (this.modifyBtn) this.modifyBtn.textContent = t("stylesPane.modify", this);
    if (this.inspectorBtn)
      this.inspectorBtn.textContent =
        this.mode === "styles"
          ? t("stylesPane.inspector", this)
          : t("stylesPane.backToStyles", this);
    const empty = this.listEl?.querySelector<HTMLElement>(".empty");
    if (empty) empty.textContent = t("stylesPane.empty", this);
  }
}

export default DocenStylesPane;
