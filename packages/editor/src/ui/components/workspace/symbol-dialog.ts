import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** The starter symbol palette — the characters Word's Symbol dialog offers
 *  that a document author most often reaches for (typography, math, arrows,
 *  currency, enclosed numerals, miscellany). Flat grid, no font switcher. */
const SYMBOLS = [
  "©",
  "®",
  "™",
  "§",
  "¶",
  "†",
  "‡",
  "•",
  "…",
  "‰",
  "°",
  "′",
  "″",
  "℃",
  "℉",
  "№",
  "±",
  "×",
  "÷",
  "≈",
  "≠",
  "≤",
  "≥",
  "∞",
  "√",
  "∑",
  "∏",
  "∫",
  "π",
  "µ",
  "Ω",
  "∆",
  "→",
  "←",
  "↑",
  "↓",
  "↔",
  "⇐",
  "⇒",
  "⇑",
  "⇓",
  "⇔",
  "∼",
  "∝",
  "∴",
  "∵",
  "⊂",
  "⊃",
  "€",
  "£",
  "¥",
  "¢",
  "₩",
  "₽",
  "①",
  "②",
  "③",
  "④",
  "⑤",
  "⑥",
  "⑦",
  "⑧",
  "⑨",
  "⑩",
  "★",
  "☆",
  "♦",
  "♠",
  "♣",
  "♥",
  "♪",
  "♫",
  "☀",
  "☂",
  "✓",
  "✗",
  "☐",
  "☑",
  "◼",
  "◻",
];

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(380px, 92vw);
  }
  .sym-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
  }
  /* Word's character grid: fixed square cells, hover + selected states. */
  .sym-grid {
    display: grid;
    grid-template-columns: repeat(8, 1fr);
    gap: 2px;
    max-height: 264px;
    overflow: auto;
  }
  .sym-grid button {
    aspect-ratio: 1;
    border: 1px solid transparent;
    border-radius: 4px;
    background: transparent;
    font-size: 18px;
    line-height: 1;
    cursor: pointer;
    font-family: inherit;
  }
  .sym-grid button:hover {
    background: var(--colorNeutralBackground1Hover, #f5f5f5);
  }
  .sym-grid button.selected {
    border-color: var(--colorBrandForeground1, #0f6cbd);
    background: var(--colorBrandBackground2, #ebf3fc);
  }
  /* Word's preview strip: the character large, its code point beside it. */
  .sym-preview {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 8px 10px;
    border: 1px solid var(--colorNeutralStroke2, #e0e0e0);
    border-radius: 4px;
    min-height: 56px;
    box-sizing: border-box;
  }
  .sym-preview .glyph {
    font-size: 34px;
    line-height: 1;
    min-width: 40px;
    text-align: center;
  }
  .sym-preview .codepoint {
    font-size: 12px;
    color: var(--colorNeutralForeground3, #616161);
    font-variant-numeric: tabular-nums;
  }
`;

const template = html<DocenSymbolDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="sym-body">
      <div class="sym-grid" ${ref("gridEl")}></div>
      <div class="sym-preview">
        <span class="glyph" ${ref("glyphEl")}></span>
        <span class="codepoint" ${ref("codepointEl")}></span>
      </div>
    </div>
    <div slot="action">
      <fluent-button
        appearance="accent"
        ${ref("insertBtn")}
        @click="${(x) => x.insertSymbol()}"
      ></fluent-button>
      <fluent-button ${ref("closeBtn")} @click="${(x) => x.hide()}"></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-symbol-dialog>` — MS Office "Symbol" dialog. A flat grid of common
 * Unicode symbols with a live preview; picking one and pressing Insert emits
 * `symbol:insert` { char } and stays open (Word keeps the dialog up so several
 * symbols can be inserted in a row). The host listens and inserts into the
 * editor at the caret.
 */
@customElement({ name: "docen-symbol-dialog", template, styles })
class DocenSymbolDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable gridEl?: HTMLElement;
  @observable glyphEl?: HTMLElement;
  @observable codepointEl?: HTMLElement;
  @observable insertBtn?: HTMLElement;
  @observable closeBtn?: HTMLElement;

  /** The symbol highlighted in the grid (also the Insert button's payload). */
  #selected = "";
  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#renderGrid();
    this.#unobserveLang = observeLang(() => this.#applyLabels());
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  show(): void {
    if (!this.#selected) this.#select(SYMBOLS[0]);
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  #select(char: string): void {
    this.#selected = char;
    if (this.glyphEl) this.glyphEl.textContent = char;
    if (this.codepointEl) {
      const code = char.codePointAt(0)?.toString(16).toUpperCase().padStart(4, "0") ?? "";
      this.codepointEl.textContent = `U+${code}  ${char}`;
    }
    this.gridEl?.querySelectorAll("button").forEach((b) => {
      b.classList.toggle("selected", b.textContent === char);
    });
  }

  #insert(): void {
    if (!this.#selected) return;
    this.$emit("symbol:insert", { char: this.#selected });
  }

  /** Template-visible Insert handler (FAST templates live outside the class,
   *  so a `#`-private method can't be referenced from the binding). */
  insertSymbol(): void {
    this.#insert();
  }

  #renderGrid(): void {
    if (!this.gridEl) return;
    this.gridEl.replaceChildren();
    for (const char of SYMBOLS) {
      const cell = document.createElement("button");
      cell.type = "button";
      cell.textContent = char;
      cell.addEventListener("click", () => this.#select(char));
      cell.addEventListener("dblclick", () => {
        this.#select(char);
        this.#insert();
      });
      this.gridEl.append(cell);
    }
    if (this.#selected) this.#select(this.#selected);
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("symbol.title", this);
    if (this.insertBtn) this.insertBtn.textContent = t("symbol.insert", this);
    if (this.closeBtn) this.closeBtn.textContent = t("symbol.close", this);
  }
}

export default DocenSymbolDialog;
