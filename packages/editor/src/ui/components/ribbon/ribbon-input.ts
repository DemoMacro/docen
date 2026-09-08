import { FASTElement, attr, css, customElement, html, ref } from "@microsoft/fast-element";

const styles = css`
  :host {
    display: inline-flex;
    width: 112px;
  }
  fluent-text-input {
    width: 100%;
    min-width: 0;
  }
`;

const template = html<DocenRibbonInput>`
  <fluent-text-input
    appearance="outline"
    part="field"
    value="${(x) => x.hostValue ?? ""}"
    :currentValue="${(x) => x.hostValue ?? ""}"
    ${ref("field")}
  ></fluent-text-input>
`;

/**
 * `<docen-ribbon-input value="4.39 厘米" event="drawing-width">` — Word's
 * numeric measure box (the Size groups): typeable text with no drop-down.
 * Enter or blur (when the text changed since the last commit) emits `command`
 * with `{ event, value }`; a `value` attribute change mirrors into the field
 * so the host can keep the box reporting the selection's live extent.
 */
@customElement({ name: "docen-ribbon-input", template, styles })
class DocenRibbonInput extends FASTElement {
  @attr({ attribute: "value" }) hostValue?: string;
  @attr event?: string;

  field?: HTMLElement & { currentValue: string };

  private lastCommit = "";

  connectedCallback(): void {
    super.connectedCallback();
    this.lastCommit = this.hostValue ?? "";
    // Enter commits the way Word's box does; blur commits only what changed.
    this.addEventListener("keydown", (e: KeyboardEvent) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      this.commit();
    });
    this.addEventListener("focusout", () => this.commit());
  }

  /** A mirrored extent also resets the commit baseline, so the commit that
   *  lands this value doesn't fire again on the blur that follows. */
  hostValueChanged(): void {
    this.lastCommit = this.hostValue ?? "";
  }

  private commit(): void {
    // currentValue is what the user sees mid-edit; `value` only updates when
    // Fluent's own commit machinery runs, which this box bypasses.
    const text = this.field?.currentValue ?? "";
    if (text === this.lastCommit) return;
    this.lastCommit = text;
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: { event: this.event ?? "", value: text, source: this },
      }),
    );
  }
}

export default DocenRibbonInput;
