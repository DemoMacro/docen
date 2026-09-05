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

/** What the Document Inspector found (Word's 检查问题): the comment count and
 *  the distinct revision-record count. The host rescans after each removal and
 *  re-hands the JSON in, so the rows fall to zero in place. */
export interface InspectFindings {
  comments: number;
  revisions: number;
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(360px, 92vw);
  }
  .inspect-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
  }
  .inspect-row {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 7px 2px;
    border-block-end: 1px solid var(--colorNeutralStroke2, #e0e0e0);
  }
  .inspect-row:last-child {
    border-block-end: none;
  }
  .inspect-row span {
    font-size: 13px;
    flex: 1;
  }
  .inspect-row b {
    font-size: 14px;
    font-variant-numeric: tabular-nums;
  }
`;

const template = html<DocenInspectDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="inspect-body" ${ref("bodyEl")}>
      <div class="inspect-row">
        <span></span>
        <b></b>
        <fluent-button
          appearance="outline"
          ${ref("commentsBtn")}
          @click="${(x) => x.$emit("inspect:clear-comments")}"
        ></fluent-button>
      </div>
      <div class="inspect-row">
        <span></span>
        <b></b>
        <fluent-button
          appearance="outline"
          ${ref("revisionsBtn")}
          @click="${(x) => x.$emit("inspect:accept-revisions")}"
        ></fluent-button>
      </div>
    </div>
    <div slot="action">
      <fluent-button
        appearance="accent"
        ${ref("closeBtn")}
        @click="${(x) => x.hide()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-inspect-dialog>` — MS Office "Document Inspector" (检查问题) for the
 * two content kinds the editor models. The host scans (comment count, distinct
 * revision records) and hands the JSON `findings` attribute in before `show()`;
 * each row's button emits `inspect:clear-comments` / `inspect:accept-revisions`
 * for the host to act on, and the host re-hands the scan so the counts fall to
 * zero (the buttons disable at zero, Word's "all clear" state).
 */
@customElement({ name: "docen-inspect-dialog", template, styles })
class DocenInspectDialog extends FASTElement {
  @attr findings?: string;

  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable bodyEl?: HTMLElement;
  @observable commentsBtn?: HTMLElement & { disabled: boolean; textContent: string };
  @observable revisionsBtn?: HTMLElement & { disabled: boolean; textContent: string };
  @observable closeBtn?: HTMLElement & { textContent: string };

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

  findingsChanged(): void {
    this.#renderRows();
  }

  show(): void {
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("inspect.title", this);
    if (this.commentsBtn) this.commentsBtn.textContent = t("inspect.remove", this);
    if (this.revisionsBtn) this.revisionsBtn.textContent = t("inspect.accept", this);
    if (this.closeBtn) this.closeBtn.textContent = t("inspect.close", this);
    this.#renderRows();
  }

  #renderRows(): void {
    if (!this.bodyEl) return;
    let findings: Partial<InspectFindings> = {};
    try {
      findings = this.findings ? (JSON.parse(this.findings) as Partial<InspectFindings>) : {};
    } catch {
      findings = {};
    }
    const rows = this.bodyEl.querySelectorAll(".inspect-row");
    const paint = (
      row: Element | undefined,
      label: string,
      count: number,
      btn?: HTMLElement & { disabled: boolean; textContent: string },
    ): void => {
      if (!row) return;
      const span = row.querySelector("span");
      const b = row.querySelector("b");
      if (span) span.textContent = label;
      if (b) b.textContent = String(count);
      if (btn) btn.disabled = count === 0;
    };
    paint(rows[0], t("inspect.comments", this), findings.comments ?? 0, this.commentsBtn);
    paint(rows[1], t("inspect.revisions", this), findings.revisions ?? 0, this.revisionsBtn);
  }
}

export default DocenInspectDialog;
