import {
  FASTElement,
  css,
  customElement,
  html,
  observable,
  ref,
  repeat,
} from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** One date/time format — the OOXML date-picture switch ("the instruction")
 *  and the Intl options that render today's sample the list shows. */
export interface DateFormatItem {
  /** OOXML date picture, e.g. `yyyy年M月d日` — goes after `DATE \@ `. */
  instruction: string;
  /** Renders the sample text (and the static insert). */
  intl: Intl.DateTimeFormatOptions;
}

/** Word's Date and Time formats, per offered language (most common first). */
export const DATE_FORMATS: Record<string, DateFormatItem[]> = {
  "zh-CN": [
    { instruction: "yyyy年M月d日", intl: { year: "numeric", month: "long", day: "numeric" } },
    { instruction: "yyyy/M/d", intl: { year: "numeric", month: "numeric", day: "numeric" } },
    {
      instruction: "yyyy年M月d日 HH:mm",
      intl: {
        year: "numeric",
        month: "long",
        day: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      },
    },
    { instruction: "HH:mm", intl: { hour: "2-digit", minute: "2-digit", hour12: false } },
  ],
  "en-US": [
    { instruction: "MMMM d, yyyy", intl: { year: "numeric", month: "long", day: "numeric" } },
    { instruction: "M/d/yyyy", intl: { year: "numeric", month: "numeric", day: "numeric" } },
    {
      instruction: "MMMM d, yyyy h:mm am/pm",
      intl: {
        year: "numeric",
        month: "long",
        day: "numeric",
        hour: "numeric",
        minute: "2-digit",
        hour12: true,
      },
    },
    { instruction: "h:mm am/pm", intl: { hour: "numeric", minute: "2-digit", hour12: true } },
  ],
};

/** The languages the dialog offers (Word lists the editing languages). */
export const DATE_LANGUAGES: readonly { tag: string; name: string }[] = [
  { tag: "zh-CN", name: "中文（中国）" },
  { tag: "en-US", name: "English (United States)" },
];

/** Today's text for a format (the list sample and the static insert). */
export function formatDateTime(tag: string, item: DateFormatItem): string {
  return new Intl.DateTimeFormat(tag, item.intl).format(new Date());
}

const langOptionTemplate = html<(typeof DATE_LANGUAGES)[number], DocenDateTimeDialog>`
  <fluent-option value="${(l) => l.tag}">${(l) => l.name}</fluent-option>
`;

const formatTemplate = html<{ tag: string; index: number }, DocenDateTimeDialog>`
  <fluent-option value="${(x) => x.index}">
    ${(x) =>
      formatDateTime(
        x.tag,
        DATE_FORMATS[x.tag][x.index] ?? DATE_FORMATS[x.tag][0] ?? { instruction: "", intl: {} },
      )}
  </fluent-option>
`;

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(320px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 14px;
    font-size: 13px;
  }
  fluent-dropdown,
  fluent-listbox {
    width: 100%;
    min-width: 0;
  }
  fluent-field {
    align-self: flex-start;
  }
`;

const template = html<DocenDateTimeDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <fluent-dropdown type="combobox" appearance="outline" ${ref("langDropdown")}>
        <fluent-listbox popover="manual" tabindex="-1">
          ${repeat(() => DATE_LANGUAGES, langOptionTemplate)}
        </fluent-listbox>
        <input
          slot="control"
          role="combobox"
          aria-haspopup="listbox"
          type="combobox"
          size="1"
          style="width:100%;box-sizing:border-box"
          ${ref("langInput")}
        />
      </fluent-dropdown>
      <fluent-listbox ${ref("formatListbox")}>
        ${repeat((x) => {
          const tag = x.langDropdown?.value ?? "zh-CN";
          return (DATE_FORMATS[tag] ?? DATE_FORMATS["zh-CN"]).map((_, index) => ({ tag, index }));
        }, formatTemplate)}
      </fluent-listbox>
      <fluent-field label-position="after">
        <fluent-checkbox slot="input" ${ref("autoBox")}></fluent-checkbox>
        <label slot="label">${(x) => t("dateDialog.updateAuto", x)}</label>
      </fluent-field>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.insertDate()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-date-time-dialog>` — Word's Date and Time dialog (Insert → Date
 * and Time): pick a language and format (the list shows today's sample),
 * optionally "update automatically" (a DATE field instead of static text).
 * Commits via `date-time:insert` `{ text, instruction? }` — `instruction`
 * carries the OOXML date picture when the field form was picked.
 */
@customElement({ name: "docen-date-time-dialog", template, styles })
class DocenDateTimeDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable langDropdown?: HTMLElement & { value: string | null };
  @observable langInput?: HTMLInputElement;
  @observable formatListbox?: HTMLElement & { value: string | null };
  @observable autoBox?: HTMLElement & { checked: boolean };
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

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

  show(): void {
    if (this.langDropdown && !this.langDropdown.value) this.langDropdown.value = "zh-CN";
    this.#selectFormat(0);
    if (this.autoBox) this.autoBox.checked = false;
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Mark only option `index` selected. The listbox's own `value` getter lags
   *  its options (FAST syncs it through selectedIndex bookkeeping), so the
   *  OK handler reads the DOM state this keeps authoritative. */
  #selectFormat(index: number): void {
    this.#formatOptions().forEach((opt, i) => {
      opt.selected = i === index;
    });
  }

  /** The selected format index (see #selectFormat for why not `listbox.value`). */
  #selectedFormat(): number {
    const at = this.#formatOptions().findIndex((opt) => opt.selected);
    return at >= 0 ? at : 0;
  }

  /** The dialog's format options, typed for their FAST `selected` flag. */
  #formatOptions(): { selected: boolean }[] {
    return [...(this.formatListbox?.querySelectorAll("fluent-option") ?? [])] as unknown as {
      selected: boolean;
    }[];
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  insertDate(): void {
    const tag = this.langDropdown?.value ?? "zh-CN";
    const formats = DATE_FORMATS[tag] ?? DATE_FORMATS["zh-CN"];
    const item = formats[Math.min(this.#selectedFormat(), formats.length - 1)];
    if (!item) return;
    this.$emit("date-time:insert", {
      text: formatDateTime(tag, item),
      ...(this.autoBox?.checked ? { instruction: `DATE \\@ "${item.instruction}"` } : {}),
    });
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("dateDialog.title", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenDateTimeDialog;
