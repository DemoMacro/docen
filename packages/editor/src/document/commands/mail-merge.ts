import type { JSONContent } from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import { DocAttrStep } from "@tiptap/pm/transform";

import { t } from "../../ui";

/** The recipient data source pasted into the Select Recipients dialog — a
 *  CSV/TSV grid whose first row names the merge fields. Session-scoped: no
 *  OOXML part carries recipient rows, so it lives in doc attrs only. */
export interface MergeRecipients {
  headers: string[];
  rows: string[][];
}

/** The Start Mail Merge document kind — "letters" merges one copy (one
 *  section) per recipient; "directory" lists every record in one stream. */
export type MergeType = "letters" | "directory";

/** One quoted/literal CSV cell — reads until the delimiter or end of line. */
const readCell = (line: string, start: number, delimiter: string): [string, number] => {
  if (line[start] === '"') {
    let value = "";
    let i = start + 1;
    for (;;) {
      const quote = line.indexOf('"', i);
      if (quote < 0) return [value + line.slice(i), line.length];
      if (line[quote + 1] === '"') {
        value += line.slice(i, quote) + '"';
        i = quote + 2;
        continue;
      }
      // Closing quote — skip a trailing delimiter so the caller lands on the
      // next cell (the plain branch's `end + 1`).
      if (line[quote + 1] === delimiter) return [value + line.slice(i, quote), quote + 2];
      return [value + line.slice(i, quote), quote + 1];
    }
  }
  const end = line.indexOf(delimiter, start);
  return end < 0 ? [line.slice(start), line.length] : [line.slice(start, end), end + 1];
};

const splitLine = (line: string, delimiter: string): string[] => {
  const cells: string[] = [];
  let i = 0;
  for (;;) {
    const [value, next] = readCell(line, i, delimiter);
    cells.push(value.trim());
    if (next >= line.length) return cells;
    i = next;
  }
};

/** Parse pasted CSV/TSV text into the recipients grid — the delimiter is
 *  whichever the header row uses first (Excel clipboard copies are TSV). A
 *  header row plus at least one record is required. */
export function parseRecipients(text: string): MergeRecipients | null {
  const lines = text
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .filter((line) => line.trim() !== "");
  if (lines.length < 2) return null;
  const delimiter = lines[0].includes("\t") ? "\t" : ",";
  const headers = splitLine(lines[0], delimiter).filter(Boolean);
  if (headers.length === 0) return null;
  const rows = lines.slice(1).map((line) => {
    const cells = splitLine(line, delimiter);
    return headers.map((_, i) => cells[i] ?? "");
  });
  return { headers, rows };
}

/** ` MERGEFIELD Name \* MERGEFORMAT ` → "Name" (null when not a merge field). */
export function mergeFieldName(instruction: string): string | null {
  const match = /^\s*MERGEFIELD\s+(\S+)/i.exec(instruction);
  return match ? match[1] : null;
}

/** The merge field cached display — Word's «Name» chevron shape. */
export const mergeFieldDisplay = (name: string): string => `«${name}»`;

/** The seed for an inline MERGEFIELD — a simple field whose cached value is
 *  the chevron display (what Word shows before a merge). */
export const mergeFieldSeed = (name: string): JSONContent => ({
  type: "inlinePassthrough",
  attrs: {
    data: JSON.stringify({
      simpleField: {
        instruction: `MERGEFIELD ${name}`,
        cachedValue: mergeFieldDisplay(name),
      },
    }),
  },
});

/** Clone the doc with every MERGEFIELD's cached display swapped to the value
 *  the recipient row gives — the preview pass over the render JSON (and the
 *  Finish & Merge export). Fields the row doesn't name keep their chevrons. */
export function applyRecipientsRow(
  doc: JSONContent,
  recipients: MergeRecipients,
  rowIndex: number,
): JSONContent {
  const row = recipients.rows[rowIndex] ?? [];
  const values = new Map(
    recipients.headers.map((header, i) => [header.trim().toLowerCase(), row[i] ?? ""]),
  );
  const walk = (node: JSONContent): JSONContent => {
    const next: JSONContent = { ...node };
    const data = (node.attrs as { data?: unknown } | undefined)?.data;
    if (typeof data === "string") {
      try {
        const parsed = JSON.parse(data) as {
          simpleField?: { instruction?: string; cachedValue?: string };
        };
        const field = parsed.simpleField;
        const name = field ? mergeFieldName(field.instruction ?? "") : null;
        const value = name ? values.get(name.toLowerCase()) : undefined;
        if (field && name && value !== undefined) {
          next.attrs = {
            ...node.attrs,
            data: JSON.stringify({
              ...parsed,
              simpleField: { ...field, cachedValue: value },
            }),
          };
        }
      } catch {
        /* opaque payload — leave untouched */
      }
    }
    if (Array.isArray(node.content)) next.content = node.content.map(walk);
    return next;
  };
  return walk(doc);
}

/** The mail merge commands' view of the host — resolved per call so the
 *  controller can be built before a document opens. */
export interface MailMergeHost {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The story bridge — merge fields target the active story. */
  bridge(): { activeEditor(): Editor; focus(): void } | undefined;
  /** The host element — the i18n language source. */
  element(): HTMLElement;
  /** Re-run the live render (preview toggles and record navigation). */
  rerender(): void;
}

/**
 * The Mailings tab's merge commands, split out of the host element: the
 * recipient data source (doc.attrs.recipients), the MERGEFIELD seeds, the
 * preview index, and the merge assembly input for Finish & Merge (one
 * section per recipient — Word's Edit Individual Documents shape).
 */
export class MailMergeCommands {
  /** The record the preview shows — null when preview is off. */
  #previewIndex: number | null = null;

  constructor(private readonly host: MailMergeHost) {}

  #target(): Editor | null | undefined {
    return this.host.bridge()?.activeEditor() ?? this.host.editor();
  }

  /** The document's recipient data source — doc.attrs.recipients. */
  recipients(): MergeRecipients | null {
    const attrs = this.host.editor()?.state.doc.attrs as
      | { recipients?: MergeRecipients | null }
      | undefined;
    const { headers, rows } = attrs?.recipients ?? {};
    return headers && rows && rows.length > 0 ? { headers, rows } : null;
  }

  mergeType(): MergeType {
    const attrs = this.host.editor()?.state.doc.attrs as
      | { mergeType?: MergeType | null }
      | undefined;
    return attrs?.mergeType === "directory" ? "directory" : "letters";
  }

  /** The record the preview shows, clamped into range — null when off. */
  previewRow(): number | null {
    const recipients = this.recipients();
    if (this.#previewIndex === null || !recipients) return null;
    return Math.min(Math.max(this.#previewIndex, 0), recipients.rows.length - 1);
  }

  isPreviewing(): boolean {
    return this.#previewIndex !== null;
  }

  /** Recipients dialog commit — replace the data source. An empty grid
   *  clears the attr (null) and turns the preview off. */
  readonly onRecipientsOk = (event: Event): void => {
    const { recipients } = (event as CustomEvent<{ recipients?: MergeRecipients | null }>).detail;
    const target = this.#target();
    if (!target) return;
    this.#previewIndex = null;
    target.commands.command(({ tr }) => {
      tr.step(new DocAttrStep("recipients", recipients ?? null));
      return true;
    });
    this.host.rerender();
  };

  /** Start Mail Merge menu — record the document kind (letters vs directory;
   *  Finish & Merge shapes the output by it). */
  setMergeType(type: MergeType): void {
    const target = this.#target();
    if (!target) return;
    target.commands.command(({ tr }) => {
      tr.step(new DocAttrStep("mergeType", type));
      return true;
    });
  }

  /** Insert Merge Field — seed a MERGEFIELD simple field at the caret. */
  insertMergeField(name: string): void {
    const target = this.#target();
    if (!target || !name.trim()) return;
    const { from } = target.state.selection;
    target.view.dispatch(
      target.state.tr.insert(from, target.schema.nodeFromJSON(mergeFieldSeed(name.trim()))),
    );
    this.host.bridge()?.focus();
  }

  /** Address Block — one paragraph per recipient column (the full record as
   *  an address-style block; delete the lines a letter doesn't need). */
  insertAddressBlock(): void {
    const recipients = this.recipients();
    const target = this.#target();
    if (!target || !recipients) return;
    const { from } = target.state.selection;
    const paragraphs = recipients.headers.map((header) =>
      target.schema.nodeFromJSON({
        type: "paragraph",
        content: [mergeFieldSeed(header)],
      }),
    );
    target.view.dispatch(target.state.tr.insert(from, paragraphs));
    this.host.bridge()?.focus();
  }

  /** Greeting Line — "Dear «first column»," as its own paragraph. */
  insertGreetingLine(): void {
    const target = this.#target();
    if (!target) return;
    const element = this.host.element();
    const name = this.recipients()?.headers[0] ?? "Name";
    const paragraph = target.schema.nodeFromJSON({
      type: "paragraph",
      content: [
        { type: "text", text: `${t("merge.greetingPrefix", element)} ` },
        mergeFieldSeed(name),
        { type: "text", text: t("merge.greetingSuffix", element) },
      ],
    });
    const { from } = target.state.selection;
    target.view.dispatch(target.state.tr.insert(from, paragraph));
    this.host.bridge()?.focus();
  }

  /** Preview Results — toggle between the chevrons and the first record. */
  togglePreview(): void {
    if (this.#previewIndex === null) {
      if (!this.recipients()) return;
      this.#previewIndex = 0;
    } else {
      this.#previewIndex = null;
    }
    this.host.rerender();
  }

  firstRecord(): void {
    if (!this.recipients()) return;
    this.#previewIndex = 0;
    this.host.rerender();
  }

  lastRecord(): void {
    const recipients = this.recipients();
    if (!recipients) return;
    this.#previewIndex = recipients.rows.length - 1;
    this.host.rerender();
  }
}
