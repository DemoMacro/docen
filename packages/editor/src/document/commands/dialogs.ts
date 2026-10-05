import type { ChartOptions, JSONContent, StylesOptions } from "@docen/docx";
import {
  buildCustomMultilevelLevels,
  buildListLevels,
  detectHeadingLevel,
  indexNumberings,
  nextMultilevelReference,
} from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import { formatNumber } from "@docen/layout";
import type { Node as PMNode } from "@tiptap/pm/model";
import { NodeSelection, TextSelection } from "@tiptap/pm/state";
import { DocAttrStep } from "@tiptap/pm/transform";

import type { FontDialogPatch } from "../../ui/components/workspace/font-dialog";
import { textCounter, wordCounter } from "../addin";
import {
  CROSS_REFERENCE_CONTENTS,
  crossReferenceContentAvailable,
  crossReferenceInstruction,
  type CrossRefContent,
  type CrossReferenceTarget,
} from "../cross-reference";
import type {
  ChartDataPatch,
  DrawingPropertiesPatch,
  ParagraphDialogPatch,
  TablePropertiesPatch,
} from "../extensions/commands";
import { collectListReferences } from "../extensions/commands";
import { noteRefId } from "../extensions/notes";
import {
  CAPTION_SEPARATOR_CHARS,
  customPropertiesOf,
  evaluateField,
  fieldRef,
  finiteNumber,
  parseFieldInstruction,
  seqChapterLevel,
  SEQ_NUMBER_FORMATS,
  type FieldBookmark,
  type FieldContext,
  type FieldFrame,
  type FieldRef,
} from "../fields";

/** One footnote/endnote body entry in documentExtras (`{ id, children }` with
 *  office-open's nested `paragraph`/run shapes — edited only as opaque text
 *  here). */
interface NoteEntry {
  id?: number;
  children?: unknown;
  [key: string]: unknown;
}

/** The slice of documentExtras the note commands touch. */
interface NoteExtras {
  footnotes?: NoteEntry[];
  endnotes?: NoteEntry[];
  contentTypes?: { overrides?: Array<{ partName?: string; contentType?: string }> };
  [key: string]: unknown;
}

/** The run between a caption's label+number and its text — Word's separator
 *  options (hyphen/period/colon/en dash/em dash) plus the classic space and
 *  CJK colon/period, stripped when the caption text starts. */
const CAPTION_SEPARATOR = /^[\s:：\-—–.。]+/;
const CAPTION_SEPARATOR_CHAR = /[\s:：\-—–.。]/;

/** Zero-width passthrough marks (w:bookmarkStart/End, comment range markers,
 *  proofing markers) — they occupy a position but carry no content, so a
 *  content-range scan must step past them. */
const MARK_PASSTHROUGH_KEYS = [
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "proofErr",
] as const;

/** Fields whose value reads the bookmark table — Update All Fields resolves
 *  them after the pass that renumbers captions (see the two phases). */
const BOOKMARK_FIELD_NAMES = new Set(["REF", "PAGEREF", "NOTEREF"]);

/** One caption label's settings entry (settings.xml `w:captions/w:caption`,
 *  carried in documentExtras.settings) — the per-label chapter-number shape
 *  Word persists alongside the inserted field. */
interface CaptionSetting {
  /** The label word (w:name). */
  name: string;
  /** Whether the label includes chapter numbers (w:chapNum). */
  chapterNumber?: boolean;
  /** The heading level chapters start at (w:heading, 1-9). */
  heading?: number;
  /** The chapter-number separator token (w:sep). */
  sep?: string;
  /** The label's number format (w:numFmt, ST_NumberFormat). */
  numFmt?: string;
}

/** The SEQ numbering state computed in one document-order walk — what the
 *  caption commit and Update All Fields read. */
interface SeqWalk {
  /** Field-atom position → the ordinal this occurrence takes. */
  ordinals: Map<number, number>;
  /** Field-atom position → the chapter number in effect for its `\s` level
   *  (only when a heading of that level precedes the field). */
  chapters: Map<number, string>;
  /** Label → occurrences since its last reset (the state at the walk's end;
   *  with a `before` bound, at that position). */
  counts: Map<string, number>;
  /** Heading level (1-9) → its chapter count at the walk's end. */
  chapterCounts: number[];
}

/** The dialog commands' view of the host — resolved per call so the controller
 *  can be built before a document opens (the editor and the story bridge both
 *  arrive later). */
export interface DialogsHost {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The story bridge — dialog commits target the active story (header/footer
   *  stories included); pageOf backs the PAGEREF cached values. */
  bridge():
    | { activeEditor(): Editor; focus(): void; pageOf(pos: number): number | null }
    | undefined;
  /** The host element — shadow-DOM root for the dialog components and the
   *  i18n language source. */
  element(): HTMLElement;
  /** Status-bar language mirror (Word shows the selection's language there) —
   *  re-run after the proofing-language commit. */
  syncStatusLanguage(): void;
  /** The open document's filename (FILENAME fields). Absent headless. */
  filename?: () => string | undefined;
  /** The rendered page frame at a position (PAGE/NUMPAGES/SECTION/… and
   *  PAGEREF) — the pagination's own view of the caret. Absent headless:
   *  page-dependent fields keep their cached value. */
  fieldFrame?: (pos: number) => FieldFrame | undefined;
  /** The UI language's "above"/"below" for the `\p` switch (Word renders the
   *  words in the document language). Absent = English. */
  positionTerms?: () => { above: string; below: string };
}

/**
 * The Home/References dialog commits, split out of the host element: the
 * paragraph/font/table-properties patches, the proofing language, the Chinese
 * layout pair (phonetic guide, two lines in one), the multilevel list
 * definition, the caption/cross-reference/bookmark field seeds, and the
 * general field insert/update pair — see the matching *-dialog.ts components
 * for the UI side.
 */
export class DialogCommands {
  constructor(private readonly host: DialogsHost) {}

  #target(): Editor | null | undefined {
    return this.host.bridge()?.activeEditor() ?? this.host.editor();
  }

  /** Paragraph dialog 确定 — stamp the committed patch onto the selection. */
  readonly onParagraphOk = (event: CustomEvent<ParagraphDialogPatch | undefined>): void => {
    const patch = event.detail;
    if (!patch) return;
    const target = this.#target();
    target?.commands["paragraph-dialog-apply"]?.(patch);
  };

  /** Paragraph dialog's Set As Default — write the patch onto the Normal
   *  style, so every paragraph inheriting it picks the values up. */
  readonly onParagraphDefault = (event: CustomEvent<ParagraphDialogPatch | undefined>): void => {
    const patch = event.detail;
    if (!patch) return;
    const target = this.#target();
    target?.commands["paragraph-dialog-default"]?.(patch);
  };

  // The Font dialog's OK — the patch is the selection's absolute run state
  // (Office commits the dialog atomically): everything lands in ONE chained
  // transaction, so a single undo reverts the whole dialog. (Separate
  // commands can't be fired in sequence off one cached commands object:
  // Tiptap's non-chain commands capture their transaction state once, so a
  // second dispatch applies a stale tr and PM throws "mismatched transaction".)
  readonly onFontOk = (event: CustomEvent<FontDialogPatch | undefined>): void => {
    const patch = event.detail;
    if (!patch) return;
    const target = this.#target();
    if (!target) return;
    const chain = target.chain();
    // Native attrs ride one textStyle setMark (attrNative null = absent).
    chain.setMark("textStyle", {
      font: patch.font,
      size: patch.size ? Number(patch.size) : null,
      doubleStrike: patch.doubleStrike || null,
      smallCaps: patch.smallCaps || null,
      allCaps: patch.allCaps || null,
      vanish: patch.hidden || null,
    });
    if (patch.bold) chain.setMark("bold");
    else chain.unsetMark("bold");
    if (patch.italic) chain.setMark("italic");
    else chain.unsetMark("italic");
    if (patch.strike) chain.setMark("strike");
    else chain.unsetMark("strike");
    if (patch.underlineStyle) chain["underline-style"](patch.underlineStyle, patch.underlineColor);
    else chain.unsetMark("underline");
    // Sub-/superscript are mutually exclusive marks — commit the checked one
    // and clear both when neither is.
    if (patch.superscript) chain.setMark("superscript");
    else if (patch.subscript) chain.setMark("subscript");
    else {
      chain.unsetMark("superscript");
      chain.unsetMark("subscript");
    }
    chain.run();
    this.host.bridge()?.focus();
  };

  /** Table Properties dialog 确定 — rewrite the caret table's alignment and
   *  left indent (the dialog prefills from the same table's attrs). */
  readonly onTablePropertiesOk = (event: CustomEvent<TablePropertiesPatch | undefined>): void => {
    const patch = event.detail;
    if (!patch) return;
    const target = this.#target();
    target?.commands["table-properties-apply"]?.(patch);
    this.host.bridge()?.focus();
  };

  /** Size-and-Position dialog 确定 — stamp the committed geometry (cm) onto
   *  the selected floating drawing. */
  readonly onDrawingPropertiesOk = (
    event: CustomEvent<DrawingPropertiesPatch | undefined>,
  ): void => {
    const patch = event.detail;
    if (!patch) return;
    const target = this.#target();
    target?.commands["drawing-properties-apply"]?.(patch);
    this.host.bridge()?.focus();
  };

  /** Chart Design tab → Edit Data — open the grid dialog prefilled from the
   *  selected chart's payload. Modal: the NodeSelection stays on the chart
   *  while the dialog is up, so the commit rides the chart-data-apply
   *  command (no pending-position bookkeeping needed). */
  chartEditAtSelection(): void {
    const target = this.#target();
    const sel = target?.state.selection;
    if (!sel || !(sel instanceof NodeSelection) || sel.node.type.name !== "chart") return;
    this.#chartDialog()?.show((sel.node.attrs.chart as ChartOptions | null) ?? null);
  }

  /** Edit Data dialog 确定 — stamp the grid's values onto the selected chart. */
  readonly onChartOk = (event: Event): void => {
    const { patch } = (event as CustomEvent<{ patch?: ChartDataPatch }>).detail ?? {};
    if (!patch) return;
    const target = this.#target();
    target?.commands["chart-data-apply"]?.(JSON.stringify(patch));
    this.host.bridge()?.focus();
  };

  #chartDialog(): { show(chart: ChartOptions | null): void } | null | undefined {
    return this.host.element().shadowRoot?.querySelector("docen-chart-data-dialog") as
      | { show(chart: ChartOptions | null): void }
      | null
      | undefined;
  }

  /** Language dialog 确定 — commit the selection's proofing language
   *  (w:lang). With a bare caret the setting rides the run the caret sits in
   *  (extended over the adjacent text sharing its marks) — the viewless input
   *  path has no stored-marks lane, so a caret-only commit would otherwise
   *  be dropped and the dialog would never stick. */
  readonly onLanguageOk = (event: Event): void => {
    const { value, noProof } = (event as CustomEvent<{ value?: string; noProof?: boolean }>)
      .detail ?? { value: undefined, noProof: false };
    const target = this.#target();
    if (!value || !target) return;
    const { selection } = target.state;
    const attrs = { language: { value }, noProof: noProof ? true : null };
    if (selection.empty) {
      const $from = selection.$from;
      const base = $from.marks();
      const sameRun = (node: PMNode) =>
        node.isText &&
        node.marks.length === base.length &&
        node.marks.every((m) => base.some((b) => b.eq(m)));
      let from = selection.from;
      let to = selection.from;
      const parentStart = $from.start();
      $from.parent.forEach((child, off) => {
        if (!sameRun(child)) return;
        const start = parentStart + off;
        const end = start + child.nodeSize;
        if (end <= selection.from) from = Math.min(from, start);
        else if (start <= selection.from) {
          from = Math.min(from, start);
          to = Math.max(to, end);
        }
      });
      target
        .chain()
        .setTextSelection({ from, to })
        .setMark("textStyle", attrs)
        .setTextSelection(selection.from)
        .run();
    } else {
      target.chain().setMark("textStyle", attrs).run();
    }
    this.host.bridge()?.focus();
    this.host.syncStatusLanguage();
  };

  // ── Phonetic guide (拼音指南) ──

  /** The selection's phonetic state for the dialog: the per-character text,
   *  the readings already on its runs (blank where unannotated), the first
   *  ruby mark's alignment, and the selection bounds. Null when the selection
   *  is empty, spans paragraphs, or holds anything but text (the guide splits
   *  the run per character — mixed content and cross-paragraph ranges don't
   *  split). */
  selectionPhonetic(): {
    chars: string[];
    readings: string[];
    alignment: string | null;
    from: number;
    to: number;
  } | null {
    const editor = this.#target();
    if (!editor) return null;
    const { from, to, empty, $from, $to } = editor.state.selection;
    if (empty || !$from.sameParent($to)) return null;
    const { doc } = editor.state;
    let plain = true;
    doc.nodesBetween(from, to, (node) => {
      // (nodesBetween yields the ancestors too — only an inline non-text node
      // inside the range blocks the split.)
      if (node.isInline && !node.isText) plain = false;
    });
    if (!plain) return null;
    const chars = doc.textBetween(from, to).split("");
    const readings = chars.map(() => "");
    let alignment: string | null = null;
    doc.nodesBetween(from, to, (node, pos) => {
      if (!node.isText) return;
      const ruby = (node.marks ?? []).find((m) => m.type.name === "ruby");
      if (!ruby) return;
      alignment ??= (ruby.attrs.alignment as string) ?? null;
      // This editor writes one node per base character carrying its whole
      // reading; a parsed multi-character node has no reliable per-character
      // split, so its reading lands whole on the first character.
      const start = Math.max(from, pos);
      const end = Math.min(to, pos + node.nodeSize);
      if (end > start && start - from < readings.length)
        readings[start - from] = String(ruby.attrs.text ?? "");
    });
    return { chars, readings, alignment, from, to };
  }

  /** Home → Font → Phonetic guide — open the per-character reading dialog
   *  (Word grays the button on an empty selection; a non-text or
   *  cross-paragraph selection is a no-op here). */
  phoneticOpen(): void {
    const dialog = this.host
      .element()
      .shadowRoot?.querySelector("docen-phonetic-dialog") as unknown as {
      show(chars: string[], readings: string[], alignment: string | null): void;
    } | null;
    const state = this.selectionPhonetic();
    if (!dialog || !state) return;
    dialog.show(state.chars, state.readings, state.alignment);
  }

  /** Phonetic dialog 确定 — split the selection into per-character runs, each
   *  carrying a ruby mark with its reading (a blank reading leaves that
   *  character unannotated). The base run's own marks ride every character;
   *  the annotation font is half the base size (Word's default). */
  readonly onPhoneticOk = (event: Event): void => {
    const { chars, readings, alignment } =
      (
        event as CustomEvent<{
          chars?: string[];
          readings?: string[];
          alignment?: string;
        }>
      ).detail ?? {};
    const target = this.#target();
    if (!target || !chars || !readings || chars.length === 0) return;
    const { from, to, empty, $from } = target.state.selection;
    if (empty) return;
    const carried = $from.marks();
    const baseSize =
      (carried.find((m) => m.type.name === "textStyle")?.attrs.size as number | null) ?? null;
    const { schema, tr } = target.state;
    const nodes = chars.map((ch, i) => {
      const marks = readings[i]
        ? [
            ...carried,
            schema.mark("ruby", {
              text: readings[i],
              alignment: alignment ?? "center",
              fontSize: baseSize != null ? Math.round(baseSize / 2) : null,
              baseFontSize: baseSize,
              raise: null,
              languageId: null,
              dirty: null,
            }),
          ]
        : carried;
      return schema.text(ch, marks);
    });
    const next = tr.replaceWith(from, to, nodes);
    next.setSelection(TextSelection.create(next.doc, from, from + chars.length));
    target.view.dispatch(next);
    this.host.bridge()?.focus();
  };

  /** Phonetic dialog 清除读音 — strip the ruby marks off the selection. */
  readonly onPhoneticClear = (): void => {
    const target = this.#target();
    if (!target || target.state.selection.empty) return;
    target.chain().unsetMark("ruby").run();
    this.host.bridge()?.focus();
  };

  // ── Two Lines in One (双行合一 / 合并字符, Home → Paragraph → Chinese Layout) ──

  /** The selection's two-in-one state for the dialog: its text and whether a
   *  bracket pair is already on. Null when the selection is empty. */
  selectionTwoInOne(): { text: string; brackets: boolean } | null {
    const editor = this.#target();
    if (!editor || editor.state.selection.empty) return null;
    const { from, to } = editor.state.selection;
    let brackets = false;
    editor.state.doc.nodesBetween(from, to, (node) => {
      if (!node.isText || brackets) return;
      const style = (node.marks ?? []).find((m) => m.type.name === "textStyle");
      const layout = style?.attrs.eastAsianLayout as
        | { combine?: unknown; combineBrackets?: unknown }
        | undefined;
      if (
        layout &&
        (layout.combine === true || layout.combine === "1") &&
        typeof layout.combineBrackets === "string" &&
        layout.combineBrackets !== "none"
      )
        brackets = true;
    });
    return { text: editor.state.doc.textBetween(from, to), brackets };
  }

  /** Home → Paragraph → Chinese Layout — open the two-lines-in-one dialog
   *  (Word grays the button on an empty selection). */
  twoInOneOpen(): void {
    const dialog = this.host
      .element()
      .shadowRoot?.querySelector("docen-two-in-one-dialog") as unknown as {
      show(text: string, brackets: boolean): void;
    } | null;
    const state = this.selectionTwoInOne();
    if (!dialog || !state) return;
    dialog.show(state.text, state.brackets);
  }

  /** Two-in-one dialog 确定 — stamp the eastAsianLayout combine mark on the
   *  selection; a dialog-edited text swaps the range for the new text carrying
   *  the selection's own marks (the combine attrs merged into its textStyle). */
  readonly onTwoInOneOk = (event: Event): void => {
    const { text, brackets } =
      (event as CustomEvent<{ text?: string; brackets?: boolean }>).detail ?? {};
    const target = this.#target();
    if (!target || !text || target.state.selection.empty) return;
    const { from, to, $from } = target.state.selection;
    const layout = { combine: true, combineBrackets: brackets ? "round" : null };
    if (text !== target.state.doc.textBetween(from, to)) {
      const { schema } = target.state;
      const carried = $from
        .marks()
        .map((m) =>
          m.type.name === "textStyle" ? schema.mark("textStyle", { ...m.attrs, ...layout }) : m,
        );
      target
        .chain()
        .command(({ tr }) => {
          tr.replaceWith(from, to, schema.text(text, carried));
          return true;
        })
        .run();
    } else {
      target.chain().setMark("textStyle", { eastAsianLayout: layout }).run();
    }
    this.host.bridge()?.focus();
  };

  /** Home → Paragraph → Multilevel — open the Define New Multilevel List
   *  dialog (no prefill: the editor has no current-list readback, the dialog
   *  resets to a fresh cascade). */
  defineListOpen(): void {
    const dialog = this.host
      .element()
      .shadowRoot?.querySelector("docen-define-list-dialog") as unknown as {
      show(): void;
    } | null;
    dialog?.show();
  }

  /** Define-list dialog 确定 — register the defined levels as a document
   *  numbering definition (compile passes doc.attrs.numbering straight
   *  through, and a "custom" reference is not a generated one, so the
   *  definition is never rebuilt) and stamp the selection's paragraphs with
   *  the fresh reference, all in one transaction. */
  readonly onDefineListOk = (event: Event): void => {
    const { levels } =
      (event as CustomEvent<{ levels?: { format: string; text: string }[] }>).detail ?? {};
    const target = this.#target();
    if (!target || !levels?.length) return;
    const { state } = target;
    const reference = nextMultilevelReference(
      collectListReferences(state.doc),
      (state.doc.attrs as { numbering?: unknown }).numbering,
      "custom",
    );
    const numbering = (state.doc.attrs as { numbering?: { abstractNumberings?: unknown[] } })
      .numbering;
    target
      .chain()
      .command(({ tr }) => {
        // Doc attrs sit outside the position space (doc.nodeAt(0) is the
        // first block), so they go through the dedicated doc-attr step; the
        // stamps below keep their positions either way.
        tr.step(
          new DocAttrStep("numbering", {
            ...numbering,
            abstractNumberings: [
              ...(numbering?.abstractNumberings ?? []),
              { reference, levels: buildCustomMultilevelLevels(levels) },
            ],
          }),
        );
        state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
          if (node.type.name !== "paragraph") return true;
          const attrs = node.attrs as Record<string, unknown>;
          const level = (attrs.numbering as { level?: number } | null | undefined)?.level;
          tr.setNodeMarkup(pos, undefined, {
            ...attrs,
            bullet: null,
            numbering: { reference, level: typeof level === "number" ? level : 0 },
          });
          return false;
        });
        return true;
      })
      .run();
    this.host.bridge()?.focus();
  };

  // ── Caption / Cross-reference / Bookmark ──

  /** The next free bookmark id — one past the highest id already carried by
   *  any bookmarkStart/bookmarkEnd passthrough in the document (OOXML marks
   *  pairs by a document-unique integer). */
  nextBookmarkId(target: Editor): number {
    let max = -1;
    const scan = (node: JSONContent): void => {
      for (const child of node.content ?? []) {
        if (child.type === "inlinePassthrough" || child.type === "passthrough") {
          try {
            const data = JSON.parse(String(child.attrs?.data ?? "{}")) as {
              bookmarkStart?: { id?: number };
              bookmarkEnd?: { id?: number };
            };
            for (const id of [data.bookmarkStart?.id, data.bookmarkEnd?.id]) {
              if (typeof id === "number" && id > max) max = id;
            }
          } catch {
            // opaque verbatim blobs without bookmark data — skip
          }
        }
        scan(child);
      }
    };
    scan(target.getJSON());
    return max + 1;
  }

  /** The SEQ counters in one document-order walk. Word's rules: a heading
   *  advances its level's chapter count (deeper levels reset), and a label
   *  whose SEQ field carries `\s <level>` restarts its sequence at every
   *  heading at or above that level — so captions renumber from the document
   *  itself, not from their caches. The reset lands on the heading, before
   *  any later field is visited: a `before` bound (the caption commit's
   *  insertion point, F9's caret) must still see the pending reset of its
   *  chapter, or the first caption of a new chapter would continue the
   *  previous count.
   *
   *  The chapter number is the heading's occurrence ordinal: the runtime model
   *  carries no heading number (list numbering is the projection's).
   *  Continuous heading numbering matches Word; multilevel prefixes (`1.1`)
   *  and unnumbered headings differ from Word's STYLEREF rendering. */
  #seqWalk(editor: Editor, before?: number): SeqWalk {
    const styles = (editor.state.doc.attrs as { styles?: StylesOptions }).styles;
    const ordinals = new Map<number, number>();
    const chapters = new Map<number, string>();
    const counts = new Map<string, number>();
    const chapterCounts = Array.from({ length: 10 }, () => 0);
    // Label → the `\s` level of its most recent switched field: a heading at
    // or above that level restarts the label's sequence. Recorded from the
    // fields visited so far, so a heading resets every label whose sequence
    // has already begun.
    const resetLevels = new Map<string, number>();
    editor.state.doc.descendants((node, pos) => {
      if (before != null && pos >= before) return false;
      if (node.type.name === "paragraph") {
        const level = detectHeadingLevel(
          {
            heading: (node.attrs.heading as string) || undefined,
            style: (node.attrs.style as string) || undefined,
            outlineLevel: node.attrs.outlineLevel as number | undefined,
          },
          styles,
        );
        if (level != null) {
          chapterCounts[level] = (chapterCounts[level] ?? 0) + 1;
          for (let l = level + 1; l <= 9; l++) chapterCounts[l] = 0;
          for (const [label, resetLevel] of resetLevels) {
            if (resetLevel >= level) counts.delete(label);
          }
        }
        return true;
      }
      if (node.type.name !== "inlinePassthrough") return true;
      const data = node.attrs.data;
      if (typeof data !== "string") return true;
      let field: ReturnType<typeof parseFieldInstruction>;
      try {
        const branch = JSON.parse(data) as Record<string, unknown>;
        const ref = fieldRef(branch);
        if (!ref?.instruction) return true;
        field = parseFieldInstruction(ref.instruction);
      } catch {
        return true;
      }
      if (field.name !== "SEQ") return true;
      const label = field.args[0];
      if (!label) return true;
      const level = seqChapterLevel(field.switches.s);
      if (level != null) resetLevels.set(label, level);
      const ordinal = (counts.get(label) ?? 0) + 1;
      counts.set(label, ordinal);
      ordinals.set(pos, ordinal);
      if (level != null && chapterCounts[level]! > 0)
        chapters.set(pos, String(chapterCounts[level]));
      return true;
    });
    return { ordinals, chapters, counts, chapterCounts };
  }

  /** The document's caption settings (settings.xml `w:captions`) — the
   *  per-label chapter-number shape the caption commit persists. */
  #captionSettings(editor: Editor): CaptionSetting[] {
    const extras = (editor.state.doc.attrs ?? {}) as {
      documentExtras?: { settings?: { captions?: { captions?: unknown } } };
    };
    const list = extras.documentExtras?.settings?.captions?.captions;
    if (!Array.isArray(list)) return [];
    return list.filter(
      (entry): entry is CaptionSetting =>
        !!entry && typeof entry === "object" && typeof (entry as CaptionSetting).name === "string",
    );
  }

  /** The document's caption separators as characters: label → separator
   *  (Word's "Use separator"). Labels with no setting keep the evaluator's
   *  hyphen default. */
  #captionSeparators(editor: Editor): ReadonlyMap<string, string> | undefined {
    const map = new Map<string, string>();
    for (const setting of this.#captionSettings(editor)) {
      const char = setting.sep ? CAPTION_SEPARATOR_CHARS[setting.sep] : undefined;
      if (char) map.set(setting.name, char);
    }
    return map.size > 0 ? map : undefined;
  }

  /** The caption settings with one label's entry upserted (w:caption) — the
   *  rest of documentExtras/settings passes through verbatim. */
  #patchCaptionSettings(editor: Editor, entry: CaptionSetting): Record<string, unknown> {
    const extras = {
      ...(editor.state.doc.attrs.documentExtras as Record<string, unknown> | undefined),
    };
    const settings = { ...(extras.settings as Record<string, unknown> | undefined) };
    const captions = { ...(settings.captions as Record<string, unknown> | undefined) };
    const list = Array.isArray(captions.captions)
      ? ([...captions.captions] as Record<string, unknown>[])
      : [];
    const index = list.findIndex((c) => c?.name === entry.name);
    if (index >= 0) list[index] = { ...list[index], ...entry };
    else list.push({ ...entry });
    captions.captions = list;
    settings.captions = captions;
    extras.settings = settings;
    return extras;
  }

  // ── Cross-reference targets (交叉引用) ──
  // The scan classifies each paragraph: caption (Caption style / SEQ field),
  // heading (detectHeadingLevel), numbered item (a resolvable list number),
  // otherwise the bookmarks it carries are user targets. Note reference atoms
  // are their own targets. Word creates a hidden `_Ref…` bookmark around a
  // referenced heading/numbered item on first use; the commit does the same
  // (onCrossRefOk), while captions/user bookmarks keep their existing one.

  /** One bookmark pair inside a paragraph: name plus the inner (start..end)
   *  range. */
  #bookmarkPairs(node: PMNode, pos: number): { name: string; from: number; to: number }[] {
    const open: { id?: number; name: string; from: number }[] = [];
    const out: { name: string; from: number; to: number }[] = [];
    node.forEach((child, offset) => {
      if (child.type.name !== "inlinePassthrough") return;
      const at = pos + 1 + offset;
      try {
        const data = JSON.parse(String(child.attrs.data ?? "{}")) as {
          bookmarkStart?: { id?: number; name?: string };
          bookmarkEnd?: { id?: number };
        };
        if (data.bookmarkStart?.name) {
          open.push({
            id: data.bookmarkStart.id,
            name: data.bookmarkStart.name,
            from: at + child.nodeSize,
          });
          return;
        }
        if (data.bookmarkEnd) {
          const index =
            data.bookmarkEnd.id != null
              ? open.findIndex((entry) => entry.id === data.bookmarkEnd!.id)
              : 0;
          const hit = open[index >= 0 ? index : 0];
          if (!hit) return;
          out.push({ name: hit.name, from: hit.from, to: at });
          open.splice(index >= 0 ? index : 0, 1);
        }
      } catch {
        // opaque verbatim blob — not a bookmark pair
      }
    });
    return out;
  }

  /** The plain text inside [from, to): text runs, field caches, and the note
   *  reference atoms' displayed numbers (the projection's first-reference
   *  ordinal — endnotes keep their lowercase-Roman look). */
  #crossRefText(
    node: PMNode,
    pos: number,
    from: number,
    to: number,
    notes: Map<number, { kind: "footnote" | "endnote"; id: number; display: string }>,
  ): string {
    let text = "";
    node.forEach((child, offset) => {
      const at = pos + 1 + offset;
      const end = at + child.nodeSize;
      if (at < from || end > to) return;
      if (child.type.name === "text") {
        text += child.textContent;
        return;
      }
      if (child.type.name !== "inlinePassthrough") return;
      const note = notes.get(at);
      if (note) {
        text += note.display;
        return;
      }
      const data = child.attrs.data;
      if (typeof data !== "string") return;
      try {
        const ref = fieldRef(JSON.parse(data) as Record<string, unknown>);
        if (ref?.result != null) text += ref.result;
      } catch {
        // opaque verbatim blob — not a field
      }
    });
    return text;
  }

  /** The paragraph's first SEQ field's label ("Figure" in `SEQ Figure \* …`),
   *  or null — the caption scan's rule (same read as the table-of-figures). */
  #seqLabelOf(node: PMNode): string | null {
    let label: string | null = null;
    node.forEach((child) => {
      if (label || child.type.name !== "inlinePassthrough") return;
      try {
        const data = JSON.parse(String(child.attrs.data ?? "{}")) as {
          simpleField?: { instruction?: string };
        };
        const m = /^SEQ\s+(\S+)/i.exec((data.simpleField?.instruction ?? "").trim());
        if (m) label = m[1]!;
      } catch {
        // opaque payload — not a SEQ field
      }
    });
    return label;
  }

  /** The end position of the paragraph's first SEQ field (inclusive), or
   *  null — the label+number range of a bookmark-less caption. */
  #seqEnd(node: PMNode, pos: number): number | null {
    let end: number | null = null;
    node.forEach((child, offset) => {
      if (end != null || child.type.name !== "inlinePassthrough") return;
      try {
        const data = JSON.parse(String(child.attrs.data ?? "{}")) as {
          simpleField?: { instruction?: string };
        };
        if (/^SEQ\s/i.test((data.simpleField?.instruction ?? "").trim()))
          end = pos + 1 + offset + child.nodeSize;
      } catch {
        // opaque payload — not a SEQ field
      }
    });
    return end;
  }

  /** The absolute position of a caption's first non-separator character at or
   *  after `from` (Word's "Only caption text" starts past the separator).
   *  Zero-width marks — the label's `bookmarkEnd` sits exactly at `from` —
   *  are not text; the scan steps past them to the first real character (or
   *  inline atom), so the resulting bookmark never carries the separator. */
  #captionTextStart(node: PMNode, pos: number, from: number): number | undefined {
    let start: number | undefined;
    node.forEach((child, offset) => {
      if (start != null) return;
      const at = pos + 1 + offset;
      const end = at + child.nodeSize;
      if (end <= from) return;
      if (child.type.name === "text") {
        const text = child.textContent;
        for (let i = Math.max(0, from - at); i < text.length; i++) {
          if (!CAPTION_SEPARATOR_CHAR.test(text[i]!)) {
            start = at + i;
            return;
          }
        }
        return;
      }
      if (child.type.name === "inlinePassthrough") {
        const data = child.attrs.data;
        if (typeof data === "string") {
          try {
            const branch = JSON.parse(data) as Record<string, unknown>;
            if (MARK_PASSTHROUGH_KEYS.some((key) => key in branch)) return;
          } catch {
            /* opaque verbatim blob — treated as content below */
          }
        }
      }
      if (at >= from) start = at;
    });
    return start;
  }

  /** Paragraph position → its rendered list number ("1", "1.1", "第1章"),
   *  trailing periods stripped (Word's REF `\n` shape). The counter walk
   *  mirrors the projection's marker substitution (per-reference counters,
   *  deeper levels reset) over the doc's numbering definitions plus the
   *  generated ones editor-created lists compile to. Style-chain numbering
   *  is not resolved — the runtime model carries it only in the style
   *  definition the per-paragraph projection expands, so a heading numbered
   *  through its style has no `\n` value here (documented deferral). */
  #paragraphNumbers(doc: PMNode): Map<number, string> {
    const definitions = indexNumberings((doc.attrs as { numbering?: unknown }).numbering);
    const generated = new Map<string, { format: string; text: string }[]>();
    const levelsOf = (reference: string): { format: string; text: string }[] => {
      const own = definitions.get(reference);
      if (own) return own;
      let built = generated.get(reference);
      if (!built) {
        built = (buildListLevels(reference) ?? []).map((level) => ({
          format: level.format != null ? String(level.format) : "decimal",
          text: level.text ?? "",
        }));
        if (built.length > 0) generated.set(reference, built);
      }
      return built;
    };
    const counters = new Map<string, number[]>();
    const out = new Map<number, string>();
    doc.descendants((node, pos) => {
      if (node.type.name !== "paragraph") return true;
      const attrs = node.attrs as Record<string, unknown>;
      // The built-in bullet sugar has no level definition and no number.
      if (attrs.bullet) return true;
      const reference = (attrs.numbering as { reference?: string } | null | undefined)?.reference;
      if (typeof reference !== "string" || !reference) return true;
      const level = Number((attrs.numbering as { level?: number }).level) || 0;
      const index = Math.max(0, Math.min(8, Math.trunc(level)));
      const levels = levelsOf(reference);
      const def = levels[index];
      if (!def || def.format === "bullet" || def.format === "none") return true;
      const counts = counters.get(reference) ?? [];
      counters.set(reference, counts);
      counts[index] = (counts[index] ?? 0) + 1;
      counts.length = index + 1;
      const marker = (def.text || "%1.").replace(/%([1-9])/g, (_, k: string) => {
        const at = Number(k) - 1;
        return formatNumber(levels[at]?.format ?? def.format, counts[at] ?? 1);
      });
      out.set(pos, marker.replace(/[.。]+$/, ""));
      return true;
    });
    return out;
  }

  /** The document's cross-reference candidates: caption, heading, numbered
   *  item, bookmark, footnote, and endnote targets in document order, each
   *  with its referenceable text and the range a hidden bookmark wraps. */
  crossReferenceTargets(): CrossReferenceTarget[] {
    const editor = this.#target();
    if (!editor) return [];
    return this.#crossReferenceTargetsIn(editor, editor.state.doc);
  }

  /** The candidate scan over one document snapshot — Update All Fields runs
   *  it again against the transaction's phase-1 doc so references quote the
   *  numbering just written (positions are stable: attribute patches keep
   *  node sizes). */
  #crossReferenceTargetsIn(editor: Editor, doc: PMNode): CrossReferenceTarget[] {
    const attrs = doc.attrs as { styles?: StylesOptions; documentExtras?: unknown };
    const styles = attrs.styles;
    const extras = (attrs.documentExtras ?? {}) as NoteExtras;
    const numbers = this.#paragraphNumbers(doc);
    const ordinals = { footnote: new Map<number, number>(), endnote: new Map<number, number>() };
    const out: CrossReferenceTarget[] = [];
    const pageAt = (pos: number): number | undefined => {
      const frame = this.#frameAt(editor, pos);
      const physical = this.host.bridge()?.pageOf(pos);
      return frame?.page ?? (typeof physical === "number" ? physical + 1 : undefined);
    };

    doc.descendants((node, pos) => {
      if (node.type.name !== "paragraph") return true;
      const nodeAttrs = node.attrs as Record<string, unknown>;
      const pairs = this.#bookmarkPairs(node, pos);
      const contentFrom = pos + 1;
      const contentTo = pos + 1 + node.content.size;
      const page = pageAt(pos);

      // Note reference atoms: the projection's first-reference-order ordinal
      // becomes the displayed number (`NOTEREF`'s value).
      const notes = new Map<
        number,
        { kind: "footnote" | "endnote"; id: number; display: string }
      >();
      node.forEach((child, offset) => {
        if (child.type.name !== "inlinePassthrough") return;
        const data = child.attrs.data;
        if (typeof data !== "string") return;
        let ref: ReturnType<typeof noteRefId>;
        try {
          ref = noteRefId(JSON.parse(data) as Record<string, unknown>);
        } catch {
          return;
        }
        if (!ref || ref.id == null) return;
        const kind = ref.kind === "endnoteReference" ? "endnote" : "footnote";
        const seen = ordinals[kind];
        let ordinal = seen.get(ref.id);
        if (ordinal == null) {
          ordinal = seen.size + 1;
          seen.set(ref.id, ordinal);
        }
        notes.set(pos + 1 + offset, {
          kind,
          id: ref.id,
          display: kind === "footnote" ? String(ordinal) : formatNumber("lowerRoman", ordinal),
        });
      });
      const texts = pairs.map((pair) => this.#crossRefText(node, pos, pair.from, pair.to, notes));

      const seqLabel = this.#seqLabelOf(node);
      const heading = detectHeadingLevel(
        {
          heading: (nodeAttrs.heading as string) || undefined,
          style: (nodeAttrs.style as string) || undefined,
          outlineLevel: nodeAttrs.outlineLevel as number | undefined,
        },
        styles,
      );
      const number = numbers.get(pos);
      const caption = nodeAttrs.style === "Caption" || seqLabel != null;
      const used = new Set<string>();

      if (caption) {
        // The structural label+number range runs to the SEQ field's end; the
        // `_Ref` pair wrapping exactly that range is the caption's own pair
        // (Word's shape). Other hidden pairs (an "entire caption" or "only
        // caption text" reference's own bookmark) do not displace it.
        const seqEnd = this.#seqEnd(node, pos);
        const labelRangeEnd = seqEnd ?? contentTo;
        // The caption's own pair ends exactly at the SEQ field (Word wraps
        // label+number); an "entire caption" or "only caption text" reference's
        // bookmark ends at the paragraph end and does not displace it.
        const labelPair = pairs
          .filter((entry) => entry.name.startsWith("_Ref") && entry.to === labelRangeEnd)
          .sort((a, b) => a.from - b.from)[0];
        const pair = labelPair ?? pairs.find((entry) => entry.name.startsWith("_Ref")) ?? pairs[0];
        const labelEnd = labelPair ? labelPair.to : (seqEnd ?? pair?.to ?? contentTo);
        const index = pair ? pairs.indexOf(pair) : -1;
        const labelText =
          index >= 0
            ? texts[index]!.trim()
            : this.#crossRefText(node, pos, contentFrom, labelEnd, notes).trim();
        const rawAfter = this.#crossRefText(node, pos, labelEnd, contentTo, notes);
        const captionText = rawAfter.replace(CAPTION_SEPARATOR, "");
        if (pair) used.add(pair.name);
        out.push({
          key: `caption:${pair?.name ?? pos}`,
          kind: "caption",
          name: pair?.name ?? "",
          ...(labelPair ? { labelName: labelPair.name } : {}),
          ...(seqLabel ? { label: seqLabel } : {}),
          pos,
          text: labelText,
          captionFull: `${labelText}${rawAfter}`.trim(),
          ...(captionText
            ? {
                captionText,
                captionTextFrom: this.#captionTextStart(node, pos, labelEnd),
              }
            : {}),
          ...(number ? { number } : {}),
          ...(page != null ? { page } : {}),
          listText: `${labelText}${rawAfter}`.trim(),
          contentFrom,
          contentTo,
          labelTo: labelEnd,
          ...(pair ? { bookmarkFrom: pair.from, bookmarkTo: pair.to } : {}),
        });
      }

      if (heading != null && !caption) {
        const index = Math.max(
          0,
          pairs.findIndex((pair) => pair.name.startsWith("_Ref")),
        );
        const pair = pairs[index];
        if (pair) used.add(pair.name);
        out.push({
          key: `heading:${pos}`,
          kind: "heading",
          name: pair?.name ?? "",
          pos,
          text: pair
            ? this.#crossRefText(node, pos, pair.from, pair.to, notes).trim()
            : node.textContent.trim(),
          ...(number ? { number } : {}),
          ...(page != null ? { page } : {}),
          listText: node.textContent.trim(),
          contentFrom,
          contentTo,
          ...(pair ? { bookmarkFrom: pair.from, bookmarkTo: pair.to } : {}),
        });
      }

      if (number != null && !caption) {
        const index = Math.max(
          0,
          pairs.findIndex((pair) => pair.name.startsWith("_Ref")),
        );
        const pair = pairs[index];
        if (pair) used.add(pair.name);
        out.push({
          key: `numbered:${pos}`,
          kind: "numbered",
          name: pair?.name ?? "",
          pos,
          text: node.textContent.trim(),
          number,
          ...(page != null ? { page } : {}),
          listText: `${number} ${node.textContent.trim()}`.trim(),
          contentFrom,
          contentTo,
          ...(pair ? { bookmarkFrom: pair.from, bookmarkTo: pair.to } : {}),
        });
      }

      // Note reference atoms are their own targets (an existing bookmark pair
      // around the atom supplies the name).
      for (const [at, note] of notes) {
        const pair = pairs.find((entry) => at >= entry.from && at < entry.to);
        if (pair) used.add(pair.name);
        const notePage = pageAt(at);
        const noteEntry = (note.kind === "footnote" ? extras.footnotes : extras.endnotes)?.find(
          (entry) => entry.id === note.id,
        );
        const body = noteEntry
          ? this.#noteTextOf(noteEntry.children)
              .split("\n")
              .find((line) => line.trim() !== "")
              ?.trim()
          : "";
        out.push({
          key: `${note.kind}:${at}`,
          kind: note.kind,
          name: pair?.name ?? "",
          pos: at,
          text: note.display,
          number: note.display,
          ...(notePage != null ? { page: notePage } : {}),
          listText: body || note.display,
          contentFrom: at,
          contentTo: at + 1,
          ...(pair ? { bookmarkFrom: pair.from, bookmarkTo: pair.to } : {}),
        });
      }

      // The paragraph's remaining bookmark pairs are user targets.
      pairs.forEach((pair, i) => {
        if (used.has(pair.name)) return;
        out.push({
          key: `bookmark:${pair.name}`,
          kind: "bookmark",
          name: pair.name,
          pos,
          text: texts[i]!.trim(),
          ...(page != null ? { page } : {}),
          listText: texts[i]!.trim() || pair.name,
          contentFrom,
          contentTo,
          bookmarkFrom: pair.from,
          bookmarkTo: pair.to,
        });
      });

      return true;
    });
    return out;
  }

  /** Caption dialog 确定 — seed a Caption-styled paragraph beside the caret's
   *  paragraph carrying the next SEQ field (cached, so the projection paints
   *  the number without field evaluation) wrapped in a _Ref bookmark pair
   *  (the cross-reference target), in one transaction. The field code carries
   *  the picked number format (`\*`) and, with "Include chapter number", the
   *  heading level (`\s`); the separator lands in the document's caption
   *  settings (settings.xml `w:captions`), where the SEQ evaluator reads it —
   *  Word's split between field code and document setting.
   *
   *  Word interop deviation: Word writes the chapter as a separate
   *  `STYLEREF <level> \s` field and caches only the sequence number in the
   *  SEQ result; docen folds the chapter prefix into the SEQ cache (the
   *  projection paints a single field atom — see the SEQ evaluator). An
   *  in-Word update therefore drops the prefix, and Update All cannot refresh
   *  a real Word STYLEREF.
   *
   *  The Caption style definition joins the document styles when absent
   *  (compile passes doc.attrs.styles straight through). */
  readonly onCaptionOk = (event: Event): void => {
    const { label, text, position, excludeLabel, chapterNumber, heading, sep, format } =
      (
        event as CustomEvent<{
          label?: string;
          text?: string;
          position?: string;
          excludeLabel?: boolean;
          chapterNumber?: boolean;
          heading?: number;
          sep?: string;
          format?: string;
        }>
      ).detail ?? {};
    const target = this.#target();
    if (!target || !label) return;
    const { state } = target;
    const $from = state.selection.$from;
    if ($from.parent.type.name !== "paragraph") return;
    const insertPos = position === "above" ? $from.before($from.depth) : $from.after($from.depth);
    const level =
      chapterNumber &&
      typeof heading === "number" &&
      Number.isInteger(heading) &&
      heading >= 1 &&
      heading <= 9
        ? heading
        : undefined;
    const numberFormat = format && SEQ_NUMBER_FORMATS[format] ? format : "ARABIC";
    const separatorToken = sep && CAPTION_SEPARATOR_CHARS[sep] ? sep : "hyphen";
    const instruction = `SEQ ${label} \\* ${numberFormat}${level != null ? ` \\s ${level}` : ""}`;
    // The sequence state at the insertion point — reset-aware, so a caption
    // inserted mid-document carries the ordinal it will keep. The separator
    // rides the context directly: the settings write below lands after the
    // cached value is evaluated.
    const walk = this.#seqWalk(target, insertPos);
    const seq = (walk.counts.get(label) ?? 0) + 1;
    const chapter =
      level != null && walk.chapterCounts[level]! > 0
        ? String(walk.chapterCounts[level])
        : undefined;
    const cachedValue =
      evaluateField(instruction, {
        ...this.#fieldBase(target),
        captionSeparators: new Map([[label, CAPTION_SEPARATOR_CHARS[separatorToken]!]]),
        sequences: new Map([[label, seq]]),
        ...(chapter != null ? { chapters: new Map([[level!, chapter]]) } : {}),
      }) ?? String(seq);
    const bookmarkId = this.nextBookmarkId(target);
    const name = `_Ref${String(bookmarkId).padStart(8, "0")}`;
    const seed = (data: object): JSONContent =>
      ({
        type: "inlinePassthrough",
        attrs: { data: JSON.stringify(data) },
      }) as JSONContent;
    const caption: JSONContent = {
      type: "paragraph",
      attrs: { style: "Caption" },
      // The _Ref bookmark wraps only label + number (Word's shape), so a
      // cross-reference's "label and number" content reads "图 1" without the
      // caption text.
      content: [
        seed({ bookmarkStart: { id: bookmarkId, name } }),
        ...(excludeLabel ? [] : [{ type: "text", text: `${label} ` }]),
        seed({ simpleField: { instruction, cachedValue } }),
        seed({ bookmarkEnd: { id: bookmarkId } }),
        ...(text ? [{ type: "text", text: `: ${text}` }] : []),
      ],
    };
    const styles = { ...((state.doc.attrs.styles ?? {}) as Record<string, unknown>) };
    const paragraphStyles = (styles.paragraphStyles ?? []) as { id?: string }[];
    const nextExtras = this.#patchCaptionSettings(target, {
      name: label,
      chapterNumber: level != null,
      ...(level != null ? { heading: level } : {}),
      sep: separatorToken,
      numFmt: SEQ_NUMBER_FORMATS[numberFormat] ?? "decimal",
    });
    target
      .chain()
      .command(({ tr }) => {
        if (!paragraphStyles.some((s) => s.id === "Caption")) {
          tr.step(
            new DocAttrStep("styles", {
              ...styles,
              paragraphStyles: [
                ...paragraphStyles,
                {
                  id: "Caption",
                  name: "caption",
                  basedOn: "Normal",
                  next: "Normal",
                  uiPriority: 35,
                  quickFormat: true,
                },
              ],
            }),
          );
        }
        tr.step(new DocAttrStep("documentExtras", nextExtras));
        // Doc attrs sit outside the position space, so the doc-attr steps and
        // the insertion below compose in either order.
        tr.insert(insertPos, target.schema.nodeFromJSON(caption));
        return true;
      })
      .run();
    this.host.bridge()?.focus();
  };

  /** Cross-reference dialog 确定 — seed a cached REF/PAGEREF/NOTEREF field at
   *  the caret for the picked target × content, Word's field shapes:
   *
   *  - text / label / entire caption / caption text → `REF <name> \h`
   *  - heading or numbered-item number → `REF <name> \n \h`
   *  - footnote/endnote number → `NOTEREF <name> \h`
   *  - page number → `PAGEREF <name> \h`
   *  - above/below → `REF|NOTEREF <name> \p \h`
   *
   *  A target without a bookmark gains Word's hidden `_Ref…` pair around the
   *  range the chosen content needs (the whole heading/paragraph, a caption's
   *  label+number, its caption text, or the whole caption) in the same
   *  transaction; the cache is evaluated from the live document, so the
   *  insert and a later F9 / Update All Fields agree. */
  readonly onCrossRefOk = (event: Event): void => {
    const { key, content } =
      (event as CustomEvent<{ key?: string; content?: CrossRefContent }>).detail ?? {};
    const editor = this.#target();
    if (!editor || !key || !content) return;
    const target = this.crossReferenceTargets().find((entry) => entry.key === key);
    if (!target || !CROSS_REFERENCE_CONTENTS[target.kind].includes(content)) return;
    // Options the dialog disables are not inserted as empty references.
    if (!crossReferenceContentAvailable(target, content)) return;
    const requested = this.#crossReferenceRange(target, content);
    if (!requested) return;
    // Reuse the target's bookmark when it covers the requested range — for
    // "label and number" the caption's own `_Ref` pair; for text/number/page/
    // above-below the target's own bookmark (a user bookmark's range is the
    // reference). Word creates a hidden `_Ref…` pair for the rest.
    const reuse =
      content === "label"
        ? (target.labelName ??
          (target.bookmarkFrom === requested.from && target.bookmarkTo === requested.to
            ? target.name
            : ""))
        : content === "entire" || content === "captionText"
          ? target.bookmarkFrom === requested.from && target.bookmarkTo === requested.to
            ? target.name
            : ""
          : (target.name ?? "");
    const range = reuse ? null : requested;
    const bookmarkId = range ? this.nextBookmarkId(editor) : null;
    const name = range ? `_Ref${String(bookmarkId).padStart(8, "0")}` : reuse;
    const instruction = crossReferenceInstruction(target.kind, content, name);
    if (!instruction) return;

    const { from } = editor.state.selection;
    // Above/below resolves from the reading-order relation (before the
    // bookmark and field inserts shift either position).
    const position = target.pos < from ? "above" : "below";
    const terms = this.#positionTerms();
    const text = this.#crossReferenceText(target, content);
    const extra = new Map<string, FieldBookmark>();
    if (range) extra.set(name, this.#fieldBookmark(editor, target.pos, text, target.number));
    const value =
      evaluateField(instruction, {
        ...this.#fieldBase(editor, this.#bookmarks(editor, extra)),
        frame: this.#frameAt(editor, from),
        selfPos: from,
        positionTerms: terms,
      }) ?? this.#crossReferenceFallback(target, content, position, terms);

    const seed = (data: object): JSONContent =>
      ({
        type: "inlinePassthrough",
        attrs: { data: JSON.stringify(data) },
      }) as JSONContent;
    const tr = editor.state.tr;
    if (range) {
      tr.insert(
        tr.mapping.map(range.from),
        editor.schema.nodeFromJSON(seed({ bookmarkStart: { id: bookmarkId, name } })),
      );
      tr.insert(
        tr.mapping.map(range.to),
        editor.schema.nodeFromJSON(seed({ bookmarkEnd: { id: bookmarkId } })),
      );
    }
    tr.insert(
      tr.mapping.map(from),
      editor.schema.nodeFromJSON(seed({ simpleField: { instruction, cachedValue: value } })),
    );
    editor.view.dispatch(tr);
    this.host.bridge()?.focus();
  };

  /** The hidden-bookmark range a content choice needs: the caption's
   *  label+number or caption text when it is narrower than the target's own
   *  content range (whole caption, heading/numbered paragraph, note atom). */
  #crossReferenceRange(
    target: CrossReferenceTarget,
    content: CrossRefContent,
  ): { from: number; to: number } | null {
    if (content === "label") {
      return target.labelTo != null ? { from: target.contentFrom, to: target.labelTo } : null;
    }
    if (content === "captionText") {
      return target.captionTextFrom != null
        ? { from: target.captionTextFrom, to: target.contentTo }
        : null;
    }
    return { from: target.contentFrom, to: target.contentTo };
  }

  /** The inner text a freshly created hidden bookmark carries (what REF/
   *  NOTEREF re-derives); undefined for page/above-below content, which
   *  ignores the text. */
  #crossReferenceText(target: CrossReferenceTarget, content: CrossRefContent): string | undefined {
    if (content === "entire") return target.captionFull ?? target.text;
    if (content === "captionText") return target.captionText;
    if (content === "page" || content === "aboveBelow") return undefined;
    return target.text;
  }

  /** The cached value when no frame/evaluator can supply one — the target's
   *  own text (Word caches what it just inserted), the page fallback 1, or
   *  the above/below word. */
  #crossReferenceFallback(
    target: CrossReferenceTarget,
    content: CrossRefContent,
    position: "above" | "below",
    terms: { above: string; below: string },
  ): string {
    const text = this.#crossReferenceText(target, content);
    if (text != null) return text;
    if (content === "number") return target.number ?? "";
    if (content === "page") return String(target.page ?? 1);
    return position === "above" ? terms.above : terms.below;
  }

  /** One bookmark-table entry for REF/PAGEREF/NOTEREF — the target's page
   *  frame and displayed page with the supplied inner text/number. */
  #fieldBookmark(editor: Editor, pos: number, text?: string, number?: string): FieldBookmark {
    const frame = this.#frameAt(editor, pos);
    const physical = this.host.bridge()?.pageOf(pos);
    const page = frame?.page ?? (typeof physical === "number" ? physical + 1 : undefined);
    return {
      ...(text ? { text } : {}),
      ...(number ? { number } : {}),
      pos,
      ...(page != null ? { page } : {}),
      ...(frame?.pageFormat ? { pageFormat: frame.pageFormat } : {}),
    };
  }

  /** The `\p` switch's words in the host UI's language (Word renders them in
   *  the document language). English when the host supplies none (headless). */
  #positionTerms(): { above: string; below: string } {
    return this.host.positionTerms?.() ?? { above: "above", below: "below" };
  }

  // ── Footnote/endnote bodies ──
  // Note content lives in doc attrs documentExtras.footnotes/endnotes —
  // outside the PM doc — so the body has no caret of its own: the dialog is
  // the editing surface, the reference atom in the text is the anchor, and
  // the NotesCleanup extension prunes bodies whose last reference is deleted.

  /** The reference atom under the caret — the caret sits right after a
   *  right-click collapse, so the atom starting at the caret and the one
   *  ending just before it both count. */
  noteTarget(): { kind: "footnote" | "endnote"; id: number; pos: number } | null {
    const editor = this.host.editor();
    if (!editor) return null;
    const { from } = editor.state.selection;
    for (const at of [from, from - 1]) {
      if (at < 0) continue;
      const node = editor.state.doc.nodeAt(at);
      if (!node || node.type.name !== "inlinePassthrough") continue;
      const data = node.attrs.data;
      if (typeof data !== "string") continue;
      try {
        const ref = noteRefId(JSON.parse(data) as Record<string, unknown>);
        if (ref?.id == null) continue;
        return {
          kind: ref.kind === "endnoteReference" ? "endnote" : "footnote",
          id: ref.id,
          pos: at,
        };
      } catch {
        /* malformed atom JSON — not a note reference */
      }
    }
    return null;
  }

  /** The note's plain text — one line per note paragraph (the textarea's
   *  model); only text branches speak, the note-number marker stays silent. */
  #noteTextOf(children: unknown): string {
    const paragraphText = (p: unknown): string => {
      if (typeof p === "string") return p;
      if (!p || typeof p !== "object") return "";
      const runs = (p as { children?: unknown }).children;
      if (!Array.isArray(runs)) return "";
      let text = "";
      for (const run of runs) {
        if (typeof run === "string") text += run;
        else if (
          run &&
          typeof run === "object" &&
          typeof (run as { text?: unknown }).text === "string"
        )
          text += (run as { text: string }).text;
      }
      return text;
    };
    if (!Array.isArray(children)) return "";
    const lines: string[] = [];
    for (const child of children) {
      if (typeof child === "string") {
        lines.push(child);
      } else if (child && typeof child === "object" && "paragraph" in child) {
        lines.push(paragraphText((child as { paragraph: unknown }).paragraph));
      } else {
        lines.push(paragraphText(child));
      }
    }
    return lines.join("\n");
  }

  /** Rebuild the note body from textarea lines — each line becomes a
   *  paragraph in the original first paragraph's style, the leading
   *  note-number marker run kept on the first (Word's canonical note body;
   *  a missing marker is re-injected by the notes stringify on export). */
  #noteChildrenOf(text: string, previous: unknown): unknown[] {
    const first = Array.isArray(previous) ? previous[0] : undefined;
    const firstParagraph =
      first && typeof first === "object" && "paragraph" in first
        ? (first as { paragraph: unknown }).paragraph
        : first;
    const style =
      firstParagraph && typeof firstParagraph === "object"
        ? (firstParagraph as { style?: unknown }).style
        : undefined;
    const firstRuns = (
      firstParagraph && typeof firstParagraph === "object"
        ? (firstParagraph as { children?: unknown }).children
        : undefined
    ) as unknown[] | undefined;
    const marker = (firstRuns ?? []).find(
      (run) =>
        run &&
        typeof run === "object" &&
        ((run as Record<string, unknown>).footnoteRef === true ||
          (run as Record<string, unknown>).endnoteRef === true),
    );
    return text.split("\n").map((line, i) => ({
      ...(style != null ? { style } : {}),
      children: [...(i === 0 && marker ? [marker] : []), ...(line ? [{ text: line }] : [])],
    }));
  }

  /** Insert → Footnote/Endnote: open the dialog empty; the commit (onNoteOk's
   *  insert branch) drops the reference atom at the caret and appends the
   *  note body to documentExtras. */
  noteInsert(kind: "footnote" | "endnote"): void {
    this.#noteTarget = null;
    this.#noteDialog()?.show(kind, "");
  }

  /** Context menu → 编辑脚注/尾注: prefill the dialog with the referenced
   *  note's text; the commit rewrites that note's body in place. */
  noteEditAtSelection(): void {
    const target = this.noteTarget();
    if (!target) return;
    const note = this.#noteOf(target.kind, target.id);
    this.#noteTarget = target;
    this.#noteDialog()?.show(target.kind, note ? this.#noteTextOf(note.children) : "");
  }

  /** Context menu → 删除脚注/尾注: delete the reference atom (Word deletes a
   *  note by deleting its reference); the cleanup extension prunes the body
   *  when its last reference is gone. */
  noteDeleteAtSelection(): void {
    const editor = this.host.editor();
    const target = this.noteTarget();
    if (!editor || !target) return;
    editor.view.dispatch(editor.state.tr.delete(target.pos, target.pos + 1));
  }

  readonly onNoteOk = (event: Event): void => {
    const { kind, text } =
      (event as CustomEvent<{ kind?: "footnote" | "endnote"; text?: string }>).detail ?? {};
    if ((kind !== "footnote" && kind !== "endnote") || text == null) return;
    const editing = this.#noteTarget;
    this.#noteTarget = null;
    const editor = this.host.editor();
    if (!editor) return;
    if (editing) {
      this.#rewriteNote(kind, editing.id, text);
      return;
    }
    if (!text) return;
    this.#insertNoteBody(kind, text);
  };

  /** The pending edit target between noteEditAtSelection and the dialog's OK
   *  (a modal dialog — one pending target at a time; null = insert mode). */
  #noteTarget: { kind: "footnote" | "endnote"; id: number; pos: number } | null | undefined;

  #noteDialog(): { show(kind: "footnote" | "endnote", text: string): void } | null | undefined {
    return this.host.element().shadowRoot?.querySelector("docen-note-dialog") as
      | { show(kind: "footnote" | "endnote", text: string): void }
      | null
      | undefined;
  }

  #extras(): NoteExtras {
    const editor = this.host.editor();
    return (editor?.state.doc.attrs.documentExtras ?? {}) as NoteExtras;
  }

  #noteOf(kind: "footnote" | "endnote", id: number): NoteEntry | undefined {
    const extras = this.#extras();
    return (kind === "endnote" ? extras.endnotes : extras.footnotes)?.find(
      (note) => note.id === id,
    );
  }

  /** The edit commit: swap the referenced note's body for the rebuilt lines
   *  (one attrs transaction — the projection repaints the page-bottom note). */
  #rewriteNote(kind: "footnote" | "endnote", id: number, text: string): void {
    const editor = this.host.editor();
    if (!editor) return;
    const channel = kind === "endnote" ? "endnotes" : "footnotes";
    const extras = this.#extras();
    const notes = extras[channel] ?? [];
    let swapped = false;
    const next = notes.map((note) => {
      if (note.id !== id) return note;
      swapped = true;
      return { ...note, children: this.#noteChildrenOf(text, note.children) };
    });
    if (!swapped) return;
    editor.view.dispatch(
      editor.state.tr.setDocAttribute("documentExtras", { ...extras, [channel]: next }),
    );
  }

  /** The insert commit (the old #insertNote, dialog-fed): a reference atom at
   *  the caret plus the note body — Word's canonical *Text paragraph leading
   *  with the note-number marker run — and the content-types override the
   *  packer needs for the first note of a kind. */
  #insertNoteBody(kind: "footnote" | "endnote", text: string): void {
    const editor = this.host.editor();
    if (!editor) return;
    const channel = `${kind}s` as "footnotes" | "endnotes";
    const extras = this.#extras();
    const notes = extras[channel] ?? [];
    const id = notes.reduce((max, note) => Math.max(max, Number(note.id ?? 0)), 0) + 1;
    const Note = kind === "footnote" ? "FootnoteText" : "EndnoteText";
    const ref = editor.schema.nodeFromJSON({
      type: "inlinePassthrough",
      attrs: { data: JSON.stringify({ [`${kind}Reference`]: id }) },
    } as JSONContent);
    const documentExtras: NoteExtras = {
      ...extras,
      [channel]: [
        ...notes,
        {
          id,
          // One *Text paragraph per textarea line, the note-number marker run
          // leading the first (the seed supplies both — the builder lifts them).
          children: this.#noteChildrenOf(text, [
            { style: Note, children: [{ [`${kind}Ref`]: true }] },
          ]),
        },
      ],
    };
    const partName = `/word/${channel}.xml`;
    if (
      extras.contentTypes &&
      !extras.contentTypes.overrides?.some((o) => o.partName === partName)
    ) {
      documentExtras.contentTypes = {
        ...extras.contentTypes,
        overrides: [
          ...(extras.contentTypes.overrides ?? []),
          {
            partName,
            contentType: `application/vnd.openxmlformats-officedocument.wordprocessingml.${channel}+xml`,
          },
        ],
      };
    }
    editor.view.dispatch(
      editor.state.tr
        .insert(editor.state.selection.from, ref)
        .setDocAttribute("documentExtras", documentExtras),
    );
  }

  // ── Fields (域) ──
  // Fields ride the inlinePassthrough atom (attrs.data = the office-open
  // simpleField/complexField/formField branch); the painter already renders
  // them, so the editing surface here is insert (Field dialog), edit (the
  // same dialog prefilled), update (re-derive the cache, Word's F9), and the
  // form-field checkbox flip.

  /** The field atom under the caret — the caret sits right after a
   *  right-click collapse, so the atom starting at the caret and the one
   *  ending just before it both count. */
  fieldTarget(): { pos: number; branch: Record<string, unknown>; ref: FieldRef } | null {
    const editor = this.#target();
    if (!editor) return null;
    const { from } = editor.state.selection;
    for (const at of [from, from - 1]) {
      if (at < 0) continue;
      const node = editor.state.doc.nodeAt(at);
      if (!node || node.type.name !== "inlinePassthrough") continue;
      const data = node.attrs.data;
      if (typeof data !== "string") continue;
      try {
        const branch = JSON.parse(data) as Record<string, unknown>;
        const ref = fieldRef(branch);
        if (ref) return { pos: at, branch, ref };
      } catch {
        /* malformed atom JSON — not a field */
      }
    }
    return null;
  }

  /** Insert → Field: open the dialog empty. */
  fieldInsert(): void {
    this.#fieldTarget = null;
    this.#fieldDialog()?.show();
  }

  /** Context menu → 编辑域: prefill the dialog with the field's instruction;
   *  the commit rewrites it in place and re-derives the cached value. */
  fieldEditAtSelection(): void {
    const hit = this.fieldTarget();
    if (!hit?.ref.instruction) return;
    this.#fieldTarget = hit.pos;
    this.#fieldDialog()?.show(hit.ref.instruction);
  }

  /** Context menu → 更新域 (Word's F9): re-derive the cached value from the
   *  live document. Page-dependent fields resolve against the host's pinned
   *  pagination (`fieldFrame`); fields the context can't evaluate keep their
   *  cache. */
  fieldUpdateAtSelection(): void {
    const editor = this.#target();
    const hit = this.fieldTarget();
    if (!editor || !hit) return;
    const value = evaluateField(
      hit.ref.instruction ?? "",
      this.#fieldContext(editor, hit.pos, hit.ref.instruction),
    );
    if (value == null) return;
    const patch = hit.ref.kind === "simpleField" ? { cachedValue: value } : { result: value };
    const next = this.#patchField(hit.branch, hit.ref.kind, patch);
    if (!next) return;
    editor.view.dispatch(editor.state.tr.setNodeAttribute(hit.pos, "data", JSON.stringify(next)));
  }

  /** Update All Fields (Word's select-all + F9): re-derive every field atom's
   *  cached value in document order and commit the changes in one transaction.
   *  SEQ occurrences take their own ordinal per label (the caption sequence),
   *  restarting at each heading their `\s` switch names; `\*` formats the
   *  number and `\s` prefixes the chapter number, so a deleted or moved
   *  caption renumbers from the document itself. REF/PAGEREF/NOTEREF resolve
   *  through the bookmark table, and page-dependent fields use the host's
   *  pinned pagination when it exists. Returns the number of fields whose
   *  cache changed.
   *
   *  Two phases in one transaction: every other field (captions' SEQ
   *  included) lands first, then the bookmark table is re-scanned from that
   *  transaction doc, then the REF family resolves — a reference to a caption
   *  whose SEQ renumbered in this pass quotes the new number immediately
   *  instead of needing a second update. Positions are stable across the
   *  phases (attribute patches never change node sizes).
   *
   *  Limitation: a complex field whose result is structured (`resultRunsXml`
   *  — TOC entries and nested fields) is patched on its flat `result` only;
   *  paint keeps the structured runs, so such fields must be refreshed through
   *  their own commands (the TOC updaters). Rewriting result XML here would
   *  mean regenerating runs, not patching a cache. */
  updateAllFields(): number {
    const editor = this.#target();
    if (!editor) return 0;
    const base = this.#fieldBase(editor);
    const seq = this.#seqWalk(editor);
    const tr = editor.state.tr;
    const changed = new Set<number>();

    const contextOf = (
      pos: number,
      field: ReturnType<typeof parseFieldInstruction>,
      bookmarks: ReadonlyMap<string, FieldBookmark>,
    ): FieldContext => {
      const context: FieldContext = {
        ...base,
        bookmarks,
        frame: this.#frameAt(editor, pos),
        selfPos: pos,
      };
      if (field.name === "SEQ") {
        const label = field.args[0];
        const ordinal = label ? seq.ordinals.get(pos) : undefined;
        if (label && ordinal != null) context.sequences = new Map([[label, ordinal]]);
        const level = seqChapterLevel(field.switches.s);
        const chapter = level != null ? seq.chapters.get(pos) : undefined;
        if (level != null && chapter != null) context.chapters = new Map([[level, chapter]]);
      }
      return context;
    };

    const apply = (
      doc: PMNode,
      bookmarks: ReadonlyMap<string, FieldBookmark>,
      refFamily: boolean,
    ): void => {
      doc.descendants((node, pos) => {
        if (node.type.name !== "inlinePassthrough") return true;
        const data = node.attrs.data;
        if (typeof data !== "string") return true;
        let branch: Record<string, unknown>;
        try {
          branch = JSON.parse(data) as Record<string, unknown>;
        } catch {
          return true;
        }
        const ref = fieldRef(branch);
        if (!ref?.instruction || ref.kind === "formField") return true;
        const field = parseFieldInstruction(ref.instruction);
        if (BOOKMARK_FIELD_NAMES.has(field.name) !== refFamily) return true;
        const value = evaluateField(ref.instruction, contextOf(pos, field, bookmarks));
        if (value == null || value === ref.result) return true;
        const patch = ref.kind === "simpleField" ? { cachedValue: value } : { result: value };
        const next = this.#patchField(branch, ref.kind, patch);
        if (!next) return true;
        tr.setNodeAttribute(pos, "data", JSON.stringify(next));
        changed.add(pos);
        return true;
      });
    };

    // Phase 1: everything but the bookmark-reading REF family.
    apply(editor.state.doc, base.bookmarks ?? new Map(), false);
    // Phase 2: the REF family against the phase-1 doc's bookmarks.
    apply(tr.doc, this.#bookmarks(editor, undefined, tr.doc), true);

    if (changed.size === 0) return 0;
    editor.view.dispatch(tr);
    return changed.size;
  }

  /** Context menu → checkbox toggle: flip the form field's checked flag (the
   *  painter re-renders the box from it). */
  fieldToggleCheckboxAtSelection(): void {
    const editor = this.#target();
    const hit = this.fieldTarget();
    if (!editor || hit?.ref.kind !== "formField") return;
    const form = hit.branch.formField as { checkBox?: Record<string, unknown> } | undefined;
    if (!form?.checkBox) return;
    const next = {
      ...hit.branch,
      formField: { ...form, checkBox: { ...form.checkBox, checked: !hit.ref.checked } },
    };
    editor.view.dispatch(editor.state.tr.setNodeAttribute(hit.pos, "data", JSON.stringify(next)));
  }

  readonly onFieldOk = (event: Event): void => {
    const { instruction } = (event as CustomEvent<{ instruction?: string }>).detail ?? {};
    const pos = this.#fieldTarget;
    this.#fieldTarget = null;
    const editor = this.#target();
    if (!editor || !instruction) return;
    if (pos != null) {
      // Edit: rewrite the atom's instruction, carrying the rest of the branch
      // through; the cache re-derives, keeping the old value for instructions
      // the engine can't evaluate.
      const node = editor.state.doc.nodeAt(pos);
      const data = node?.attrs.data;
      if (!node || node.type.name !== "inlinePassthrough" || typeof data !== "string") return;
      try {
        const branch = JSON.parse(data) as Record<string, unknown>;
        const ref = fieldRef(branch);
        if (!ref || ref.kind === "formField") return;
        const value =
          evaluateField(instruction, this.#fieldContext(editor, pos, instruction)) ??
          ref.result ??
          "";
        const patch =
          ref.kind === "simpleField"
            ? { instruction, cachedValue: value }
            : { instruction, result: value };
        const next = this.#patchField(branch, ref.kind, patch);
        if (next)
          editor.view.dispatch(editor.state.tr.setNodeAttribute(pos, "data", JSON.stringify(next)));
      } catch {
        /* malformed atom JSON — nothing to rewrite */
      }
      return;
    }
    // Insert: a simple field at the caret. PAGE/NUMPAGES cache nothing (the
    // painter resolves them per page); an unevaluable instruction caches an
    // empty result, honestly blank until Word/our updater fills it.
    const seed = editor.schema.nodeFromJSON({
      type: "inlinePassthrough",
      attrs: {
        data: JSON.stringify({
          simpleField: {
            instruction,
            cachedValue:
              evaluateField(
                instruction,
                this.#fieldContext(editor, editor.state.selection.from, instruction),
              ) ?? "",
          },
        }),
      },
    } as JSONContent);
    editor.view.dispatch(editor.state.tr.insert(editor.state.selection.from, seed));
    this.host.bridge()?.focus();
  };

  /** The pending edit target between fieldEditAtSelection and the dialog's OK
   *  (a modal dialog — one pending atom at a time; null = insert mode). */
  #fieldTarget: number | null | undefined;

  #fieldDialog(): { show(prefill?: string): void } | null | undefined {
    return this.host.element().shadowRoot?.querySelector("docen-field-dialog") as
      | { show(prefill?: string): void }
      | null
      | undefined;
  }

  /** The evaluation context slices every field update shares: the doc's core
   *  properties (AUTHOR/TITLE/SUBJECT/KEYWORDS/COMMENTS/CREATEDATE/…), the
   *  filename (FILENAME), the core revision (REVNUM), custom properties
   *  (DOCPROPERTY), the bookmark table (REF/PAGEREF/NOTEREF, with `extras`
   *  merged for bookmarks a commit is about to create), the `\p` words, the
   *  caption separators (SEQ chapter numbering) and the live word/char totals
   *  (NUMWORDS/NUMCHARS). The "now" clock is per call — an update command is a
   *  moment, not a render. */
  #fieldBase(
    editor: Editor,
    bookmarks: ReadonlyMap<string, FieldBookmark> = this.#bookmarks(editor),
  ): Omit<FieldContext, "frame" | "sequences" | "chapters"> {
    const doc = editor.state.doc;
    const attrs = (doc.attrs ?? {}) as {
      core?: Record<string, unknown>;
      documentExtras?: Record<string, unknown>;
    };
    const text = doc.textBetween(0, doc.content.size, "\n", "");
    const revision = finiteNumber(attrs.core?.revision);
    const filename = this.host.filename?.();
    const customProperties = customPropertiesOf(attrs.documentExtras);
    const captionSeparators = this.#captionSeparators(editor);
    return {
      now: new Date(),
      core: attrs.core ?? {},
      words: wordCounter(text),
      chars: textCounter(text),
      ...(filename != null ? { filename } : {}),
      ...(revision != null ? { revision } : {}),
      ...(customProperties ? { customProperties } : {}),
      ...(captionSeparators ? { captionSeparators } : {}),
      positionTerms: this.#positionTerms(),
      bookmarks,
    };
  }

  /** REF/PAGEREF/NOTEREF targets: each bookmark's inner text/number plus the
   *  DISPLAYED page its start sits on (the host's page frame — restart and
   *  numFmt applied) with that page's section format. Falls back to the
   *  bridge's physical page only when no frame is available (headless);
   *  absent = the bookmark keeps its cached value. `extras` carries bookmarks
   *  created by the transaction a commit is building; `doc` scans a document
   *  snapshot other than the live one (Update All's phase-1 doc). */
  #bookmarks(
    editor: Editor,
    extras?: ReadonlyMap<string, FieldBookmark>,
    doc?: PMNode,
  ): ReadonlyMap<string, FieldBookmark> {
    const map = new Map<string, FieldBookmark>();
    for (const target of this.#crossReferenceTargetsIn(editor, doc ?? editor.state.doc)) {
      if (!target.name) continue;
      map.set(target.name, this.#fieldBookmark(editor, target.pos, target.text, target.number));
    }
    if (extras) for (const [name, bookmark] of extras) map.set(name, bookmark);
    return map;
  }

  /** The rendered page frame at a document position — the host's pagination
   *  view (page number/count, section number/span, numFmt). Null headless, and
   *  null inside a header/footer story: the host's page map is the body's, and
   *  a story position maps to no page (page fields there keep their cache and
   *  paint through the furniture's live page context instead). */
  #frameAt(editor: Editor, pos?: number): FieldFrame | undefined {
    if (editor !== this.host.editor()) return undefined;
    return this.host.fieldFrame?.(pos ?? editor.state.selection.from);
  }

  /** Field evaluation context from the live document — the body text backs
   *  NUMWORDS/NUMCHARS (the status bar's counters; notes stay out, Word's
   *  default), the core properties the document attrs carry, and the host's
   *  pinned pagination for the page-dependent fields. With an instruction,
   *  a SEQ field also gets its live ordinal and `\s` chapter number at `pos`
   *  (the position-bound walk applies the chapter's pending reset, so F9 and
   *  edit agree with Update All Fields). */
  #fieldContext(editor: Editor, pos?: number, instruction?: string): FieldContext {
    const selfPos = pos ?? editor.state.selection.from;
    const context: FieldContext = {
      ...this.#fieldBase(editor),
      frame: this.#frameAt(editor, selfPos),
      selfPos,
    };
    if (!instruction) return context;
    const field = parseFieldInstruction(instruction);
    const label = field.name === "SEQ" ? field.args[0] : undefined;
    if (!label) return context;
    const walk = this.#seqWalk(editor, selfPos);
    context.sequences = new Map([[label, (walk.counts.get(label) ?? 0) + 1]]);
    const level = seqChapterLevel(field.switches.s);
    if (level != null && walk.chapterCounts[level]! > 0)
      context.chapters = new Map([[level, String(walk.chapterCounts[level])]]);
    return context;
  }

  /** The branch with the field's flat shape patched (simpleField's
   *  cachedValue / complexField's instruction+result); null for form fields,
   *  which the dialog doesn't edit. */
  #patchField(
    branch: Record<string, unknown>,
    kind: FieldRef["kind"],
    patch: Record<string, unknown>,
  ): Record<string, unknown> | null {
    if (kind === "formField") return null;
    const field = branch[kind];
    if (!field || typeof field !== "object") return null;
    return { ...branch, [kind]: { ...(field as Record<string, unknown>), ...patch } };
  }
}
