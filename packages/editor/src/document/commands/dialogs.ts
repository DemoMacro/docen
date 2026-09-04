import type { JSONContent } from "@docen/docx";
import { buildCustomMultilevelLevels, nextMultilevelReference } from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import { TextSelection } from "@tiptap/pm/state";
import { DocAttrStep } from "@tiptap/pm/transform";

import type { FontDialogPatch } from "../../ui/components/workspace/font-dialog";
import type {
  DrawingPropertiesPatch,
  ParagraphDialogPatch,
  TablePropertiesPatch,
} from "../extensions/commands";
import { collectListReferences } from "../extensions/commands";

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
}

/**
 * The Home/References dialog commits, split out of the host element: the
 * paragraph/font/table-properties patches, the proofing language, the Chinese
 * layout pair (phonetic guide, two lines in one), the multilevel list
 * definition, and the caption/cross-reference/bookmark field seeds — see the
 * matching *-dialog.ts components for the UI side.
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

  /** Language dialog 确定 — commit the selection's proofing language (w:lang). */
  readonly onLanguageOk = (event: Event): void => {
    const { value, noProof } = (event as CustomEvent<{ value?: string; noProof?: boolean }>)
      .detail ?? { value: undefined, noProof: false };
    const target = this.#target();
    if (!value || !target) return;
    if (target.state.selection.empty) return;
    target
      .chain()
      .setMark("textStyle", { language: { value }, noProof: noProof ? true : null })
      .run();
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

  /** The next SEQ number for a caption label — one past the SEQ fields with
   *  the same label already in the document (each label's sequence restarts
   *  at 1; the fields carry the cached results the projection paints). */
  #nextSeqNumber(target: Editor, label: string): number {
    let count = 0;
    target.state.doc.descendants((node) => {
      if (node.type.name !== "inlinePassthrough") return true;
      try {
        const data = JSON.parse(String(node.attrs.data ?? "{}")) as {
          simpleField?: { instruction?: string };
        };
        const tokens = (data.simpleField?.instruction ?? "").trim().split(/\s+/);
        if (tokens[0]?.toUpperCase() === "SEQ" && tokens[1] === label) count++;
      } catch {
        // opaque verbatim blobs without field data — skip
      }
      return true;
    });
    return count + 1;
  }

  /** The document's referenceable bookmarks for the cross-reference dialog —
   *  a caption's `_Ref` pair (its inner text reads "图 1") plus the user's
   *  own bookmarks, each with its bookmark position for PAGEREF page lookups. */
  crossReferenceTargets(): { name: string; text: string; kind: string; pos: number }[] {
    const target = this.#target();
    if (!target) return [];
    const out: { name: string; text: string; kind: string; pos: number }[] = [];
    target.state.doc.descendants((node, pos) => {
      if (node.type.name !== "paragraph") return true;
      let name = "";
      let open = false;
      let text = "";
      node.forEach((child) => {
        if (child.type.name === "inlinePassthrough") {
          try {
            const data = JSON.parse(String(child.attrs.data ?? "{}")) as {
              bookmarkStart?: { name?: string };
              bookmarkEnd?: unknown;
              simpleField?: { cachedValue?: string };
            };
            if (data.bookmarkStart?.name && !open) {
              name = data.bookmarkStart.name;
              open = true;
              return;
            }
            if (data.bookmarkEnd && open) {
              open = false;
              return;
            }
            if (open && data.simpleField) text += String(data.simpleField.cachedValue ?? "");
          } catch {
            // opaque verbatim blobs — skip
          }
        } else if (open && child.type.name === "text") {
          text += child.textContent ?? "";
        }
      });
      if (name)
        out.push({
          name,
          text: text.trim(),
          kind: name.startsWith("_Ref") ? "caption" : "bookmark",
          pos,
        });
      return true;
    });
    return out;
  }

  /** Caption dialog 确定 — seed a Caption-styled paragraph beside the caret's
   *  paragraph carrying the next SEQ field (cached, so the projection paints
   *  the number without field evaluation) wrapped in a _Ref bookmark pair
   *  (the cross-reference target), in one transaction. The Caption style
   *  definition joins the document styles when absent (compile passes
   *  doc.attrs.styles straight through). */
  readonly onCaptionOk = (event: Event): void => {
    const { label, text, position, excludeLabel } =
      (
        event as CustomEvent<{
          label?: string;
          text?: string;
          position?: string;
          excludeLabel?: boolean;
        }>
      ).detail ?? {};
    const target = this.#target();
    if (!target || !label) return;
    const { state } = target;
    const $from = state.selection.$from;
    if ($from.parent.type.name !== "paragraph") return;
    const seq = this.#nextSeqNumber(target, label);
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
        seed({
          simpleField: { instruction: `SEQ ${label} \\* ARABIC`, cachedValue: String(seq) },
        }),
        seed({ bookmarkEnd: { id: bookmarkId } }),
        ...(text ? [{ type: "text", text: `: ${text}` }] : []),
      ],
    };
    const styles = { ...((state.doc.attrs.styles ?? {}) as Record<string, unknown>) };
    const paragraphStyles = (styles.paragraphStyles ?? []) as { id?: string }[];
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
        // Doc attrs sit outside the position space, so the doc-attr step and
        // the insertion below compose in either order.
        tr.insert(
          position === "above" ? $from.before($from.depth) : $from.after($from.depth),
          target.schema.nodeFromJSON(caption),
        );
        return true;
      })
      .run();
    this.host.bridge()?.focus();
  };

  /** Cross-reference dialog 确定 — seed a cached REF field at the caret: a
   *  REF for "label and number" / "bookmark text" content (the bookmark's
   *  inner text), a PAGEREF for page content (resolved through the bridge's
   *  pageOf, the same geometry the TOC's page numbers use). Both carry \h —
   *  Word's hyperlink form. */
  readonly onCrossRefOk = (event: Event): void => {
    const { name, content } =
      (event as CustomEvent<{ name?: string; content?: string }>).detail ?? {};
    const target = this.#target();
    if (!target || !name) return;
    const hit = this.crossReferenceTargets().find((entry) => entry.name === name);
    if (!hit) return;
    // The bridge's pageOf is 0-based; Word's page numbers are 1-based (the
    // TOC's conversion, with 1 as the fallback when the position has no
    // laid-out page yet).
    const page = this.host.bridge()?.pageOf(hit.pos);
    const cached =
      content === "page" ? String(typeof page === "number" ? page + 1 : 1) : hit.text || "1";
    const seed: JSONContent = {
      type: "inlinePassthrough",
      attrs: {
        data: JSON.stringify({
          simpleField: {
            instruction: `${content === "page" ? "PAGEREF" : "REF"} ${name} \\h`,
            cachedValue: cached,
          },
        }),
      },
    } as JSONContent;
    const { from } = target.state.selection;
    target.view.dispatch(target.state.tr.insert(from, target.schema.nodeFromJSON(seed)));
    this.host.bridge()?.focus();
  };
}
