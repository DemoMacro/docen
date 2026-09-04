import type { JSONContent } from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import { DocAttrStep } from "@tiptap/pm/transform";

import { t } from "../../ui";

/** A bibliography source as the sources dialog edits it — office-open's
 *  SourceTypeOptions narrowed to the exposed fields (the document's own
 *  sources round-trip untouched through the open attrs value). */
export type BibliographySource = {
  tag?: string;
  sourceType?: string;
  title?: string;
  year?: string;
  publisher?: string;
  author?: { authors?: { last?: string; first?: string; corporate?: string }[] };
};

/** The references commands' view of the host — resolved per call so the
 *  controller can be built before a document opens (the editor and the story
 *  bridge both arrive later). */
export interface ReferencesHost {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The story bridge — citations target the active story (header/footer
   *  stories included). */
  bridge(): { activeEditor(): Editor; focus(): void } | undefined;
  /** The host element — the i18n language source for prompts. */
  element(): HTMLElement;
}

/**
 * The References tab's citation/bibliography commands, split out of the host
 * element: the source master list (doc.attrs.bibliography), the CITATION
 * field seed, and the Bibliography-styled block rebuild.
 */
export class ReferencesCommands {
  constructor(private readonly host: ReferencesHost) {}

  #target(): Editor | null | undefined {
    return this.host.bridge()?.activeEditor() ?? this.host.editor();
  }

  /** Mark Entry — prompt for the entry text (defaulting to the selection) and
   *  seed an `XE "…"` field at the selection's end. A cached-less fldSimple
   *  renders nothing (Word's invisible index marker) but round-trips verbatim
   *  through DOCX for Insert Index to collect. */
  markIndexEntry(target: Editor): void {
    const { empty, from, to } = target.state.selection;
    const selected = empty ? "" : target.state.doc.textBetween(from, to, " ");
    const entry = window.prompt(t("index.prompt", this.host.element()), selected)?.trim();
    if (!entry) return;
    const seed: JSONContent = {
      type: "inlinePassthrough",
      attrs: {
        data: JSON.stringify({
          simpleField: { instruction: `XE "${entry.replaceAll('"', "''")}"` },
        }),
      },
    } as JSONContent;
    target.view.dispatch(target.state.tr.insert(to, target.schema.nodeFromJSON(seed)));
    this.host.bridge()?.focus();
  }

  /** The document's bibliography sources — doc.attrs.bibliography (the Source
   *  Manager's master list, word/bibliography.xml on save). */
  bibliographySources(): BibliographySource[] {
    const attrs = this.host.editor()?.state.doc.attrs as {
      bibliography?: { sources?: BibliographySource[] } | null;
    };
    return [...(attrs.bibliography?.sources ?? [])];
  }

  /** Sources dialog commit — replace the document's source list. An empty
   *  list clears the attr (null) so no empty part is emitted. */
  readonly onSourcesOk = (event: Event): void => {
    const { sources } = (event as CustomEvent<{ sources?: BibliographySource[] }>).detail ?? {};
    const target = this.#target();
    if (!target || !sources) return;
    target.commands.command(({ tr }) => {
      tr.step(new DocAttrStep("bibliography", sources.length > 0 ? { sources } : null));
      return true;
    });
  };

  /** Citation dialog insert — seed a cached CITATION field at the caret in
   *  Word's in-text shape "(Author, Year)" (the title stands in when the
   *  source has no author). */
  readonly onCitationOk = (event: Event): void => {
    const { tag } = (event as CustomEvent<{ tag?: string }>).detail ?? {};
    const target = this.#target();
    if (!target || !tag) return;
    const source = this.bibliographySources().find((entry) => entry.tag === tag);
    const authors = (source?.author?.authors ?? [])
      .map((person) => [person.last, person.first].filter(Boolean).join(", "))
      .join("; ");
    const head = authors || source?.title || tag;
    const cached = `(${head}${source?.year ? `, ${source.year}` : ""})`;
    const seed: JSONContent = {
      type: "inlinePassthrough",
      attrs: {
        data: JSON.stringify({
          simpleField: { instruction: `CITATION "${tag}" \\l 1033`, cachedValue: cached },
        }),
      },
    } as JSONContent;
    const { from } = target.state.selection;
    target.view.dispatch(target.state.tr.insert(from, target.schema.nodeFromJSON(seed)));
    this.host.bridge()?.focus();
  };

  /** Bibliography — rebuild the Bibliography-styled block after the caret's
   *  paragraph from the document's sources (an existing block is replaced,
   *  not duplicated). The two style definitions join the document styles when
   *  absent. */
  insertBibliography(): void {
    const target = this.#target();
    if (!target) return;
    const sources = this.bibliographySources();
    if (sources.length === 0) {
      window.alert(t("bibliography.empty", this.host.element()));
      return;
    }
    const { state } = target;
    if (state.selection.$from.parent.type.name !== "paragraph") return;
    // A simplified APA entry line: Authors (Year). Title. Publisher.
    const entryText = (source: BibliographySource): string => {
      const authors = (source.author?.authors ?? [])
        .map((person) => person.corporate ?? [person.last, person.first].filter(Boolean).join(", "))
        .join("; ");
      return [
        authors,
        source.year ? `(${source.year})` : "",
        source.title ? `${source.title}.` : "",
        source.publisher ? `${source.publisher}.` : "",
      ]
        .filter(Boolean)
        .join(" ");
    };
    const block: JSONContent[] = [
      {
        type: "paragraph",
        attrs: { style: "BibliographyHeading" },
        content: [{ type: "text", text: t("bibliography.heading", this.host.element()) }],
      },
      ...sources.map((source) => ({
        type: "paragraph",
        attrs: { style: "Bibliography" },
        content: [{ type: "text", text: entryText(source) }],
      })),
    ];
    const styles = { ...((state.doc.attrs.styles ?? {}) as Record<string, unknown>) };
    const paragraphStyles = (styles.paragraphStyles ?? []) as { id?: string }[];
    const missing = (["BibliographyHeading", "Bibliography"] as const).filter(
      (id) => !paragraphStyles.some((style) => style.id === id),
    );
    const afterCaret = state.selection.$from.after(state.selection.$from.depth);
    target
      .chain()
      .command(({ tr }) => {
        if (missing.length > 0) {
          const definitions = missing.map((id) => ({
            id,
            name: id === "Bibliography" ? "bibliography" : "Bibliography Heading",
            basedOn: "Normal",
            next: "Normal",
            ...(id === "BibliographyHeading" ? { bold: true } : {}),
          }));
          tr.step(
            new DocAttrStep("styles", {
              ...styles,
              paragraphStyles: [...paragraphStyles, ...definitions],
            }),
          );
        }
        // Drop the stale block first; map the insertion anchor through the
        // deletions so the fresh block lands after the caret's paragraph
        // even when the old block sat before the caret.
        state.doc.descendants((node, at) => {
          if (node.type.name !== "paragraph") return true;
          const raw = (node.attrs as Record<string, unknown>).style;
          const style = typeof raw === "string" ? raw : "";
          if (style === "Bibliography" || style === "BibliographyHeading")
            tr.delete(at, at + node.nodeSize);
          return true;
        });
        tr.insert(
          tr.mapping.map(afterCaret),
          block.map((node) => target.schema.nodeFromJSON(node)),
        );
        return true;
      })
      .run();
    this.host.bridge()?.focus();
  }
}
