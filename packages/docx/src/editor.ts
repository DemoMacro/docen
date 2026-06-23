import { Editor, type AnyExtension } from "./core";
import { docxExtensions } from "./core";

export interface DocxEditorOptions {
  /** DOM element to mount the editor */
  element: HTMLElement;
  /** Additional Tiptap extensions */
  extensions?: AnyExtension[];
  /** Initial content (Tiptap JSON or HTML string) */
  content?: Record<string, unknown> | string;
  /** Enable spellcheck (disable for large documents) */
  spellcheck?: boolean;
  /** Editor is editable */
  editable?: boolean;
}

/**
 * Create a Tiptap editor configured for DOCX editing.
 */
export function createDocxEditor(options: DocxEditorOptions): Editor {
  const {
    element,
    extensions: extraExtensions = [],
    content,
    spellcheck = true,
    editable = true,
  } = options;

  // Dedupe: an extra extension overrides any docxExtension of the same name.
  // The editor layer passes PageDocument (extends Document) to switch the doc
  // schema to `doc > page+`; without this filter both register name "doc" and
  // Tiptap warns "Duplicate extension names found: ['doc']".
  const extraNames = new Set(
    extraExtensions.map((e) => (e as { name?: string }).name).filter(Boolean),
  );
  const extensions = [
    ...docxExtensions.filter((e) => !extraNames.has((e as { name?: string }).name)),
    ...extraExtensions,
  ];

  const editor = new Editor({
    element,
    extensions,
    content: content ?? { type: "doc", content: [{ type: "paragraph" }] },
    editable,
    editorProps: {
      attributes: {
        spellcheck: spellcheck ? "true" : "false",
      },
    },
  });

  return editor;
}
