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

  const editor = new Editor({
    element,
    extensions: [...docxExtensions, ...extraExtensions],
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
