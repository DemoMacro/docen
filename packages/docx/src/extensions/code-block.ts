import { CodeBlockLowlight } from "@tiptap/extension-code-block-lowlight";

import type { JSONContent } from "../core";

/**
 * CodeBlock extension — CodeBlockLowlight with a DOCX renderDocx.
 *
 * DOCX has no dedicated code-block element. A code block maps to a single
 * paragraph styled "Code" with a monospace run font. Inline marks (syntax-
 * highlight tokens) and line breaks (`\n` → `<w:br/>`) are handled by
 * DocxManager's shared inline-content compilation; resolveCodeBlock reassembles
 * the text. This module owns only the paragraph-level style.
 *
 * `language` has no OOXML carrier and is intentionally dropped (known lossy).
 */

/** codeBlock node → paragraph properties (style "Code" + monospace font). */
export function renderDocx(_node: JSONContent): Record<string, unknown> {
  return { style: "Code", run: { font: "Consolas" } };
}

// DocxManager calls renderDocx for the paragraph style; resolveCodeBlock handles
// parsing (break → "\n"). No parseDocx — codeBlock has no DOCX attrs to recover.
export const CodeBlock = CodeBlockLowlight.extend({ renderDocx });
