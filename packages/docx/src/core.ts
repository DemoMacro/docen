/**
 * DOCX extension core — re-exports @tiptap/core with DOCX augmentations.
 *
 * This module serves the same role as @tiptap/core but extended with
 * renderDocx/parseDocx support (declared via module augmentation in types.d.ts).
 *
 * Consumers should import from "@docen/docx" instead of "@tiptap/core"
 * to get the full DOCX-aware type system.
 */

// Re-export tiptap/core fundamentals (used throughout the package)
export type { JSONContent, AnyExtension, Extensions } from "@tiptap/core";
export { Editor, Extension, Node, Mark } from "@tiptap/core";

// DOCX extension registry
export { docxExtensions } from "./extensions";
