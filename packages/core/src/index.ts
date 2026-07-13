/**
 * @docen/core — pure rendering layer mapping OOXML drawing data to LeaferJS /
 * ECharts elements. Shared across the docx (Tiptap), pptx (LeaferJS), and xlsx
 * (RevoGrid) editors in `@docen/editor`.
 *
 * This package owns the `render*` (data → element) and `parse*` (element →
 * data) mappings plus geometry math and export. It deliberately owns **no
 * editing semantics** (selection, undo, keyboard/IME, handles, property panes)
 * — those live in `@docen/editor`'s web components and NodeViews.
 *
 * @module
 */

export * from "./geometry";
export * from "./style";
export * from "./image";
export * from "./export";
