/**
 * @docen/core — the scene painter for the docen editors: it instantiates a
 * laid-out `LayoutDoc` (@docen/layout) as a LeaferJS tree. Consumed by the
 * canvas editors in `@docen/editor`.
 *
 * The painter owns the `LayoutDoc → Leafer` mapping and **no editing
 * semantics** (selection, undo, keyboard/IME, handles, property panes) and
 * **no layout decisions** (the engine in @docen/layout owns all geometry).
 *
 * @module
 */

export * from "./painter";
