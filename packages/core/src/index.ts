/**
 * @docen/core — shared rendering layer for the docen editors: geometry, style,
 * and image helpers plus the scene painter that instantiates a laid-out
 * `LayoutDoc` (@docen/layout) as a LeaferJS tree. Consumed by the canvas
 * editors in `@docen/editor`.
 *
 * This package owns the `LayoutDoc → Leafer` mapping plus geometry math and
 * export. It deliberately owns **no editing semantics** (selection, undo,
 * keyboard/IME, handles, property panes) and **no layout decisions** (the
 * engine in @docen/layout owns all geometry) — those live in `@docen/layout`
 * and `@docen/editor` respectively.
 *
 * @module
 */

export * from "./geometry";
export * from "./style";
export * from "./image";
export * from "./export";
export * from "./painter";
