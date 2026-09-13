/**
 * Format-neutral geometry — pure computation, no DOM, no Leafer. Exposed as a
 * subpath export (`@docen/core/geometry`) so Node-side consumers (layout
 * projections, specs) can evaluate preset shapes without loading the
 * browser-only painter barrel.
 *
 * @module
 */

export * from "./drawing-props";
export * from "./preset-shape";
