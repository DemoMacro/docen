/**
 * The drawing editor domain — the format-independent layer behind selecting
 * and editing drawings (pictures, shapes, groups) on the canvas: hit/box
 * geometry, the selection frame and crop overlays, the gesture state
 * machine, and the docx hit→node resolution. Shared by the document editor
 * today; the presentation/workbook editors adopt the same layer.
 */
export * from "./target";
export * from "./gestures";
export * from "./docx";
export * from "./geometry";
export * from "./overlay";
export * from "./crop-overlay";
