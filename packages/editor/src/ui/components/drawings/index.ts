/**
 * Drawing web components — first-class interactive components for Office
 * DrawingML objects (image, shape, chart, smartart).
 *
 * Each component is a standalone web component (not a Tiptap NodeView): it
 * speaks OOXML data in (JSON `attrs` attribute) and edited geometry out
 * (`change` CustomEvent). The docx `ImageView` / `WpsShapeView` NodeViews are
 * thin adapters that mount these components and forward `change` to a
 * ProseMirror transaction; a property pane, dialog, or the future pptx slide
 * canvas can mount the same components independently.
 *
 * @module
 */

export { default as DocenImage } from "./image";
