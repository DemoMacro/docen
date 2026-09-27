// The scene projection — PresentationOptions → the drawing members the core
// painter paints. Public surface: projectPresentation and the projected
// shapes; the batch map lives in presentation.ts.

export {
  projectPresentation,
  type ProjectedPresentation,
  type ProjectedSlideBackground,
  type ProjectedSlide,
  type ProjectedSlideMember,
} from "./presentation";
export { textBlocks, type TextFieldContext } from "./text";
