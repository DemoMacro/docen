// The PPTX engine for docen editors: re-exports the OOXML parse/generate
// surface from @office-open/pptx and adds the projection into the drawing
// members the core painter paints.

export {
  generatePresentation,
  parsePresentation,
  parsePresentationSync,
  type PresentationOptions,
  type SlideOptions,
  type SlideChild,
  type SlideSize,
} from "@office-open/pptx";
export { projectPresentation, type ProjectedPresentation, type ProjectedSlide } from "./scene";
