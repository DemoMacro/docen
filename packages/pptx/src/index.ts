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
  type TableOptions,
  type TableCellOptions,
  type TransitionOptions,
  type TransitionType,
  type SlideAnimation,
} from "@office-open/pptx";
export {
  projectPresentation,
  type ProjectedPresentation,
  type ProjectedSlideBackground,
  type ProjectedSlide,
} from "./scene";
export { textBlocks, type TextFieldContext } from "./scene";
export { tableGridOf, type CellOrigin } from "./scene/tables";
export {
  memberAt,
  memberByPath,
  offsetMemberByPath,
  resizeMemberByPath,
  type MemberHit,
} from "./scene/walk";
