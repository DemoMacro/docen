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
  type SlideCommentOptions,
  type SlideSize,
  type TableOptions,
  type TableCellOptions,
  type TransitionOptions,
  type TransitionType,
  type AnimationDuration,
  type SlideAnimation,
} from "@office-open/pptx";
export type {
  CompressionOptions,
  OutputByType,
  OutputType,
  PackerOptions,
} from "@office-open/pptx";
export type { DataType, SyncDataType } from "@office-open/core";
export type { ColorSchemeOptions, FontSchemeOptions } from "@office-open/core";
export type { ChartOptions, ShapeOptions } from "@office-open/pptx";
export type {
  BulletOptions,
  CustomGeometryOptions,
  ParagraphDescriptorOptions,
  ShapeType,
  StrikeStyle,
  TextAlignment,
  TextBodyOptions,
  TextBreakOptions,
  TextRunOptions,
  UnderlineStyle,
  TextVertical,
} from "@office-open/core/drawing";
export type { CategoryPointOptions, ChartValue } from "@office-open/core";
export {
  projectPresentation,
  type ProjectedPresentation,
  type ProjectedSlideBackground,
  type ProjectedSlide,
  type ProjectedSlideMember,
} from "./scene";
export { textBlocks, type TextFieldContext } from "./scene";
export { resizeTableGrid, tableGridOf, type CellOrigin } from "./scene/tables";
export {
  memberAt,
  memberByPath,
  offsetMemberByPath,
  renderableChild,
  resizeMemberByPath,
  type MemberHit,
} from "./scene/walk";
