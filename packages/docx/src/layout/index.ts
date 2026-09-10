export { projectDocumentOptions, projectFlowBox, type ProjectedSection } from "./project";
export type { MarkupDisplay } from "./project/context";
// Page-level projection types live in @docen/layout (the painter consumes
// them without depending on this package).
export type {
  ProjectedFlowBox,
  ProjectedPageBackground,
  ProjectedPageFurniture,
} from "@docen/layout";
