// DOCX pipeline re-exported from @docen/docx.
//
// parseDOCX / generateDOCX are high-level JSONContent ↔ DOCX binary converters
// (mirroring parseHTML/generateHTML). generateDOCX defaults to running
// prepareDocument (http image embedding) and carries outputType as its generic T
// via `packer.type`; generateDOCXSync / generateDOCXStream cover sync + streaming.
export {
  generateDOCX,
  generateDOCXStream,
  generateDOCXSync,
  parseDOCX,
  patchDOCX,
} from "@docen/docx";
export type { DocxGenerateOptions, DocxPatchContent, DocxPatchOptions } from "@docen/docx";
