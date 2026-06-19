import { patchDocument } from "@office-open/docx";
import type { OutputByType, OutputType, SectionChild } from "@office-open/docx";

import type { JSONContent } from "../core";
import { compileDocument } from "./docx";
import { prepareDocument, type PrepareStep } from "./prepare";

/**
 * A patch's replacement content, expressed as Tiptap JSON. `compileDocument`
 * converts it to office-open `SectionChild[]` before applying.
 */
export interface DocxPatchContent {
  content: JSONContent;
}

/**
 * Options for patching a DOCX template.
 *
 * Mirrors the legacy `@docen/export-docx` DocxPatchOptions minus `exportOptions`
 * — `compileDocument` already derives all styling from the Tiptap attrs, so no
 * separate export configuration is needed.
 */
export interface DocxPatchOptions<T extends OutputType = OutputType> {
  /** Template DOCX to patch. */
  template: Parameters<typeof patchDocument>[0]["data"];
  /** Placeholder name → replacement content. */
  patches: Record<string, DocxPatchContent>;
  /**
   * Pre-compilation steps run on each patch's content in place (default:
   * `prepareImages()`). `false` skips; `PrepareStep[]` runs custom steps.
   * Required when patch content references http image URLs.
   */
  prepare?: boolean | PrepareStep[];
  /** Custom placeholder delimiters (default `{{` / `}}`). */
  placeholderDelimiters?: { start?: string; end?: string };
  /** Keep the template's paragraph styles on patched content. */
  keepOriginalStyles?: boolean;
  /** Recurse into nested patches. */
  recursive?: boolean;
  /** Output container type. */
  outputType: T;
}

/**
 * Patch a DOCX template by replacing placeholders with Tiptap-JSON content.
 *
 * Each patch's `content` is prepared (default: `prepareImages`, in place) then
 * compiled (`compileDocument` → `DocumentOptions`); its first section's
 * `children` become the replacement. Patching is delegated to
 * `@office-open/docx`'s `patchDocument`.
 */
export async function patchDOCX<T extends OutputType>(
  options: DocxPatchOptions<T>,
): Promise<OutputByType[T]> {
  const {
    template,
    patches,
    prepare = true,
    placeholderDelimiters,
    keepOriginalStyles,
    recursive,
    outputType,
  } = options;

  const patchesObject: Record<string, { type: "document"; children: SectionChild[] }> = {};

  for (const [key, patchContent] of Object.entries(patches)) {
    if (prepare !== false) {
      await prepareDocument(patchContent.content, prepare === true ? undefined : prepare);
    }
    const docOpts = compileDocument(patchContent.content);
    const children = (docOpts.sections?.[0]?.children ?? []) as SectionChild[];
    patchesObject[key] = { type: "document", children };
  }

  return patchDocument({
    outputType,
    data: template,
    placeholders: patchesObject,
    ...(keepOriginalStyles !== undefined && { keepOriginalStyles }),
    ...(recursive !== undefined && { recursive }),
    ...(placeholderDelimiters && {
      placeholderDelimiters: {
        start: placeholderDelimiters.start ?? "{{",
        end: placeholderDelimiters.end ?? "}}",
      },
    }),
  });
}
