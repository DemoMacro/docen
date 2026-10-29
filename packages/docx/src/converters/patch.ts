import type { PatchRouteStep } from "@office-open/core";
import { parseDocument, patchDocument } from "@office-open/docx";
import type {
  DocumentPatchCommand,
  DocumentOptions,
  OutputByType,
  OutputType,
  SectionChild,
} from "@office-open/docx";

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
 * A high-level placeholder facade over office-open's ordered
 * `DocumentPatchCommand` API.
 */
export interface DocxPatchOptions<T extends OutputType = OutputType> {
  /** Template DOCX to patch. */
  template: Parameters<typeof parseDocument>[0];
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
  const { template, patches, prepare = true, placeholderDelimiters, outputType } = options;

  const delimiters = {
    start: placeholderDelimiters?.start ?? "{{",
    end: placeholderDelimiters?.end ?? "}}",
  };
  const document = await parseDocument(template);
  const replacements = new Map<string, SectionChild[]>();

  for (const [key, patchContent] of Object.entries(patches)) {
    if (prepare !== false) {
      await prepareDocument(patchContent.content, prepare === true ? undefined : prepare);
    }
    const docOpts = compileDocument(patchContent.content);
    replacements.set(key, docOpts.sections?.[0]?.children ?? []);
  }

  const commands = buildPlaceholderCommands(document.sections ?? [], replacements, delimiters);

  return patchDocument({
    outputType,
    data: template,
    commands,
    placeholderDelimiters: delimiters,
  });
}

function buildPlaceholderCommands(
  sections: NonNullable<DocumentOptions["sections"]>,
  replacements: ReadonlyMap<string, SectionChild[]>,
  delimiters: { start: string; end: string },
): DocumentPatchCommand[] {
  const commands: DocumentPatchCommand[] = [];
  const targets = [...sections.keys()].reverse().flatMap((index) => {
    const child = sections[index];
    if (!child) return [];

    const section = sections[index]!;
    const children =
      "children" in section && Array.isArray(section.children) ? section.children : [];
    return [...children.keys()].reverse().flatMap((childIndex): PlaceholderTarget[] => {
      const target = children[childIndex];
      if (!target) return [];

      const baseRoute: PatchRouteStep[] = [
        { index },
        { property: "children" },
        { index: childIndex },
      ];
      return placeholderTargets(target, baseRoute, delimiters, replacements);
    });
  });

  for (const target of targets) {
    for (const [key, value] of replacements) {
      if (target.placeholder !== key) continue;

      const simpleText = value.length === 1 ? simpleTextOf(value[0]!) : undefined;
      if (simpleText !== undefined) {
        const command: Extract<DocumentPatchCommand, { type: "replace"; value: string }> = {
          type: "replace",
          at: { field: "sections", route: target.route, text: { placeholder: key, all: true } },
          value: simpleText,
        };
        commands.push(command);
        continue;
      }
      if (target.isOnlyPlaceholder) {
        commands.push({ type: "remove", at: { field: "sections", route: target.route } });
      } else {
        commands.push({
          type: "replace",
          at: { field: "sections", route: target.route, text: { placeholder: key, all: true } },
          value: "",
        });
      }
      const parentRoute = target.route.slice(0, -1);
      const positionStep = target.route.at(-1);
      if (!positionStep || !("index" in positionStep)) continue;
      commands.push({
        type: "insert",
        at: { field: "sections", route: parentRoute },
        position: target.isOnlyPlaceholder ? positionStep.index : positionStep.index + 1,
        value,
      });
    }
  }

  return commands;
}

interface PlaceholderTarget {
  route: PatchRouteStep[];
  placeholder: string;
  isOnlyPlaceholder: boolean;
}

function placeholderTargets(
  value: unknown,
  route: PatchRouteStep[],
  delimiters: { start: string; end: string },
  replacements: ReadonlyMap<string, SectionChild[]>,
): PlaceholderTarget[] {
  if (typeof value !== "object" || value === null) return [];

  const paragraph = "paragraph" in value ? value.paragraph : undefined;
  if (paragraph !== undefined) {
    return [...replacements.keys()].flatMap((key): PlaceholderTarget[] => {
      const placeholder = `${delimiters.start}${key}${delimiters.end}`;
      if (!containsText(paragraph, placeholder)) return [];
      return [
        { route, placeholder: key, isOnlyPlaceholder: textOf(paragraph).trim() === placeholder },
      ];
    });
  }

  if (Array.isArray(value)) {
    return [...value.keys()].reverse().flatMap((index): PlaceholderTarget[] => {
      const child = value[index];
      if (!child) return [];
      return placeholderTargets(child, [...route, { index }], delimiters, replacements);
    });
  }

  return Object.entries(value)
    .reverse()
    .flatMap(([property, child]): PlaceholderTarget[] => {
      if (typeof child !== "object" || child === null) return [];
      return placeholderTargets(child, [...route, { property }], delimiters, replacements);
    });
}

function simpleTextOf(value: SectionChild): string | undefined {
  if (!("paragraph" in value)) return undefined;
  const text = textOf(value.paragraph);
  return text.length > 0 ? text : undefined;
}

function containsText(value: unknown, needle: string): boolean {
  return textOf(value).includes(needle);
}

function textOf(value: unknown): string {
  if (typeof value === "string") return value;
  if (Array.isArray(value)) return value.map(textOf).join("");
  if (typeof value !== "object" || value === null) return "";
  return Object.values(value).map(textOf).join("");
}
