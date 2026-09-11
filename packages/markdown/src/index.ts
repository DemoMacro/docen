import { parseMarkdownTokens } from "./parse";
import { renderMarkdownBlocks } from "./render";
import type {
  MarkdownMapper,
  MdAlignment,
  MdBlock,
  MdInline,
  MdListItem,
  MdMark,
  MdTableCell,
  MdTableRow,
} from "./types";

/**
 * Parse Markdown into the target document model. The mapper carries the
 * format knowledge — `@docen/docx` binds this with its own mapper and
 * re-exports the familiar one-argument `parseMarkdown`.
 */
export function parseMarkdown<T>(markdown: string, mapper: MarkdownMapper<T>): T[] {
  return parseMarkdownTokens(markdown).flatMap((block) => mapper.parseBlock(block));
}

/**
 * Render the target document model back to Markdown. The mapper projects the
 * document onto the neutral IR first; DOCX-only nodes degrade to their text.
 */
export function generateMarkdown<T>(doc: T, mapper: MarkdownMapper<T>): string {
  return renderMarkdownBlocks(mapper.toBlocks(doc));
}

export type {
  MarkdownMapper,
  MdAlignment,
  MdBlock,
  MdInline,
  MdListItem,
  MdMark,
  MdTableCell,
  MdTableRow,
};
export { parseMarkdownTokens, renderMarkdownBlocks };
