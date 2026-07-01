import type { LevelsOptions } from "@office-open/docx";
import { LevelFormat } from "@office-open/docx";

import { OrderedList as OrderedListBase } from "./tiptap";

/**
 * OrderedList extension — owns the DOCX expression of an ordered list.
 *
 * A Tiptap orderedList maps to a sequence of paragraphs referencing one
 * abstractNum (decimal). The reference is keyed by `start` so lists with the
 * same start share a definition; DocxManager gives each list its own instance
 * for independent counting. This module owns that abstractNum shape; DocxManager
 * owns the cross-paragraph tree walk, start recovery, and numbering-instance
 * bookkeeping.
 */

/** Reference prefix for generated ordered-list abstractNum definitions. */
export const ORDERED_REFERENCE_PREFIX = "docen-ordered";

/** lvlText per nesting depth (level 0 → "%1.", … level 8 → "%9."). */
export const ORDERED_LEVEL_TEXT = ["%1.", "%2.", "%3.", "%4.", "%5.", "%6.", "%7.", "%8.", "%9."];

/**
 * Build nine decimal numbering levels for an ordered list. Level 0 carries
 * `start`; deeper levels restart at 1 (Word convention).
 */
export function buildOrderedLevels(start: number): LevelsOptions[] {
  return Array.from(
    { length: 9 },
    (_, level): LevelsOptions => ({
      level,
      format: LevelFormat.DECIMAL,
      start: level === 0 ? start : 1,
      text: ORDERED_LEVEL_TEXT[level],
    }),
  );
}

// DocxManager builds the abstractNum via buildOrderedLevels and tracks each
// list's start/instance itself. The extension also carries `numbering` —
// the source abstractNum reference (when the list came from parseDOCX) — so the
// round-trip reuses the original numbering definition instead of regenerating
// the default; DOCX-only (not rendered to HTML), null for editor-created lists.
export const OrderedList = OrderedListBase.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      numbering: { default: null, rendered: false },
    };
  },
});
