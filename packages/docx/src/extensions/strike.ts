import type { RunOptions } from "@office-open/docx";

import { Strike as BaseStrike } from "./tiptap";
import { attrNative } from "./utils";

/**
 * Strike mark extension with nested office-open attrs.
 *
 * OOXML represents strikethrough on a run via two mutually exclusive booleans:
 * `strike` (single) and `doubleStrike` (double). The mark itself is "single
 * strikethrough"; the `doubleStrike` attr flips it to double. DOCX round-trip
 * is near-identity for the doubleStrike flag; `strike` itself is implied by the
 * mark's presence and handled in renderDocx/parseDocx.
 *
 * Mark attribute-level renderHTML is delegated to the base Strike extension
 * (renders `<s>`); only the DOCX flag needs custom handling.
 */

// ── DOCX serialization (exported for DocxManager) ──

/**
 * attrs → run properties.
 *
 * The mark presence itself means single strikethrough. When `doubleStrike` is
 * set, emit the OOXML double-strike flag instead (the two are mutually
 * exclusive in OOXML). DocxManager calls this with `mark.attrs`.
 */
export function renderDocx(attrs: Record<string, unknown>): Partial<RunOptions> {
  return attrs.doubleStrike ? { doubleStrike: true } : { strike: true };
}

/**
 * run properties → attrs.
 *
 * Only the `doubleStrike` flag needs to round-trip; single strike is implied by
 * the mark's existence. `strike` is structural/semantic (the mark) and skipped.
 */
export function parseDocx(runOpts: RunOptions): Record<string, unknown> {
  return { doubleStrike: runOpts.doubleStrike ?? null };
}

// ── Extension ──

export const Strike = BaseStrike.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      // OOXML double-strike flag (stored verbatim; no dedicated CSS equivalent)
      doubleStrike: attrNative(),
    };
  },

  // Mark attribute-level renderHTML is handled by the base Strike extension (<s>).

  renderDocx,
  parseDocx,
});
