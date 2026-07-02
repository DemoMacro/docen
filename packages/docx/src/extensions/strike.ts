import type { RunOptions } from "@office-open/docx";
import { Strike as BaseStrike } from "@tiptap/extension-strike";

import { attrNative } from "./utils";

/**
 * Strike mark extension with nested office-open attrs.
 *
 * OOXML represents strikethrough on a run via two mutually exclusive booleans:
 * `strike` (single) and `doubleStrike` (double). The mark itself is "single
 * strikethrough"; the `doubleStrike` attr flips it to double. DOCX round-trip
 * is near-identity for the doubleStrike flag; `strike` itself is implied by the
 * mark's presence.
 *
 * Mark attribute-level renderHTML is delegated to the base Strike extension
 * (renders `<s>`); only the DOCX flag needs custom handling.
 */
export const Strike = BaseStrike.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      // OOXML double-strike flag (stored verbatim; no dedicated CSS equivalent)
      doubleStrike: attrNative(),
    };
  },

  // Single strike is implied by the mark; doubleStrike flips to OOXML
  // double-strike (the two are mutually exclusive), so only that flag round-trips.
  renderDocx: (attrs: Record<string, unknown>): Partial<RunOptions> =>
    attrs.doubleStrike ? { doubleStrike: true } : { strike: true },
  parseDocx: (runOpts: RunOptions): Record<string, unknown> | null =>
    runOpts.strike ? { doubleStrike: runOpts.doubleStrike ?? null } : null,
});
