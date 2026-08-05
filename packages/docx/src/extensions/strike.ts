import type { RunOptions } from "@office-open/docx";
import { Strike as BaseStrike } from "@tiptap/extension-strike";

/**
 * Strike mark — editor interaction only (toolbar toggle, `<s>`).
 *
 * OOXML has two mutually exclusive strikethrough booleans: `strike` (single)
 * and `doubleStrike` (double). Both are three-state and ride TextStyle's native
 * attrs for round-trip + CSS cascade, so this mark only surfaces the
 * "single strike = true" case for editing. doubleStrike=true (no single-strike
 * mark) round-trips purely through TextStyle.
 */
export const Strike = BaseStrike.extend({
  renderDocx: () => ({ strike: true }),
  parseDocx: (opts: RunOptions): Record<string, unknown> | null => (opts.strike ? {} : null),
});
