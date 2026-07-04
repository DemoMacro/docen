import type { DocenAddin, DocenHost } from "./types";

/**
 * Declare an add-in with full typing on the host generic.
 *
 * Pure identity at runtime — the value is in the type: spell out the host
 * subtype (`defineAddin<DocumentHost>(…)`) so each `commands` handler and
 * `taskPanes.render` receives a properly typed host, with the addin shape
 * checked at the call site rather than only where it's consumed.
 */
export function defineAddin<THost extends DocenHost = DocenHost>(
  addin: DocenAddin<THost>,
): DocenAddin<THost> {
  return addin;
}
