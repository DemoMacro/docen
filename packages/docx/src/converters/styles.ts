import type { BorderOptions, ParagraphOptions, StylesOptions } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";

import type { ResolveContext } from "../extensions/types";
import { resolveFontName } from "../extensions/utils";
import {
  defaultParagraphStyleId,
  indexParagraphStyles,
  mergeStyleChain,
  pStyleIdFromKey,
  type StyleEntry,
} from "../style-cascade";

// Re-export the public styles model type so consumers (the editor's Styles
// gallery) type against office-open's source of truth instead of a local
// mirror. The style-entry type and every cascade primitive live in
// style-cascade.ts (rendering-neutral); this module owns the style-facing
// editor helpers (Quick Styles gallery, caret run props, resolve/compile
// attr helpers).
export type { StylesOptions };

// ── Quick Styles gallery selection ──────────────────────────────────────────

/** A gallery-ready paragraph-style entry for the Styles combobox. */
export interface QuickStyleEntry {
  id: string;
  name: string;
}

/** The `DefaultStylesOptions` keys whose values are character styles, not
 *  paragraph styles. The Quick Styles gallery is paragraph-only, so these are
 *  excluded even when flagged `quickFormat` (the authoritative source is
 *  office-open's DefaultStylesOptions interface). */
const CHARACTER_DEFAULT_KEYS = new Set([
  "hyperlink",
  "footnoteReference",
  "footnoteTextChar",
  "endnoteReference",
  "endnoteTextChar",
]);

/**
 * The paragraph styles to list in the Quick Styles gallery, matching Word's
 * default behavior: the gallery is a *paragraph-style* selector (it applies a
 * pStyle), so only paragraph styles appear — never character styles, even
 * those flagged `quickFormat` (those live in the Styles task pane). Among
 * paragraph styles, only those flagged `quickFormat` are listed, ordered by
 * `uiPriority` (Word orders the gallery this way).
 *
 * Reads `quickFormat`/`uiPriority`/`name` straight from office-open's styles
 * model (`StylesOptions`): `paragraphStyles` (Normal + custom) and the built-in
 * named paragraph styles nested under `default` (title/heading1-9/quote/…). The
 * `default` keys that hold character styles are excluded via
 * `CHARACTER_DEFAULT_KEYS`. When a document carries no quickFormat flags at all
 * (e.g. some LibreOffice-generated files), fall back to all paragraph styles so
 * the gallery is never empty.
 */
export function quickStyles(styles: StylesOptions | null | undefined): QuickStyleEntry[] {
  if (!styles) return [];
  type Candidate = QuickStyleEntry & { uiPriority: number; quick: boolean };
  const all: Candidate[] = [];
  const seen = new Set<string>();
  const push = (id: string, style: StyleEntry): void => {
    if (seen.has(id)) return;
    seen.add(id);
    all.push({
      id,
      name: style.name || id,
      uiPriority: style.uiPriority ?? 9999,
      quick: !!style.quickFormat,
    });
  };
  for (const ps of styles.paragraphStyles ?? []) push(ps.id, ps);
  // Built-in named styles nested under `default`. Skip the `document` slot
  // (docDefaults, not a named style) and the keys that hold character styles
  // (the gallery is paragraph-only). Cast: DefaultStylesOptions has no string
  // index signature and mixes paragraph/character value types.
  const defaults = styles.default as unknown as Record<string, StyleEntry | undefined>;
  for (const [key, style] of Object.entries(defaults)) {
    if (key === "document" || CHARACTER_DEFAULT_KEYS.has(key) || !style) continue;
    push(pStyleIdFromKey(key), style);
  }

  const byPriority = (a: Candidate, b: Candidate): number => a.uiPriority - b.uiPriority;
  const quick = all.filter((s) => s.quick).sort(byPriority);
  return (quick.length > 0 ? quick : all).map(({ id, name }) => ({ id, name }));
}

/** Resolve the effective run-level properties (font name, size in points) at the
 *  caret, staying in the document's own units — no px conversion. Priority:
 *  direct run props (the textStyle mark) → the paragraph style (`styleId`) →
 *  its `basedOn` chain → the document defaults. `font` is resolved to a single
 *  display name (ascii/hAnsi/eastAsia). Returns null where nothing in the chain
 *  sets a property, so the caller can leave the box empty. */
export function effectiveRunProps(
  styles: StylesOptions | null | undefined,
  styleId: string | null | undefined,
  direct?: { font?: unknown; size?: unknown },
): { font: string | null; size: number | null } {
  let font: string | null = null;
  let size: number | null = null;

  // 1. Direct run props at the caret (textStyle mark) — highest priority.
  if (direct) {
    font = resolveFontName(direct.font);
    if (typeof direct.size === "number" && direct.size > 0) size = direct.size;
  }

  // 2. Paragraph style (styleId) → basedOn chain, via the same merge the
  //    renderer uses, so the box matches the rendered page. A pStyle-less
  //    paragraph renders as the default paragraph style (OOXML), so fall back to
  //    it when styleId is absent — matching what `.docx-default` renders.
  if ((font == null || size == null) && styles) {
    const effStyleId = styleId || defaultParagraphStyleId(styles);
    const { run } = mergeStyleChain(indexParagraphStyles(styles), effStyleId);
    if (font == null) font = resolveFontName(run.font);
    if (size == null && typeof run.size === "number" && run.size > 0) size = run.size;

    // 3. Document defaults (docDefaults run) — the final fallback.
    if (font == null || size == null) {
      const docRun = styles.default?.document?.run as Record<string, unknown> | undefined;
      if (docRun) {
        if (font == null) font = resolveFontName(docRun.font);
        if (size == null && typeof docRun.size === "number" && docRun.size > 0) size = docRun.size;
      }
    }
  }

  return { font, size };
}

// ── Attr/border helpers (shared by resolve + compile) ────────────────────────

/** Remove keys with null/undefined values. */
export function cleanAttrs(attrs: Record<string, unknown>): Record<string, unknown> {
  const result: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (value !== null && value !== undefined) result[key] = value;
  }
  return result;
}

/** Build a paragraph node from a resolved ParagraphOptions: reflective attrs
 *  parse, inline content, and null-stripped attrs. Shared by resolveParagraph's
 *  plain fallback and the list-item paragraph path so the build stays DRY.
 *  `contentPara` overrides the content source — a list item strips its task
 *  checkbox before resolving content, but attrs still come from the original. */
export function buildTextBlock(
  type: string,
  resolved: ParagraphOptions,
  ctx: ResolveContext,
  contentPara?: ParagraphOptions,
): JSONContent {
  const attrs = ctx.parseNodeAttrs(type, resolved as unknown as Record<string, unknown>);
  const content = ctx.resolveInlineContent(contentPara ?? resolved);
  const cleaned = cleanAttrs(attrs);
  const node: JSONContent = { type };
  if (Object.keys(cleaned).length > 0) node.attrs = cleaned;
  if (content.length > 0) node.content = content;
  return node;
}

/** True when a tblBorders object carries no REAL edge — every side is absent,
 *  none, or nil. office-open fills table.borders with all-`none` when the
 *  table's own <w:tblPr> defines no <w:tblBorders>, so this detects "the table
 *  has no borders of its own" to decide whether a referenced table style's
 *  borders should fill the gap. */
export function allBordersNone(borders: unknown): boolean {
  if (!borders || typeof borders !== "object") return true;
  const b = borders as Record<string, BorderOptions | undefined>;
  return (["top", "bottom", "left", "right", "insideHorizontal", "insideVertical"] as const).every(
    (k) => {
      const v = b[k];
      return !v || v.style === "none" || v.style === "nil";
    },
  );
}

/** Merge consecutive text nodes with the same marks. Used by inline container
 *  resolution (hyperlink, track-change) so a link/revision range spanning
 *  multiple runs becomes a single text node carrying the container mark. */
export function mergeTextNodes(nodes: JSONContent[]): JSONContent[] {
  const result: JSONContent[] = [];
  for (const node of nodes) {
    if (node.type === "text" && result.length > 0 && result[result.length - 1].type === "text") {
      const prev = result[result.length - 1];
      if (JSON.stringify(prev.marks) === JSON.stringify(node.marks)) {
        prev.text = (prev.text ?? "") + (node.text ?? "");
        continue;
      }
    }
    result.push({ ...node });
  }
  return result;
}
