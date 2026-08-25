import type { BorderOptions, ParagraphOptions, StylesOptions } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";

import type { ResolveContext } from "../extensions/types";
import { renderParagraphStyles, renderRunStyles, resolveFontName } from "../extensions/utils";
import {
  defaultParagraphStyleId,
  deepMergeInto,
  indexParagraphStyles,
  mergeStyleChain,
  pStyleIdFromKey,
  type StyleEntry,
} from "../style-cascade";

// Re-export the public styles model type so consumers (the editor's Styles
// gallery) type against office-open's source of truth instead of a local
// mirror. The style-entry type and every cascade primitive live in
// style-cascade.ts (rendering-neutral); this module owns the CSS route.
export type { StylesOptions };

/** Escape a style id for safe use in a CSS class selector. OOXML style ids are
 *  NCNames so this is rarely needed, but it keeps arbitrary ids safe. */
function escapeClass(id: string): string {
  return id.replace(/[^A-Za-z0-9_-]/g, (ch) => `\\${ch}`);
}

/**
 * Generate scoped CSS from a document's `styles.xml` model (`StylesOptions`) so
 * named paragraph/character styles and the document defaults render correctly
 * in the editor. All rules are scoped to `.docen-page` (the editor's page
 * surface); the editor injects the result into a `<style>` after load.
 *
 * - `default.document` → `.docen-page` base run + paragraph defaults (the doc's
 *   default font/size, line spacing, …).
 * - `paragraphStyles[]` → `.docen-page .docx-style-{id}` (run + paragraph CSS).
 *   Paragraph nodes carry `class="docx-style-{styleId}"` (see the Paragraph /
 *   Heading extensions' pStyle ↔ styleId round-trip).
 * - `characterStyles[]` → `.docen-page .docx-char-{id}` (run CSS).
 *
 * `basedOn` inheritance is deep-merged per-property into each named style's
 * rule (root first, child overrides) — see mergeStyleChain. CSS inheritance
 * follows the DOM tree, not a style's basedOn, so each `.docx-style-{id}` rule
 * carries its full ancestor chain rather than relying on source order.
 */
export function stylesToCss(styles: StylesOptions | null | undefined, scope: string): string {
  if (!styles) return "";
  const rules: string[] = [];
  // Descendant prefix for named-style rules. The caller picks the container
  // selector (the editor passes its page surface); @docen/docx itself owns no
  // such class, so the scope is parameter rather than hard-coded.
  const within = scope ? `${scope} ` : "";

  // Document defaults (w:docDefaults) — the base layer applied to EVERY
  // paragraph before named styles override (ECMA-376). Split by target:
  //  - run defaults (font/size/...) → the page surface base, inherited by all.
  //  - paragraph defaults (spacing.line, alignment, indent…) → the paragraph
  //    selectors p/h1-h6, NOT the page surface. The page's own line-height is
  //    the document-grid single pitch (page-node sets it inline); a docDefaults
  //    spacing.line placed on `.docen-page` would be overridden by that inline
  //    (inline > stylesheet) — and it is semantically a per-paragraph multiple
  //    of the grid pitch, so it belongs on paragraphs. There it overrides the
  //    inherited single pitch. Named styles (.docx-style-*, specificity 0,2,0)
  //    outrank these (0,1,1) and keep their own spacing, matching OOXML.
  const doc = styles.default?.document;
  // docDefaults carries the doc's eastAsia language (rPrDefault/lang @eastAsia);
  // forward it so *Theme fonts resolve to the right CJK font (等线/游ゴシック/…).
  const docEastAsia =
    (doc?.run as { language?: { eastAsia?: string } } | undefined)?.language?.eastAsia ?? "zh-CN";
  if (doc) {
    const runDecls = doc.run
      ? renderRunStyles(doc.run as Record<string, unknown>, docEastAsia)
      : [];
    if (runDecls.length && scope) rules.push(`${scope} { ${runDecls.join(";")} }`);
    const paraDecls = doc.paragraph
      ? renderParagraphStyles(doc.paragraph as Record<string, unknown>)
      : [];
    if (paraDecls.length) {
      const paraSel = `${within}p, ${within}h1, ${within}h2, ${within}h3, ${within}h4, ${within}h5, ${within}h6`;
      rules.push(`${paraSel} { ${paraDecls.join(";")} }`);
    }
  }

  // Named paragraph styles (custom + built-in nested under `default`) →
  // .docx-style-{id}, each with its basedOn chain deep-merged so the class
  // carries inherited properties. indexParagraphStyles dedupes the two sources
  // (a built-in may also appear in paragraphStyles) by pStyle id.
  const byId = indexParagraphStyles(styles);
  // The default paragraph style (w:default="1") is the implicit style for any
  // paragraph without a pStyle (OOXML). Its rule also matches `.docx-default` —
  // the class Paragraph/Heading renderHTML emit on a styleId-less node — so the
  // 94%-of-paragraphs body text renders as the document's real body style.
  const defaultId = defaultParagraphStyleId(styles);
  for (const id of byId.keys()) {
    const { run, paragraph } = mergeStyleChain(byId, id);
    const decls: string[] = [];
    if (Object.keys(run).length) decls.push(...renderRunStyles(run, docEastAsia));
    if (Object.keys(paragraph).length) decls.push(...renderParagraphStyles(paragraph));
    if (!decls.length) continue;
    const selector =
      id === defaultId
        ? `${within}.docx-default, ${within}.docx-style-${escapeClass(id)}`
        : `${within}.docx-style-${escapeClass(id)}`;
    rules.push(`${selector} { ${decls.join(";")} }`);
  }

  // Named character styles.
  for (const cs of styles.characterStyles ?? []) {
    const decls: string[] = [];
    if (cs.run) decls.push(...renderRunStyles(cs.run as Record<string, unknown>, docEastAsia));
    if (decls.length)
      rules.push(`${within}.docx-char-${escapeClass(cs.id)} { ${decls.join(";")} }`);
  }

  return rules.join("\n");
}
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

// ── Style baking (flatten global styles into nodes) ──────────────────────────

/** Index a document's character styles (explicit `characterStyles` + the
 *  built-in character keys under `default`) by pStyle id → run properties. Core
 *  phase: a character style's own run only (its basedOn chain is not walked —
 *  rare in practice; paragraph styles above already deep-merge). */
function indexCharacterRunStyles(styles: StylesOptions): Map<string, Record<string, unknown>> {
  const byId = new Map<string, Record<string, unknown>>();
  for (const cs of styles.characterStyles ?? []) {
    if (cs.run) byId.set(cs.id, cs.run as Record<string, unknown>);
  }
  const defaults = styles.default as unknown as Record<string, StyleEntry | undefined>;
  for (const [key, style] of Object.entries(defaults ?? {})) {
    if (!style || !style.run || key === "document") continue;
    if (!CHARACTER_DEFAULT_KEYS.has(key)) continue;
    byId.set(pStyleIdFromKey(key), style.run as Record<string, unknown>);
  }
  return byId;
}

/** Resolve one node's style inheritance. A paragraph/heading absorbs its
 *  styleId's basedOn chain (paragraph properties + the default `run`); a
 *  textStyle mark absorbs its character style's run. Other nodes (table,
 *  tableRow, tableCell, image, …) are NOT baked here — they only recurse into
 *  `content`, so a cell's paragraphs still get baked (see inlineStyles for why
 *  tables and docDefaults are intentionally out of scope). Explicit attrs/marks
 *  override the inherited values; styleId is preserved. Pure JSON — no DOM, no
 *  marks pushed onto text. */
function resolveNode(
  node: JSONContent,
  paraById: Map<string, StyleEntry>,
  charRunById: Map<string, Record<string, unknown>>,
): JSONContent {
  const out: JSONContent = { ...node };
  const attrs = node.attrs as Record<string, unknown> | undefined;
  const type = node.type;
  if (type === "paragraph" && attrs?.style) {
    const { paragraph, run } = mergeStyleChain(paraById, attrs.style as string);
    // Baked chain is the base; explicit attrs override per-property. Use
    // deepMergeInto (not object spread) so a node's schema null defaults
    // (spacing:null, indent:null, …) don't clobber the baked values, and a
    // partial override (attrs.spacing.before) merges key by key instead of
    // replacing the whole group. Nullish attrs drop out — the baked value (or
    // nothing) remains, so no dead nulls linger on the baked node.
    const mergedAttrs = deepMergeInto({ ...paragraph }, attrs);
    mergedAttrs.run = deepMergeInto(
      { ...run },
      (attrs.run as Record<string, unknown> | null | undefined) ?? {},
    );
    out.attrs = mergedAttrs;
  }
  if (node.marks) {
    out.marks = node.marks.map((m) => {
      if (m.type !== "textStyle") return m;
      const sid = (m.attrs as Record<string, unknown> | undefined)?.style as string | undefined;
      const crun = sid ? charRunById.get(sid) : undefined;
      // deepMergeInto (not spread) so the mark's schema null defaults don't
      // clobber the baked character-style run (e.g. italic:null wiping the
      // style's italic:true); explicit attrs still override per-property.
      return crun ? { ...m, attrs: deepMergeInto({ ...crun }, m.attrs ?? {}) } : m;
    });
  }
  if (node.content) {
    out.content = node.content.map((c) => resolveNode(c, paraById, charRunById));
  }
  return out;
}

/**
 * Bake a document's named styles into each node's attrs so a snippet of
 * JSONContent can be lifted into ANOTHER document (DB storage → extract →
 * recombine) without its source styles.xml. The use case is fragment
 * recombination, not editor rendering — the editor never calls this (it
 * renders via stylesToCss global CSS); this is an offline util for
 * self-contained JSON.
 *
 * Baking is intentionally PARTIAL:
 *  - A paragraph/heading or textStyle mark WITH a styleId absorbs its basedOn
 *    chain, so a "Heading 1" stays styled across documents even when the
 *    target lacks that style. styleId is kept as the semantic reference; the
 *    baked props are the rendering fallback.
 *  - A node WITHOUT a styleId (body text, a default table) is LEFT AS-IS so it
 *    follows the TARGET document's body style after recombination. Baking
 *    docDefaults here would freeze font/size onto every body paragraph and
 *    defeat "change one body style to restyle them all".
 *  - Tables are NOT baked here: resolveTable already merged the table style
 *    (borders/cellMargin) and pushed insideH/V onto cells at parse time, so a
 *    table's JSON is already self-contained. stylesToCss cannot render the
 *    interior grid (CSS border-collapse puts it on cells), so that bake MUST
 *    stay in resolveDocument — moving it here would regress the editor.
 *
 * Explicit attrs/marks always override the baked values; the style definitions
 * themselves are untouched. Pure JSON — no DOM, no marks pushed onto text.
 * Styles default to the document's own `attrs.styles` (the round-tripped
 * styles.xml model on the doc node), so `inlineStyles(doc)` needs no second
 * argument; pass `styles` to override.
 */
export function inlineStyles(json: JSONContent, styles?: StylesOptions | null): JSONContent {
  // Prefer explicitly passed styles; fall back to the document's own attrs.styles
  // so a caller can resolve a full document with no extra argument.
  const docStyles =
    styles ??
    ((json.attrs as Record<string, unknown> | undefined)?.styles as StylesOptions | undefined);
  if (!docStyles) return json;
  return resolveNode(json, indexParagraphStyles(docStyles), indexCharacterRunStyles(docStyles));
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
