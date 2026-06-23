import type { StylesOptions } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";

import { renderParagraphStyles, renderRunStyles, resolveFontName } from "../extensions/utils";

// Re-export the public styles model type so consumers (the editor's Styles
// gallery) type against office-open's source of truth instead of a local
// mirror. `StyleOptions` (singular) is an internal office-open interface and
// is not exported from the package, so we derive a style-entry type below.
export type { StylesOptions };

/** A named style entry as office-open models it: BaseParagraphStyleOptions or
 *  BaseCharacterStyleOptions (both extend the internal StyleOptions, carrying
 *  name/uiPriority/quickFormat). Derived from the public StylesOptions — not
 *  imported — because StyleOptions is not a public export of @office-open/docx. */
type StyleEntry =
  | NonNullable<StylesOptions["paragraphStyles"]>[number]
  | NonNullable<StylesOptions["characterStyles"]>[number];

/** Escape a style id for safe use in a CSS class selector. OOXML style ids are
 *  NCNames so this is rarely needed, but it keeps arbitrary ids safe. */
function escapeClass(id: string): string {
  return id.replace(/[^A-Za-z0-9_-]/g, (ch) => `\\${ch}`);
}

/** The pStyle val (class hook `docx-style-{id}`) for a built-in named style
 *  nested under DefaultStylesOptions: the key with its first letter upper-cased
 *  ("heading1" → "Heading1", "title" → "Title", "listParagraph" → "ListParagraph").
 *  This matches office-open's HeadingLevel literals / pStyle ids, so we derive
 *  the id from the key instead of hard-coding a name table. */
function pStyleIdFromKey(key: string): string {
  return key.charAt(0).toUpperCase() + key.slice(1);
}

/** The styleId of the document's default paragraph style (`w:default="1"`
 *  type="paragraph") — the implicit style applied to every paragraph WITHOUT an
 *  explicit pStyle. OOXML renders a pStyle-less paragraph as this style (usually
 *  "Normal"), so the editor's `.docx-default` class — emitted by Paragraph/
 *  Heading renderHTML on a styleId-less node — targets it. Searched in
 *  `paragraphStyles` and the built-in named styles nested under `default`
 *  (key → pStyle id). null when the document declares none. */
export function defaultParagraphStyleId(styles: StylesOptions | null | undefined): string | null {
  if (!styles) return null;
  for (const ps of styles.paragraphStyles ?? []) {
    // `default` (w:default="1") is on the runtime shape but not the public
    // StyleOptions type — read it loosely.
    if ((ps as { default?: boolean }).default) return ps.id;
  }
  const defaults = styles.default as unknown as Record<string, StyleEntry | undefined>;
  for (const [key, style] of Object.entries(defaults ?? {})) {
    if (key === "document" || !style) continue;
    if ((style as { default?: boolean }).default) return pStyleIdFromKey(key);
  }
  return null;
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
  if (doc) {
    const runDecls = doc.run ? renderRunStyles(doc.run as Record<string, unknown>) : [];
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
    if (Object.keys(run).length) decls.push(...renderRunStyles(run));
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
    if (cs.run) decls.push(...renderRunStyles(cs.run as Record<string, unknown>));
    if (decls.length)
      rules.push(`${within}.docx-char-${escapeClass(cs.id)} { ${decls.join(";")} }`);
  }

  return rules.join("\n");
}

/** Build an id → style-entry index over every paragraph style: the explicit
 *  `paragraphStyles` plus the built-in named styles nested under `default`
 *  (key → pStyle id via pStyleIdFromKey). `document` is docDefaults, not a
 *  named style, so it is excluded. A built-in that also appears in
 *  paragraphStyles is deduped by id (paragraphStyles wins on insertion order). */
function indexParagraphStyles(styles: StylesOptions): Map<string, StyleEntry> {
  const byId = new Map<string, StyleEntry>();
  for (const ps of styles.paragraphStyles ?? []) byId.set(ps.id, ps);
  const defaults = styles.default as unknown as Record<string, StyleEntry | undefined>;
  for (const [key, style] of Object.entries(defaults ?? {})) {
    if (key === "document" || !style) continue;
    byId.set(pStyleIdFromKey(key), style);
  }
  return byId;
}

/** Merge a paragraph style's run/paragraph properties with its `basedOn` chain
 *  (root first, child overrides per-property) — the OOXML inheritance model. A
 *  flat property-level merge: a child property overrides the parent's; an unset
 *  property is inherited. Shared by stylesToCss (rendering) and
 *  effectiveRunProps (the caret resolver) so the gallery box and the rendered
 *  page resolve identical values. */
function mergeStyleChain(
  byId: Map<string, StyleEntry>,
  styleId: string | null | undefined,
): { run: Record<string, unknown>; paragraph: Record<string, unknown> } {
  const chain: StyleEntry[] = [];
  const visited = new Set<string>();
  let curId = styleId || undefined;
  while (curId && !visited.has(curId)) {
    visited.add(curId);
    const style = byId.get(curId);
    if (!style) break;
    chain.unshift(style); // root first, so children override
    curId = style.basedOn ?? undefined;
  }
  const run: Record<string, unknown> = {};
  const paragraph: Record<string, unknown> = {};
  for (const style of chain) {
    // StyleEntry is a paragraph|character union; paragraph props live only on
    // the paragraph side, so access via a loose record.
    const s = style as unknown as Record<string, unknown>;
    if (s.run) Object.assign(run, s.run);
    if (s.paragraph) Object.assign(paragraph, s.paragraph);
  }
  return { run, paragraph };
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

/** Resolve one node's style inheritance: a paragraph/heading absorbs its
 *  styleId's basedOn chain (paragraph properties + the default `run`); a
 *  textStyle mark absorbs its character style's run. Explicit attrs/marks
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
  if ((type === "paragraph" || type === "heading") && attrs?.styleId) {
    const { paragraph, run } = mergeStyleChain(paraById, attrs.styleId as string);
    // resolved chain as the base; explicit attrs (incl. attrs.run) override.
    const mergedRun = { ...run, ...(attrs.run as Record<string, unknown>) };
    out.attrs = { ...paragraph, ...attrs, run: mergedRun };
  }
  if (node.marks) {
    out.marks = node.marks.map((m) => {
      if (m.type !== "textStyle") return m;
      const sid = (m.attrs as Record<string, unknown> | undefined)?.styleId as string | undefined;
      const crun = sid ? charRunById.get(sid) : undefined;
      return crun ? { ...m, attrs: { ...crun, ...m.attrs } } : m;
    });
  }
  if (node.content) {
    out.content = node.content.map((c) => resolveNode(c, paraById, charRunById));
  }
  return out;
}

/**
 * Inline a document's styles into each node's attrs — a self-contained
 * JSONContent that round-trips and renders without the styles context. A
 * paragraph/heading absorbs its styleId's `basedOn` chain (paragraph properties
 * + the paragraph's default `run`); a textStyle mark absorbs its character
 * style's run. Explicit attrs/marks always override the inherited values;
 * styleId is preserved (the semantic reference is kept). The style definitions
 * themselves are untouched — only properties are copied onto nodes.
 *
 * Styles default to the document's own `attrs.styles` (the round-tripped
 * styles.xml model carried on the doc node), so `inlineStyles(doc)` needs no
 * second argument; pass `styles` explicitly to override. Use cases:
 * cross-document paste (the snippet carries its styling) and self-contained
 * export. The editor still renders via global styles — this only merges JSON
 * properties.
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
