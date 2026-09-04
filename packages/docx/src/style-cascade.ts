// Style cascade — resolving a document's styles.xml model (StylesOptions):
// the id index, the default paragraph style, and the basedOn-chain merge.
// Rendering-neutral: the layout projection, the CSS route, and the editor's
// caret/gallery resolvers all share these primitives, so one cascade runs
// everywhere.

import type {
  StylesOptions,
  TableBordersOptions,
  TableOptions,
  TableStyleOptions,
} from "@office-open/docx";

/** Table-level cell margins (w:tblCellMar) — TableCellMarginOptions is not
 *  exported, derive it from the field that carries it. */
type TableCellMargins = NonNullable<TableOptions["margins"]>;

/** A named style entry as office-open models it: BaseParagraphStyleOptions or
 *  BaseCharacterStyleOptions (both extend the internal StyleOptions, carrying
 *  name/uiPriority/quickFormat). Derived from the public StylesOptions — not
 *  imported — because StyleOptions is not a public export of @office-open/docx. */
export type StyleEntry =
  | NonNullable<StylesOptions["paragraphStyles"]>[number]
  | NonNullable<StylesOptions["characterStyles"]>[number];

/** The pStyle val (the style's OOXML id) for a built-in named style nested
 *  under DefaultStylesOptions: the key with its first letter upper-cased
 *  ("heading1" → "Heading1", "title" → "Title", "listParagraph" →
 *  "ListParagraph"). This matches office-open's HeadingLevel literals / pStyle
 *  ids, so we derive the id from the key instead of hard-coding a name table. */
export function pStyleIdFromKey(key: string): string {
  return key.charAt(0).toUpperCase() + key.slice(1);
}

/** The styleId of the document's default paragraph style (`w:default="1"`
 *  type="paragraph") — the implicit style applied to every paragraph WITHOUT an
 *  explicit pStyle. OOXML renders a pStyle-less paragraph as this style (usually
 *  "Normal"). Searched in `paragraphStyles` and the built-in named styles nested
 *  under `default` (key → pStyle id). null when the document declares none. */
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

/** Build an id → style-entry index over every paragraph style: the explicit
 *  `paragraphStyles` plus the built-in named styles nested under `default`
 *  (key → pStyle id via pStyleIdFromKey). `document` is docDefaults, not a
 *  named style, so it is excluded. A built-in that also appears in
 *  paragraphStyles is deduped by id — the built-in wins, being set second. */
// Cache the style index by the styles object reference. A document's styles
// model is stable for its lifetime (set on load, unchanged across edits), yet
// indexParagraphStyles is called per-paragraph (detectHeadingLevel during
// resolve), per-transaction (effectiveRunProps at the caret), and per-render
// (layout projection). The WeakMap memo turns all of those into O(1) lookups after
// the first build and frees the entry when the styles object is GC'd. Callers
// treat the result as read-only (mergeStyleChain only .get()s).
const styleIndexCache = new WeakMap<StylesOptions, Map<string, StyleEntry>>();

export function indexParagraphStyles(styles: StylesOptions): Map<string, StyleEntry> {
  const cached = styleIndexCache.get(styles);
  if (cached) return cached;
  const byId = new Map<string, StyleEntry>();
  for (const ps of styles.paragraphStyles ?? []) byId.set(ps.id, ps);
  const defaults = styles.default as unknown as Record<string, StyleEntry | undefined>;
  for (const [key, style] of Object.entries(defaults ?? {})) {
    if (key === "document" || !style) continue;
    byId.set(pStyleIdFromKey(key), style);
  }
  styleIndexCache.set(styles, byId);
  return byId;
}

/** The `default` keys that carry character styles (DefaultStylesOptions types
 *  them CharacterStyleOptions; every other key is a paragraph style). */
const CHARACTER_DEFAULT_KEYS = [
  "hyperlink",
  "footnoteReference",
  "footnoteTextChar",
  "endnoteReference",
  "endnoteTextChar",
] as const;

/** Build an id → style-entry index over every character style: the explicit
 *  `characterStyles` plus the built-in character styles nested under `default`
 *  (key → style id via pStyleIdFromKey, e.g. "hyperlink" → "Hyperlink"). A
 *  built-in that also appears in characterStyles is deduped by id. WeakMap-
 *  cached per styles object, like indexParagraphStyles — the projection
 *  resolves a run's w:rStyle per run, only a handful of distinct ids exist. */
const characterStyleIndexCache = new WeakMap<StylesOptions, Map<string, StyleEntry>>();

export function indexCharacterStyles(styles: StylesOptions | undefined): Map<string, StyleEntry> {
  if (!styles) return new Map();
  const cached = characterStyleIndexCache.get(styles);
  if (cached) return cached;
  const byId = new Map<string, StyleEntry>();
  for (const cs of styles.characterStyles ?? []) byId.set(cs.id, cs);
  const defaults = styles.default as unknown as Record<string, StyleEntry | undefined>;
  for (const key of CHARACTER_DEFAULT_KEYS) {
    const style = defaults?.[key];
    if (style) byId.set(pStyleIdFromKey(key), style);
  }
  characterStyleIndexCache.set(styles, byId);
  return byId;
}

/** id → table-style index, WeakMap-cached per styles object like the
 *  paragraph/character indexes — the projection resolves a table's w:tblStyle
 *  per table, per transaction. */
const tableStyleIndexCache = new WeakMap<StylesOptions, Map<string, TableStyleOptions>>();

export function indexTableStyles(
  styles: StylesOptions | undefined,
): Map<string, TableStyleOptions> {
  if (!styles) return new Map();
  const cached = tableStyleIndexCache.get(styles);
  if (cached) return cached;
  const byId = new Map<string, TableStyleOptions>();
  for (const ts of styles.tableStyles ?? []) byId.set(ts.id, ts);
  tableStyleIndexCache.set(styles, byId);
  return byId;
}

/** Whether `v` is a plain object — an OOXML property group (spacing/indent/
 *  border/shading/font) that merges key by key — as opposed to an array
 *  (tabStops) or scalar, which replace. */
function isPlainObject(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

/** Deep-merge `source` into `target` (mutates target) per the OOXML `basedOn`
 *  model: nested property groups merge key by key (a child's spacing.before
 *  merges with, not replaces, the parent's spacing.line); arrays and scalars
 *  replace. Nullish source values are skipped so an unset child key doesn't
 *  clobber an inherited value. */
export function deepMergeInto(
  target: Record<string, unknown>,
  source: Record<string, unknown>,
): Record<string, unknown> {
  for (const [key, srcVal] of Object.entries(source)) {
    if (srcVal === null || srcVal === undefined) continue;
    const tgtVal = target[key];
    target[key] =
      isPlainObject(srcVal) && isPlainObject(tgtVal)
        ? deepMergeInto({ ...tgtVal }, srcVal)
        : isPlainObject(srcVal)
          ? { ...srcVal }
          : srcVal;
  }
  return target;
}

/** Merge a paragraph style's run/paragraph properties with its `basedOn` chain
 *  (root first, child overrides per-property) — the OOXML inheritance model.
 *  Nested property groups (spacing/indent/border/font) merge key by key; arrays
 *  and scalars replace. Shared by every style consumer (the layout projection,
 *  the CSS route, the caret resolver) so all of them resolve identical values. */
// Memoize mergeStyleChain per (byId, styleId). byId is itself memoized by
// indexParagraphStyles (one Map per StylesOptions object), so this WeakMap
// frees the cache when the styles object is GC'd. The same styleId is resolved
// for every paragraph/run that carries it (thousands of calls on a large doc,
// only dozens of distinct ids) — the chain walk + deepMergeInto is repeated
// work. Callers treat the result as read-only (consumers only read fields;
// resolveNode spreads `{...paragraph}` before merging its own attrs), so a
// shared cached value is safe.
const styleChainCache = new WeakMap<
  Map<string, StyleEntry>,
  Map<string, { run: Record<string, unknown>; paragraph: Record<string, unknown> }>
>();

export function mergeStyleChain(
  byId: Map<string, StyleEntry>,
  styleId: string | null | undefined,
): { run: Record<string, unknown>; paragraph: Record<string, unknown> } {
  if (!styleId) return { run: {}, paragraph: {} };
  const perId = styleChainCache.get(byId);
  if (perId) {
    const cached = perId.get(styleId);
    if (cached) return cached;
  }
  const chain: StyleEntry[] = [];
  const visited = new Set<string>();
  let curId: string | undefined = styleId;
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
    if (s.run) deepMergeInto(run, s.run as Record<string, unknown>);
    if (s.paragraph) deepMergeInto(paragraph, s.paragraph as Record<string, unknown>);
  }
  const result = { run, paragraph };
  const bucket =
    perId ??
    new Map<string, { run: Record<string, unknown>; paragraph: Record<string, unknown> }>();
  if (!perId) styleChainCache.set(byId, bucket);
  bucket.set(styleId, result);
  return result;
}

/** Resolve a table style's effective table-level properties (tblBorders,
 *  tblCellMar) by walking its basedOn chain (root first, child overrides) —
 *  the table-style counterpart of mergeStyleChain. office-open's parseDocument
 *  does NOT merge the referenced <w:tblStyle> into table.borders/cellMargin
 *  (they reflect only the table's own <w:tblPr>), so a "Table Grid" table whose
 *  borders live in the style needs this to render its grid. Returns empty when
 *  the style is absent or unknown. */
export function mergeTableStyleProps(
  tableStyles: StylesOptions["tableStyles"],
  styleId: string | null | undefined,
): { borders?: TableBordersOptions; margins?: TableCellMargins } {
  if (!styleId || !tableStyles) return {};
  const byId = new Map(tableStyles.map((t) => [t.id, t]));
  const chain: NonNullable<StylesOptions["tableStyles"]> = [];
  const visited = new Set<string>();
  let cur: string | undefined = styleId ?? undefined;
  while (cur && !visited.has(cur)) {
    visited.add(cur);
    const s = byId.get(cur);
    if (!s) break;
    chain.unshift(s); // root first → children override below
    cur = s.basedOn;
  }
  // Per-key inheritance: a child style's tblBorders/tblCellMar overrides only
  // the edges it declares — Word keeps the basedOn chain's remaining edges,
  // so a whole-object swap would lose the parent's other sides.
  let borders: TableBordersOptions | undefined;
  let margins: TableCellMargins | undefined;
  for (const s of chain) {
    const t = s.table;
    if (!t) continue;
    if (t.borders) borders = { ...borders, ...t.borders };
    if (t.margins) margins = { ...margins, ...t.margins };
  }
  const out: { borders?: TableBordersOptions; margins?: TableCellMargins } = {};
  if (borders) out.borders = borders;
  if (margins) out.margins = margins;
  return out;
}
