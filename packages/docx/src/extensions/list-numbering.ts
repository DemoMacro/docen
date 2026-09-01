import type { LevelsOptions } from "@office-open/docx";
import { LevelFormat } from "@office-open/docx";

/**
 * Numbering definitions for the flat list model — a list paragraph is a plain
 * paragraph carrying `bullet {level}` / `numbering {reference, level}` attrs
 * (no list tree). Editor-created lists reference a generated abstractNum
 * (`docen-bullet` / `docen-ordered-<n>`); compile registers any referenced
 * definition missing from the source numbering, and the ribbon's marker
 * variants (●/○/■, decimal/alpha/roman) map to per-variant references.
 */

/** Reference prefix for generated ordered-list abstractNum definitions. */
export const ORDERED_REFERENCE_PREFIX = "docen-ordered";

/** Reference for the editor's default bullet list definition. */
export const BULLET_REFERENCE = "docen-bullet";

/** Placeholder ordered reference stamped by HTML/Markdown import (their list
 *  syntax carries no numbering identity); assignOrderedReferences rewrites
 *  each consecutive run to a fresh docen-ordered-* reference. */
export const HTML_ORDERED_TEMP = "html-ordered";

/** Ribbon bullet-marker variants: dropdown value → level-0 glyph. */
export const BULLET_GLYPHS: Readonly<Record<string, string>> = {
  bullet: "●",
  circle: "○",
  square: "■",
};

/** Ribbon ordered-list formats: dropdown value → w:numFmt of level 0. */
export const ORDERED_FORMATS: Readonly<Record<string, LevelsOptions["format"]>> = {
  decimal: LevelFormat.DECIMAL,
  "lower-alpha": LevelFormat.LOWER_LETTER,
  "lower-roman": LevelFormat.LOWER_ROMAN,
};

/** lvlText per nesting depth — Word's default multilevel shape: every level
 *  cascades its ancestors ("%1.", "%1.%2.", "%1.%2.%3.", …), so stepping a
 *  paragraph's level visibly renumbers it. */
const ORDERED_LEVEL_TEXT = Array.from(
  { length: 9 },
  (_, level) => Array.from({ length: level + 1 }, (_, i) => `%${i + 1}`).join(".") + ".",
);

/** Word's built-in bullet level cycle (●/○/■), repeated past depth 3. */
const LEVEL_GLYPH_CYCLE = ["●", "○", "■"];

/** Build nine ordered levels; level 0 carries `format`, deeper levels decimal
 *  (Word's library lists only restyle the top level). Each level restarts at 1
 *  and indents like Word's built-in lists (0.5" per level, 0.25" hanging). */
export function buildOrderedLevels(
  format: LevelsOptions["format"] = LevelFormat.DECIMAL,
): LevelsOptions[] {
  return Array.from(
    { length: 9 },
    (_, level): LevelsOptions => ({
      level,
      format: level === 0 ? format : LevelFormat.DECIMAL,
      start: 1,
      text: ORDERED_LEVEL_TEXT[level],
      paragraph: {
        indent: { left: 720 * (level + 1), hanging: 360 },
      },
    }),
  );
}

/** Build nine bullet levels mirroring Word's built-in bullet indentation
 *  (0.5" per level, 0.25" hanging) and its glyph cycle (●/○/■); the variant
 *  glyph rides level 0. */
export function buildBulletLevels(glyph = "●"): LevelsOptions[] {
  return Array.from(
    { length: 9 },
    (_, level): LevelsOptions => ({
      level,
      format: LevelFormat.BULLET,
      text: level === 0 ? glyph : LEVEL_GLYPH_CYCLE[level % 3],
      paragraph: {
        indent: { left: 720 * (level + 1), hanging: 360 },
      },
    }),
  );
}

/** The definition a generated list reference compiles to: a bullet reference
 *  builds bullet levels (glyph from the reference suffix), an ordered
 *  reference builds ordered levels (format from the suffix). Unknown
 *  references (round-tripped `list_<numId>`) are not generated — their
 *  definitions travel in the source numbering. */
export function buildListLevels(reference: string): LevelsOptions[] | null {
  if (reference === BULLET_REFERENCE) return buildBulletLevels();
  if (reference.startsWith(`${BULLET_REFERENCE}-`)) {
    return buildBulletLevels(BULLET_GLYPHS[reference.slice(BULLET_REFERENCE.length + 1)] ?? "●");
  }
  const ordered = reference.match(new RegExp(`^${ORDERED_REFERENCE_PREFIX}-(?:([a-z-]+)-)?\\d+$`));
  if (ordered) {
    return buildOrderedLevels(
      ordered[1] ? (ORDERED_FORMATS[ordered[1]] ?? LevelFormat.DECIMAL) : LevelFormat.DECIMAL,
    );
  }
  return null;
}

/** True when a numbering reference is one of ours (compile registers missing
 *  definitions only for generated references). */
export function isGeneratedListReference(reference: string): boolean {
  return buildListLevels(reference) !== null;
}

/** The next free generated ordered reference: one past the highest existing
 *  `docen-ordered-…` suffix (counters shared across variants so references
 *  never collide), so every editor-created list numbers independently. Scans
 *  the numbering definitions and every list paragraph. A variant
 *  ("lower-alpha"…) names the level-0 format (see ORDERED_FORMATS). */
export function nextOrderedReference(
  numberingRefs: Iterable<string>,
  numbering: unknown,
  variant?: string,
): string {
  let max = 0;
  const re = new RegExp(`^${ORDERED_REFERENCE_PREFIX}(?:-[a-z-]+)?-(\\d+)$`);
  const consider = (ref: string): void => {
    const m = re.exec(ref);
    if (m) max = Math.max(max, Number(m[1]));
  };
  for (const ref of numberingRefs) consider(ref);
  const defs = (numbering as { abstractNumberings?: { reference?: string }[] } | undefined)
    ?.abstractNumberings;
  for (const def of defs ?? []) if (def.reference) consider(def.reference);
  return `${ORDERED_REFERENCE_PREFIX}-${variant ? `${variant}-` : ""}${max + 1}`;
}

/** A list paragraph's list attr (`numbering`/`bullet`) as a union of both
 *  shapes, for run scanning. */
type ListAttr = { reference?: string; level?: number };

/** Rewrite placeholder ordered references in place: each maximal run of
 *  consecutive paragraphs sharing the placeholder (a run ends at any
 *  non-member paragraph) becomes one fresh docen-ordered-* reference, so
 *  imported sibling lists number independently. */
export function assignOrderedReferences<T extends { content?: unknown[]; attrs?: unknown }>(
  json: T,
): T {
  const content = json.content;
  if (!Array.isArray(content)) return json;
  let refs: string[] = [];
  for (const node of content) {
    const ref = (node as { attrs?: { numbering?: ListAttr } })?.attrs?.numbering?.reference;
    if (typeof ref === "string") refs.push(ref);
  }
  const numbering = (json.attrs as { numbering?: unknown } | undefined)?.numbering;
  let run: number[] = [];
  const flush = (): void => {
    if (run.length === 0) return;
    const reference = nextOrderedReference(refs, numbering);
    for (const i of run) {
      const attrs = (content[i] as { attrs?: { numbering?: ListAttr } }).attrs;
      if (attrs?.numbering) attrs.numbering.reference = reference;
    }
    refs.push(reference);
    run = [];
  };
  content.forEach((node, i) => {
    const attrs = (node as { attrs?: { numbering?: ListAttr } })?.attrs;
    if (attrs?.numbering?.reference === HTML_ORDERED_TEMP) run.push(i);
    else flush();
  });
  flush();
  return json;
}
