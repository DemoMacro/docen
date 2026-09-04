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

/** Reference prefix for generated multilevel-list definitions (the List
 *  Library presets). */
export const MULTILEVEL_REFERENCE_PREFIX = "docen-multilevel";

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

// ── Multilevel List Library (多级列表样式库) ─────────────────────────────────
// Word's gallery presets, expressed as per-level (format, lvlText) shapes.
// Levels beyond the three that carry a preset's identity repeat the deepest
// shape with its level placeholder — Word defines all nine the same way.

/** One preset level's marker shape. */
export interface MultilevelPresetLevel {
  format: LevelsOptions["format"];
  text: string;
}

/** The format cycle Word's hybrid presets walk (1. / a. / i. / 4. / e. …). */
const HYBRID_CYCLE: MultilevelPresetLevel["format"][] = [
  LevelFormat.DECIMAL,
  LevelFormat.LOWER_LETTER,
  LevelFormat.LOWER_ROMAN,
];

/** The hybrid presets' lvlText per level — the level's own number only (no
 *  cascade), suffixed per preset. */
function hybridLevels(suffix: string): MultilevelPresetLevel[] {
  return Array.from({ length: 9 }, (_, level) => ({
    format: HYBRID_CYCLE[level % 3],
    text: `%${level + 1}${suffix}`,
  }));
}

/** The cascading presets' lvlText — every level chains its ancestors
 *  (%1.%2.%3.), suffixed per preset. */
function cascadeLevels(suffix: string): MultilevelPresetLevel[] {
  return Array.from({ length: 9 }, (_, level) => ({
    format: LevelFormat.DECIMAL,
    text: Array.from({ length: level + 1 }, (_, i) => `%${i + 1}`).join(".") + suffix,
  }));
}

/** CJK 章节形 — 第%1章 / 第%2节 / %3、… (chineseCounting for the first two
 *  levels, decimal 顿号 for the rest). */
function cjkChapterLevels(): MultilevelPresetLevel[] {
  return Array.from({ length: 9 }, (_, level) => {
    if (level === 0) return { format: LevelFormat.CHINESE_COUNTING, text: "第%1章" };
    if (level === 1) return { format: LevelFormat.CHINESE_COUNTING, text: "第%2节" };
    return { format: LevelFormat.DECIMAL, text: `%${level + 1}、` };
  });
}

/** The List Library presets a ribbon pick maps to. `custom` is not here —
 *  the Define New Multilevel List dialog writes its own definition into the
 *  document's numbering instead of a generated reference. */
export const MULTILEVEL_PRESETS: Readonly<Record<string, MultilevelPresetLevel[]>> = {
  // 1. / 1.1. / 1.1.1. — the cascading decimal list (the editor's default).
  cascade: cascadeLevels("."),
  // 1) / 1.1) / 1.1.1) — cascading, right parenthesis.
  "cascade-paren": cascadeLevels(")"),
  // 1. / a. / i. — Word's hybrid cycle (decimal → letter → roman, repeating).
  hybrid: hybridLevels("."),
  // 1) / a) / i)
  "hybrid-paren": hybridLevels(")"),
  // 一、 / （一） / 1. — the Chinese Word staple (中文编号).
  cjk: [
    { format: LevelFormat.CHINESE_COUNTING, text: "%1、" },
    { format: LevelFormat.CHINESE_COUNTING, text: "（%2）" },
    { format: LevelFormat.DECIMAL, text: "%3." },
    ...Array.from({ length: 6 }, (_, level) => ({
      format: LevelFormat.DECIMAL,
      text: `%${level + 4}、`,
    })),
  ],
  // 第X章 / 第X节 / X、 — the chapter/section form (中文法律/学术).
  "cjk-chapter": cjkChapterLevels(),
};

/** The preset id a generated multilevel reference carries, or null. */
export function multilevelPresetOf(reference: string): string | null {
  const m = new RegExp(`^${MULTILEVEL_REFERENCE_PREFIX}-([a-z-]+)-\\d+$`).exec(reference);
  return m?.[1] ?? null;
}

/** Build nine levels from a preset's shapes: the preset defines the format
 *  and marker text; every level restarts at 1 and indents like Word's lists
 *  (0.5" per level, 0.25" hanging). */
export function buildMultilevelLevels(preset: MultilevelPresetLevel[]): LevelsOptions[] {
  return preset.map((lvl, level) => ({
    level,
    format: lvl.format,
    start: 1,
    text: lvl.text,
    paragraph: {
      indent: { left: 720 * (level + 1), hanging: 360 },
    },
  }));
}

/** Build the nine-level definition a Define New Multilevel List commit
 *  produces: the dialog's defined levels verbatim (its `format` is already
 *  the w:numFmt token), deeper levels extending the deepest defined shape —
 *  a level-3 `%3.` defines levels 4+ as `%4.`, `%5.`, … */
export function buildCustomMultilevelLevels(
  levels: { format: string; text: string }[],
): LevelsOptions[] {
  const full: MultilevelPresetLevel[] = Array.from({ length: 9 }, (_, i) => {
    const src = levels[Math.min(i, levels.length - 1)];
    return {
      format: src.format as MultilevelPresetLevel["format"],
      text: i < levels.length ? src.text : src.text.replaceAll(`%${levels.length}`, `%${i + 1}`),
    };
  });
  return buildMultilevelLevels(full);
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
  const preset = multilevelPresetOf(reference);
  if (preset && MULTILEVEL_PRESETS[preset])
    return buildMultilevelLevels(MULTILEVEL_PRESETS[preset]);
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

/** The next free generated multilevel reference for a preset (one past the
 *  highest existing `docen-multilevel-…` suffix — presets share the counter
 *  since a re-apply rewrites the reference wholesale). */
export function nextMultilevelReference(
  numberingRefs: Iterable<string>,
  numbering: unknown,
  preset: string,
): string {
  let max = 0;
  const re = new RegExp(`^${MULTILEVEL_REFERENCE_PREFIX}-[a-z-]+-(\\d+)$`);
  const consider = (ref: string): void => {
    const m = re.exec(ref);
    if (m) max = Math.max(max, Number(m[1]));
  };
  for (const ref of numberingRefs) consider(ref);
  const defs = (numbering as { abstractNumberings?: { reference?: string }[] } | undefined)
    ?.abstractNumberings;
  for (const def of defs ?? []) if (def.reference) consider(def.reference);
  return `${MULTILEVEL_REFERENCE_PREFIX}-${preset}-${max + 1}`;
}

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
