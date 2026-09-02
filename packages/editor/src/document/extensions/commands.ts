import type { ImageAttrs } from "@docen/docx";
import { BULLET_GLYPHS, nextOrderedReference, ORDERED_FORMATS } from "@docen/docx";
import { Extension } from "@docen/docx/core";
import type { Node as PMNode, ResolvedPos } from "@tiptap/pm/model";
import type { EditorState } from "@tiptap/pm/state";
import type { Transaction } from "@tiptap/pm/state";
import { NodeSelection, TextSelection } from "@tiptap/pm/state";

/**
 * Document editor commands (Office.js-style "add-in commands") as native
 * Tiptap commands.
 *
 * Each command name (kebab-case) IS a Tiptap command on `editor.commands`, so
 * every entry point — a ribbon click, a {@link DocenKeymap} shortcut, or a
 * programmatic call — routes as `editor.chain().focus()[name](value).run()`
 * with no mapping layer (no RIBBON_COMMAND_MAP, no dispatchRibbonCommand, no
 * addin.commands bridge). Names are 1:1 with the ribbon `event` attributes and
 * the `RIBBON_ICONS` keys, so a ribbon control, its keyboard shortcut, and
 * `editor.can(name)` all resolve to the one definition here.
 *
 * Simple marks/alignment/lists wrap the built-in Tiptap commands; indent /
 * spacing / shading / border / style / case / sort stamp the office-open
 * paragraph attrs (indent/spacing/shading/border) or manipulate the doc
 * directly via the `chain` prop. `editor.can()` works on every command, so the
 * ribbon can grey-out unavailable actions precisely.
 *
 * Document-specific: workbook (RevoGrid) and presentation (LeaferJS) have
 * their own engines and do not reuse it.
 */

// Type augmentation: register every command on `editor.commands` so callers
// get autocomplete + `editor.can()` works. Each name is also the ribbon
// `event` attribute, so #onCommand does editor.chain().focus()[event](value).
declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    documentCommands: {
      // Font marks
      bold: () => ReturnType;
      italic: () => ReturnType;
      underline: () => ReturnType;
      strike: () => ReturnType;
      subscript: () => ReturnType;
      superscript: () => ReturnType;
      highlight: (value?: string) => ReturnType;
      code: () => ReturnType;
      "clear-format": () => ReturnType;
      "font-name": (font?: string) => ReturnType;
      "font-size": (size?: string) => ReturnType;
      "grow-font": () => ReturnType;
      "shrink-font": () => ReturnType;
      // Paragraph
      "align-left": () => ReturnType;
      "align-center": () => ReturnType;
      "align-right": () => ReturnType;
      justify: () => ReturnType;
      "indent-increase": () => ReturnType;
      "indent-decrease": () => ReturnType;
      "line-spacing": (mult?: string) => ReturnType;
      shading: (value?: unknown) => ReturnType;
      "font-color": (value?: unknown) => ReturnType;
      border: (side?: string) => ReturnType;
      // Lists / blocks
      "bullet-list": (variant?: string) => ReturnType;
      "ordered-list": (variant?: string) => ReturnType;
      blockquote: () => ReturnType;
      "horizontal-rule": () => ReturnType;
      "page-break": () => ReturnType;
      "column-break": () => ReturnType;
      "section-break": () => ReturnType;
      "insert-table": () => ReturnType;
      "delete-table": () => ReturnType;
      // Table context commands (the Table Design / Layout contextual tabs).
      "insert-row-above": () => ReturnType;
      "insert-row-below": () => ReturnType;
      "insert-column-left": () => ReturnType;
      "insert-column-right": () => ReturnType;
      "delete-row": () => ReturnType;
      "delete-column": () => ReturnType;
      "select-table": () => ReturnType;
      "select-table-row": () => ReturnType;
      "select-table-cell": () => ReturnType;
      "select-table-column": () => ReturnType;
      "align-cell": (value?: string) => ReturnType;
      "repeat-header-rows": () => ReturnType;
      "cell-shading": (value?: unknown) => ReturnType;
      "text-direction": () => ReturnType;
      "convert-to-text": () => ReturnType;
      "table-style": (value?: string) => ReturnType;
      "table-borders": (value?: string) => ReturnType;
      "toggle-table-look": (value?: string) => ReturnType;
      "merge-cells": () => ReturnType;
      "split-cell": () => ReturnType;
      "split-table": () => ReturnType;
      "autofit-contents": () => ReturnType;
      "autofit-window": (value?: string) => ReturnType;
      "fixed-column-width": () => ReturnType;
      "distribute-columns": () => ReturnType;
      "cell-width": (value?: string) => ReturnType;
      "cell-height": (value?: string) => ReturnType;
      link: (href?: string) => ReturnType;
      style: (styleId?: string) => ReturnType;
      "add-text": (value?: string) => ReturnType;
      // Editing
      "change-case": (mode?: string) => ReturnType;
      sort: () => ReturnType;
      "multilevel-list": (level?: string) => ReturnType;
      // Picture — names align to Office.js InlinePicture (delete / left / top).
      "delete-picture": () => ReturnType;
      "position-picture": (value?: string) => ReturnType;
      // Arrange — floating drawings (z-order, wrap, rotation, position).
      "bring-forward": () => ReturnType;
      "send-backward": () => ReturnType;
      wrap: (value?: string) => ReturnType;
      rotate: (value?: string) => ReturnType;
      position: (value?: string) => ReturnType;
      "align-objects": (value?: string) => ReturnType;
    };
  }
}

/** Ribbon event names that route to a Tiptap command (the keys of the
 *  {@link DocumentCommands} extension). `<docen-document>` greys out any ribbon
 *  control whose `event` isn't here. */
export const WIRED_DISPATCH: ReadonlySet<string> = new Set([
  "bold",
  "italic",
  "underline",
  "strike",
  "subscript",
  "superscript",
  "highlight",
  "code",
  "clear-format",
  "font-name",
  "font-size",
  "grow-font",
  "shrink-font",
  "align-left",
  "align-center",
  "align-right",
  "justify",
  "indent-increase",
  "indent-decrease",
  "line-spacing",
  "shading",
  "font-color",
  "border",
  "bullet-list",
  "ordered-list",
  "blockquote",
  "horizontal-rule",
  "page-break",
  "column-break",
  "section-break",
  "insert-table",
  "delete-table",
  "insert-row-above",
  "insert-row-below",
  "insert-column-left",
  "insert-column-right",
  "delete-row",
  "delete-column",
  "select-table",
  "select-table-row",
  "select-table-cell",
  "select-table-column",
  "align-cell",
  "repeat-header-rows",
  "cell-shading",
  "table-style",
  "table-borders",
  "toggle-table-look",
  "merge-cells",
  "split-cell",
  "split-table",
  "autofit-contents",
  "autofit-window",
  "fixed-column-width",
  "distribute-columns",
  "cell-width",
  "cell-height",
  "text-direction",
  "convert-to-text",
  "link",
  "style",
  "add-text",
  "undo",
  "redo",
  "change-case",
  "sort",
  "multilevel-list",
  "delete-picture",
  "position-picture",
  "bring-forward",
  "send-backward",
  "wrap",
  "rotate",
  "position",
  "align-objects",
  // Review tab revision tracking (the docenTrackChanges extension).
  "track-changes",
  "accept-change",
  "reject-change",
  "previous-change",
  "next-change",
]);

// ── Pure helpers (take EditorState, return data; never touch the chain) ──

/** HeadingLevel literals the style gallery recognizes as headings. */
const HEADING_LEVEL_BY_STYLE: Readonly<Record<string, 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9>> = {
  Heading1: 1,
  Heading2: 2,
  Heading3: 3,
  Heading4: 4,
  Heading5: 5,
  Heading6: 6,
  Heading7: 7,
  Heading8: 8,
  Heading9: 9,
  Title: 1,
};

// OOXML unit scales (ECMA-376) and Word defaults — irreducible conversions.
const TWIPS_PER_INCH = 1440;
/** Word's Increase/Decrease Indent moves the left indent by 0.5". */
const INDENT_STEP_TWIPS = Math.round(0.5 * TWIPS_PER_INCH);
/** OOXML border `size` is in eighths-of-a-point; Word's default border is 0.75pt. */
const DEFAULT_BORDER = {
  style: "single",
  size: Math.round(0.75 * 8),
  color: "auto",
} as const;
const BORDER_SIDES = ["top", "bottom", "left", "right"] as const;
/** Ribbon highlight color names → OOXML ST_HighlightColor tokens ("green" in
 *  the ribbon palette is the bright green; the palette's own "Green" is the
 *  dark one). */
const HIGHLIGHT_TOKENS: Readonly<Record<string, string>> = {
  yellow: "yellow",
  "bright-green": "green",
  turquoise: "cyan",
  pink: "magenta",
  red: "red",
  green: "darkGreen",
  blue: "blue",
};

/** Encode a line-spacing multiple (1.0/1.15/1.5/2.0) as OOXML w:spacing `line`.
 *  Per ECMA-376, `lineRule="auto"` expresses `line` in 240ths of a single line
 *  (240 = 1.0, 360 = 1.5); the layout engine divides by 240 to get the
 *  multiple back. */
function lineMultipleToOoxml(mult: number): number {
  return Math.round(mult * 240);
}

/** The current selection's block node, but only if it carries the office-open
 *  paragraph attrs; null otherwise (e.g. inside a list item or table cell the
 *  block differs). */
function formattableBlock(
  state: EditorState,
): { type: string; attrs: Record<string, unknown> } | null {
  const { parent } = state.selection.$from;
  return parent.type.name === "paragraph"
    ? { type: parent.type.name, attrs: (parent.attrs ?? {}) as Record<string, unknown> }
    : null;
}

// ── Flat list helpers (a list paragraph carries bullet/numbering attrs) ──

/** A paragraph's list state: which list kind it belongs to, which marker
 *  variant ("bullet"/"circle"/… / "decimal"/"lower-alpha"/…/"source" for a
 *  round-tripped reference), its numbering reference (null for the built-in
 *  bullet sugar), and its nesting level. kind null = not a list paragraph. */
interface ListState {
  kind: "bullet" | "ordered" | null;
  variant: string;
  reference: string | null;
  level: number;
}

function listStateOf(attrs: Record<string, unknown>): ListState {
  const base = { kind: null, variant: "", reference: null, level: 0 } as ListState;
  const bullet = attrs.bullet as { level?: number } | null | undefined;
  if (bullet) return { ...base, kind: "bullet", variant: "bullet", level: bullet.level ?? 0 };
  const reference = (attrs.numbering as { reference?: string } | null | undefined)?.reference;
  if (typeof reference !== "string" || !reference) return base;
  const level = (attrs.numbering as { level?: number }).level ?? 0;
  if (reference.startsWith("docen-bullet")) {
    return {
      kind: "bullet",
      variant: reference === "docen-bullet" ? "bullet" : reference.slice("docen-bullet-".length),
      reference,
      level,
    };
  }
  const m = /^docen-ordered(?:-([a-z-]+))?-\d+$/.exec(reference);
  if (m) {
    return { kind: "ordered", variant: m[1] ?? "decimal", reference, level };
  }
  // A round-tripped reference (list_<numId>) — treated as an ordered-style
  // list so the toggles can clear it or restyle it.
  return { kind: "ordered", variant: "source", reference, level };
}

/** The Word Tab semantics patch for a paragraph: a list paragraph steps its
 *  bullet/numbering level by `delta` (clamped 0–8, keeping the numbering
 *  reference); null when the paragraph is not a list. Shared by the Tab key,
 *  the indent commands, and the list drop-downs' Change List Level. */
export function listLevelStepPatch(
  attrs: Record<string, unknown>,
  delta: number,
): Record<string, unknown> | null {
  const bullet = attrs.bullet as { level?: number } | null | undefined;
  const numbering = attrs.numbering as { reference?: string; level?: number } | null | undefined;
  if (!bullet && !numbering) return null;
  const level = Math.min(8, Math.max(0, (bullet?.level ?? numbering?.level ?? 0) + delta));
  return bullet ? { bullet: { level } } : { numbering: { ...numbering, level } };
}

/** The paragraphs the selection covers, with their positions. */
function selectedParagraphs(state: EditorState): { pos: number; node: PMNode }[] {
  const { from, to } = state.selection;
  const out: { pos: number; node: PMNode }[] = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name === "paragraph") out.push({ pos, node });
    return true;
  });
  return out;
}

/** Every numbering reference the doc's list paragraphs carry — feeds the
 *  fresh-reference allocator so a new list never collides with an existing
 *  one's numbering. */
function collectListReferences(doc: PMNode): string[] {
  const refs: string[] = [];
  doc.descendants((node) => {
    if (node.type.name !== "paragraph") return true;
    const ref = listStateOf(node.attrs as Record<string, unknown>).reference;
    if (ref) refs.push(ref);
    return false;
  });
  return refs;
}

/** Toggle the selected paragraphs' list: apply the requested kind/variant
 *  (clearing the other attr — Word's bullet/numbering mutual exclusion),
 *  keep each paragraph's nesting level, or clear the list when every selected
 *  paragraph already carries exactly that kind+variant. */
function toggleList(
  state: EditorState,
  tr: { setNodeMarkup: (pos: number, type: undefined, attrs: Record<string, unknown>) => unknown },
  kind: "bullet" | "ordered",
  variant: string,
): boolean {
  const blocks = selectedParagraphs(state);
  if (blocks.length === 0) return false;
  const active = blocks.every(({ node }) => {
    const cur = listStateOf(node.attrs as Record<string, unknown>);
    return cur.kind === kind && cur.variant === variant;
  });
  let orderedRef: string | null = null;
  if (!active && kind === "ordered") {
    orderedRef = nextOrderedReference(
      collectListReferences(state.doc),
      (state.doc.attrs as { numbering?: unknown }).numbering,
      variant === "decimal" ? undefined : variant,
    );
  }
  for (const { pos, node } of blocks) {
    const attrs = node.attrs as Record<string, unknown>;
    const level = listStateOf(attrs).level;
    if (active) {
      tr.setNodeMarkup(pos, undefined, { ...attrs, bullet: null, numbering: null });
    } else if (kind === "bullet" && variant === "bullet") {
      // The default bullet rides the built-in sugar (numId 1).
      tr.setNodeMarkup(pos, undefined, { ...attrs, bullet: { level }, numbering: null });
    } else {
      const reference =
        kind === "ordered"
          ? orderedRef!
          : `docen-bullet${variant === "bullet" ? "" : `-${variant}`}`;
      tr.setNodeMarkup(pos, undefined, {
        ...attrs,
        bullet: null,
        numbering: { reference, level },
      });
    }
  }
  return true;
}

/** Current font size at the selection (textStyle.size, in points); falls back
 *  to 11pt (Word's body default) when the selection has no explicit size. */
function currentSize(state: EditorState): number {
  const mark = state.selection.$from.marks().find((m) => m.type.name === "textStyle");
  const size = (mark?.attrs as { size?: unknown } | undefined)?.size;
  return typeof size === "number" ? size : 11;
}

/** A theme-semantic color pick: themeColor (OOXML schemeClr name), val (RGB),
 *  themeTint/themeShade (OOXML tint/shade hex). */
interface ThemeColorValue {
  themeColor: string;
  val: string;
  themeTint?: string;
  themeShade?: string;
}

function isThemeColor(value: unknown): value is ThemeColorValue {
  return typeof value === "object" && value !== null && "themeColor" in value && "val" in value;
}

/** The ShadingProperties stamp for a shading pick: null clears, a theme pick
 *  carries themeFill bindings, a bare hex stores fill. undefined = unrecognized
 *  value (command declines). */
function shadingStamp(value: unknown): Record<string, unknown> | null | undefined {
  if (value === "none") return null;
  if (isThemeColor(value)) {
    const shading: Record<string, unknown> = {
      fill: value.val,
      type: "clear",
      themeFill: value.themeColor,
    };
    if (value.themeTint) shading.themeFillTint = value.themeTint;
    if (value.themeShade) shading.themeFillShade = value.themeShade;
    return shading;
  }
  if (typeof value === "string" && value) return { fill: value, type: "clear" };
  return undefined;
}

/** Depths of the enclosing table / row / cell on the selection's `$from`
 *  path (negative = absent). The table check is also the contextual-tab
 *  signal, so it is exported for the host. */
export function tableAncestry(state: EditorState): {
  tableAt: number;
  rowAt: number;
  cellAt: number;
} | null {
  return ancestryAt(state.selection.$from);
}

/** {@link tableAncestry} for an arbitrary position — Merge Cells resolves the
 *  selection's two ends independently. */
function ancestryAt($pos: ResolvedPos): {
  tableAt: number;
  rowAt: number;
  cellAt: number;
} | null {
  const { table, tableRow, tableCell } = $pos.doc.type.schema.nodes;
  let tableAt = -1;
  let rowAt = -1;
  let cellAt = -1;
  for (let d = $pos.depth; d > 0; d -= 1) {
    const node = $pos.node(d);
    if (node.type === table && tableAt < 0) tableAt = d;
    else if (node.type === tableRow && rowAt < 0) rowAt = d;
    else if (node.type === tableCell && cellAt < 0) cellAt = d;
  }
  return tableAt < 0 ? null : { tableAt, rowAt, cellAt };
}

// ── Floating drawing helpers (the Arrange commands' shared target) ───────────

/** The selected floating drawing — a NodeSelection on a floating image (its
 *  `floating` attr set) or a wps shape (floating inside its `wpsShape`
 *  payload); the stage's hit boxes produce exactly these. Null on any other
 *  selection, so Arrange greys out through editor.can(). */
function floatingDrawingAt(
  state: EditorState,
): { pos: number; attrs: Record<string, unknown>; kind: "image" | "shape" } | null {
  const sel = state.selection;
  if (!(sel instanceof NodeSelection)) return null;
  const attrs = sel.node.attrs as Record<string, unknown>;
  if (sel.node.type.name === "image") {
    return attrs.floating ? { pos: sel.from, attrs, kind: "image" } : null;
  }
  if (sel.node.type.name === "wpsShape") {
    const shape = attrs.wpsShape as Record<string, unknown> | null;
    return shape?.floating ? { pos: sel.from, attrs, kind: "shape" } : null;
  }
  return null;
}

/** The drawing's Floating object (image: a flat attr; shape: inside the
 *  wpsShape payload). */
function floatingOf(
  target: NonNullable<ReturnType<typeof floatingDrawingAt>>,
): Record<string, unknown> {
  return (
    target.kind === "image"
      ? target.attrs.floating
      : (target.attrs.wpsShape as Record<string, unknown>).floating
  ) as Record<string, unknown>;
}

/** Write a Floating back onto the drawing, shallow-copying the carrier the
 *  way PM immutability requires (image: flat; shape: the wpsShape payload). */
function withFloating(
  target: NonNullable<ReturnType<typeof floatingDrawingAt>>,
  floating: Record<string, unknown>,
): Record<string, unknown> {
  return target.kind === "image"
    ? { ...target.attrs, floating }
    : {
        ...target.attrs,
        wpsShape: { ...(target.attrs.wpsShape as Record<string, unknown>), floating },
      };
}

/** Stamp the next Floating onto the drawing (one markup write, no scroll —
 *  Arrange edits never move the caret). The markup write replaces the node,
 *  which collapses a NodeSelection to a caret — restoring it keeps the
 *  drawing selected so the command can repeat (Word's Bring Forward chains). */
function stampFloating(
  tr: Transaction,
  target: NonNullable<ReturnType<typeof floatingDrawingAt>>,
  floating: Record<string, unknown>,
): boolean {
  return stampAttrs(tr, target, withFloating(target, floating));
}

/** {@link stampFloating} for a full attrs object (rotate rewrites the image's
 *  top level or the shape's payload, not just the Floating). */
function stampAttrs(
  tr: Transaction,
  target: NonNullable<ReturnType<typeof floatingDrawingAt>>,
  attrs: Record<string, unknown>,
): boolean {
  tr.setNodeMarkup(target.pos, undefined, attrs);
  tr.setSelection(NodeSelection.create(tr.doc, target.pos));
  return true;
}

/** The 9-grid cell alignment: vertical half → the cell's verticalAlign, the
 *  horizontal half → every paragraph's alignment in the cell. */
const CELL_ALIGN: Record<string, { v: string; h: string }> = {
  tl: { v: "top", h: "left" },
  tc: { v: "top", h: "center" },
  tr: { v: "top", h: "right" },
  ml: { v: "center", h: "left" },
  mc: { v: "center", h: "center" },
  mr: { v: "center", h: "right" },
  bl: { v: "bottom", h: "left" },
  bc: { v: "bottom", h: "center" },
  br: { v: "bottom", h: "right" },
};

/** The Position gallery's nine cells → margin-relative align tokens (same
 *  key space as {@link CELL_ALIGN}; the ST_PositionAlign vocabulary both
 *  axes resolve through). */
const POSITION_ALIGN: Record<string, { v: string; h: string }> = {
  tl: { v: "top", h: "left" },
  tc: { v: "top", h: "center" },
  tr: { v: "top", h: "right" },
  ml: { v: "center", h: "left" },
  mc: { v: "center", h: "center" },
  mr: { v: "center", h: "right" },
  bl: { v: "bottom", h: "left" },
  bc: { v: "bottom", h: "center" },
  br: { v: "bottom", h: "right" },
};

/** Word's Add Text menu: a TOC level → the heading pStyle it stamps (the TOC
 *  field collects Heading 1-3), "none" returning paragraphs to body text. */
const ADD_TEXT_LEVELS: Readonly<Record<string, string | null>> = {
  "level-1": "Heading1",
  "level-2": "Heading2",
  "level-3": "Heading3",
  none: null,
};

// ── Cell Size / AutoFit measurement helpers ──────────────────────────────────

/** A UniversalMeasure string ("1.5cm") or bare number string → twips; number
 *  passes through as twips already. Mirrors the engine's UM table
 *  (docx/src/layout/project/guards.ts measureTwip) — the value spaces are the
 *  office-open length fields. */
const MEASURE_TWIP_UNITS: ReadonlyArray<readonly [string, number]> = [
  ["pt", 20],
  ["pc", 240],
  ["in", 1440],
  ["mm", 1440 / 25.4],
  ["cm", 1440 / 2.54],
  ["px", 15],
];
function parseMeasureTwip(v: unknown): number | null {
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v !== "string") return null;
  const bare = Number(v);
  if (Number.isFinite(bare)) return bare;
  const m = /^(-?[\d.]+)\s*(pt|pc|in|mm|cm|px)$/.exec(v.trim());
  if (!m) return null;
  const unit = MEASURE_TWIP_UNITS.find(([u]) => u === m[2]);
  return unit ? Number(m[1]) * unit[1] : null;
}

const CJK_CHAR = /[⺀-鿿豈-﫿！-｠　-〿]/;

/** Content-width heuristic for AutoFit Contents: no text measurer runs in the
 *  command layer, so a column's width comes from its widest cell's character
 *  count (a CJK glyph ≈ one 12pt em = 240 twips, Latin ≈ half) plus inset
 *  slack. Honest sizing for text cells; images/wide objects overflow. */
function measureTextTwip(text: string): number {
  let tw = 0;
  for (const ch of text) tw += CJK_CHAR.test(ch) ? 240 : 110;
  return tw + 120;
}

/** Word's smallest usable column — 0.5" — also the AutoFit floor. */
const MIN_COL_TWIP = 720;

type TableBordersLike = Record<string, { style: string; size: number; color: string } | undefined>;
const GRID_BORDER = { style: "single", size: 4, color: "auto" };
const NO_BORDER = { style: "none", size: 0, color: "auto" };

/** A Table Styles gallery preset: the border set plus the conditional fills —
 *  the header-row shading and the alternating body-row band. Word renders
 *  those through the table style; the editor has no style engine, so applying
 *  a preset bakes the fills onto the cells directly. */
export interface TableStylePreset {
  borders: TableBordersLike | null;
  /** Shading stamped on every tblHeader row's cells (Word's header-row
   *  conditional formatting). */
  headerFill?: string;
  /** Shading stamped on alternating body rows (Word's banded-rows
   *  conditional formatting). */
  bandFill?: string;
}

const TABLE_GRID_BORDERS: TableBordersLike = {
  top: GRID_BORDER,
  bottom: GRID_BORDER,
  left: GRID_BORDER,
  right: GRID_BORDER,
  insideHorizontal: GRID_BORDER,
  insideVertical: GRID_BORDER,
};
const TABLE_NO_BORDERS: TableBordersLike = {
  top: NO_BORDER,
  bottom: NO_BORDER,
  left: NO_BORDER,
  right: NO_BORDER,
  insideHorizontal: NO_BORDER,
  insideVertical: NO_BORDER,
};

/** Word's Table Styles gallery stand-ins, named after the built-ins they
 *  approximate (Accent 1 colors — the Office default theme). */
export const TABLE_STYLE_PRESETS: Record<string, TableStylePreset> = {
  "no-style-no-grid": { borders: TABLE_NO_BORDERS },
  "table-grid": { borders: TABLE_GRID_BORDERS },
  // Horizontal rules only, with a light band on alternating body rows.
  "light-shading": {
    borders: { top: GRID_BORDER, bottom: GRID_BORDER, insideHorizontal: GRID_BORDER },
    bandFill: "D9E2F3",
  },
  // Horizontal rules + a tinted header row.
  "light-list": {
    borders: { top: GRID_BORDER, bottom: GRID_BORDER, insideHorizontal: GRID_BORDER },
    headerFill: "8EAADB",
  },
  // Full grid + a tinted header row.
  "light-grid": { borders: TABLE_GRID_BORDERS, headerFill: "D9E2F3" },
  // Heavier outside frame + the dark Accent-1 header.
  "grid-table": {
    borders: {
      top: { style: "single", size: 8, color: "auto" },
      bottom: { style: "single", size: 8, color: "auto" },
      left: { style: "single", size: 8, color: "auto" },
      right: { style: "single", size: 8, color: "auto" },
      insideHorizontal: GRID_BORDER,
      insideVertical: GRID_BORDER,
    },
    headerFill: "4472C4",
  },
};

/** Border-side stamps for the Layout/Design borders dropdown — value matches
 *  the Home border menu (none/bottom/top/left/right/all/outside). */
function tableBordersStamp(
  value: string,
  current: TableBordersLike | null,
): TableBordersLike | null {
  if (value === "none") return TABLE_STYLE_PRESETS["no-style-no-grid"]!.borders;
  const borders: TableBordersLike = { ...current };
  if (value === "all" || value === "outside") {
    borders.top = GRID_BORDER;
    borders.bottom = GRID_BORDER;
    borders.left = GRID_BORDER;
    borders.right = GRID_BORDER;
  }
  if (value === "all") {
    borders.insideHorizontal = GRID_BORDER;
    borders.insideVertical = GRID_BORDER;
  }
  if (value === "bottom" || value === "top" || value === "left" || value === "right") {
    borders[value] = GRID_BORDER;
  }
  return borders;
}

/** Delete the table at `pos` (size `size`) and park the caret where it stood
 *  — shared by delete-table and the collapse cases of delete-row/-column. */
function deleteTableAt(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  pos: number,
  size: number,
): boolean {
  if (!dispatch) return true;
  const tr = state.tr.delete(pos, pos + size);
  tr.setSelection(TextSelection.near(tr.doc.resolve(pos)));
  dispatch(tr.scrollIntoView());
  return true;
}

/** Stamp a borders preset on the enclosing table. */
function stampTableBorders(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  borders: TableBordersLike | null,
): boolean {
  if (!borders) return false;
  const anchor = tableAncestry(state);
  if (!anchor) return false;
  if (dispatch) {
    const { $from } = state.selection;
    const table = $from.node(anchor.tableAt);
    dispatch(
      state.tr
        .setNodeMarkup($from.before(anchor.tableAt), undefined, { ...table.attrs, borders })
        .scrollIntoView(),
    );
  }
  return true;
}

/** Transform text per Word's Change Case modes. CJK sentence terminators
 *  (。！？) honoured alongside ASCII .!?. */
function transformCase(text: string, mode?: string): string {
  switch (mode) {
    case "lower":
      return text.toLowerCase();
    case "upper":
      return text.toUpperCase();
    case "capitalize":
      return text.replace(/\p{L}[\p{L}'-]*/gu, (w) => w.charAt(0).toUpperCase() + w.slice(1));
    case "toggle":
      return text.replace(/\p{L}/gu, (c) =>
        c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase(),
      );
    case "sentence":
    default:
      return text.replace(/(^\s*\p{L})|([.!?。！？]\s*\p{L})/gu, (m) => m.toUpperCase());
  }
}

// ── The extension ───────────────────────────────────────────────────────────

export const DocumentCommands = Extension.create({
  name: "documentCommands",
  addCommands() {
    return {
      // ── Font marks — wrap the built-in Tiptap toggles ──
      bold:
        () =>
        ({ commands }) =>
          commands.toggleMark("bold"),
      italic:
        () =>
        ({ commands }) =>
          commands.toggleMark("italic"),
      underline:
        () =>
        ({ commands }) =>
          commands.toggleMark("underline"),
      strike:
        () =>
        ({ commands }) =>
          commands.toggleMark("strike"),
      subscript:
        () =>
        ({ commands }) =>
          commands.toggleMark("subscript"),
      superscript:
        () =>
        ({ commands }) =>
          commands.toggleMark("superscript"),
      highlight:
        (value) =>
        ({ commands }) => {
          // "none" clears; a palette color sets its token; no value (the split
          // button's main click) applies Word's default yellow.
          if (value === "none") return commands.unsetMark("highlight");
          const token = HIGHLIGHT_TOKENS[value ?? ""] ?? "yellow";
          return commands.setMark("highlight", { color: token });
        },
      code:
        () =>
        ({ commands }) =>
          commands.toggleMark("code"),
      "clear-format":
        () =>
        ({ chain }) =>
          chain().unsetAllMarks().clearNodes().run(),
      // Font family / size — applied as textStyle mark attrs (`font` = name,
      // `size` = points). grow/shrink step the current size by 2pt.
      "font-name":
        (font) =>
        ({ commands }) =>
          commands.setMark("textStyle", { font: font ?? null }),
      "font-size":
        (size) =>
        ({ commands }) =>
          commands.setMark("textStyle", { size: size ? Number(size) : null }),
      "grow-font":
        () =>
        ({ state, commands }) =>
          commands.setMark("textStyle", { size: currentSize(state) + 2 }),
      "shrink-font":
        () =>
        ({ state, commands }) =>
          commands.setMark("textStyle", { size: Math.max(1, currentSize(state) - 2) }),

      // ── Paragraph / alignment ──
      "align-left":
        () =>
        ({ commands }) =>
          commands.updateAttributes("paragraph", { alignment: "left" }),
      "align-center":
        () =>
        ({ commands }) =>
          commands.updateAttributes("paragraph", { alignment: "center" }),
      "align-right":
        () =>
        ({ commands }) =>
          commands.updateAttributes("paragraph", { alignment: "right" }),
      justify:
        () =>
        ({ commands }) =>
          commands.updateAttributes("paragraph", { alignment: "both" }),

      // ── Indent / spacing / shading / border — stamp office-open block attrs ──
      // All four walk EVERY selected paragraph (each keeps its own existing
      // attrs — a range-spanning updateAttributes would stamp the first
      // paragraph's merged value onto the rest). Increase/decrease left
      // indent by Word's 0.5" step; a list paragraph indents to the next
      // outline level instead (the Tab semantics, not a text shift).
      "indent-increase":
        () =>
        ({ state, tr }) => {
          let touched = false;
          for (const { pos, node } of selectedParagraphs(state)) {
            const attrs = node.attrs as Record<string, unknown>;
            const list = listLevelStepPatch(attrs, 1);
            if (list) {
              tr.setNodeMarkup(pos, undefined, { ...attrs, ...list });
              touched = true;
              continue;
            }
            const current = (attrs.indent ?? {}) as { left?: number; right?: number };
            const left = Math.max(0, (current.left ?? 0) + INDENT_STEP_TWIPS);
            tr.setNodeMarkup(pos, undefined, { ...attrs, indent: { ...current, left } });
            touched = true;
          }
          return touched;
        },
      "indent-decrease":
        () =>
        ({ state, tr }) => {
          let touched = false;
          for (const { pos, node } of selectedParagraphs(state)) {
            const attrs = node.attrs as Record<string, unknown>;
            const list = listLevelStepPatch(attrs, -1);
            if (list) {
              tr.setNodeMarkup(pos, undefined, { ...attrs, ...list });
              touched = true;
              continue;
            }
            const current = (attrs.indent ?? {}) as { left?: number; right?: number };
            const left = Math.max(0, (current.left ?? 0) - INDENT_STEP_TWIPS);
            tr.setNodeMarkup(pos, undefined, { ...attrs, indent: { ...current, left } });
            touched = true;
          }
          return touched;
        },
      // Line spacing as a multiple of single (1.0/1.15/1.5/2.0); preserves
      // existing before/after. The split's main click carries no value — it
      // applies single spacing (Word's default). The dropdown's trailing
      // entries are Word's "Add Space Before/After Paragraph": 10pt (200
      // twips), not a multiple.
      "line-spacing":
        (mult) =>
        ({ state, tr }) => {
          const blocks = selectedParagraphs(state);
          if (!blocks.length) return false;
          if (mult === "add-before" || mult === "add-after") {
            const key = mult === "add-before" ? "before" : "after";
            for (const { pos, node } of blocks) {
              const attrs = node.attrs as Record<string, unknown>;
              const current = (attrs.spacing ?? {}) as Record<string, unknown>;
              tr.setNodeMarkup(pos, undefined, {
                ...attrs,
                spacing: { ...current, [key]: 200 },
              });
            }
            return true;
          }
          const parsed = parseFloat(mult ?? "");
          const m = Number.isFinite(parsed) ? parsed : 1;
          for (const { pos, node } of blocks) {
            const attrs = node.attrs as Record<string, unknown>;
            const current = (attrs.spacing ?? {}) as Record<string, unknown>;
            tr.setNodeMarkup(pos, undefined, {
              ...attrs,
              spacing: { ...current, line: lineMultipleToOoxml(m), lineRule: "auto" },
            });
          }
          return true;
        },
      // Shading follows Word's selection split: a text selection paints only
      // the selected runs (character shading via the textStyle mark); a bare
      // cursor paints the whole paragraph.
      shading:
        (value) =>
        ({ state, commands, tr }) => {
          const stamp = shadingStamp(value);
          if (stamp === undefined) return false;
          if (!state.selection.empty) {
            return commands.setMark("textStyle", { shading: stamp });
          }
          const blocks = selectedParagraphs(state);
          if (!blocks.length) return false;
          for (const { pos, node } of blocks) {
            const attrs = node.attrs as Record<string, unknown>;
            tr.setNodeMarkup(pos, undefined, { ...attrs, shading: stamp });
          }
          return true;
        },
      // Run font color: "none" clears; a theme pick stores a ColorOptions
      // (theme-bound); a bare hex stores the color directly.
      "font-color":
        (value) =>
        ({ commands }) => {
          if (value === "none") return commands.setMark("textStyle", { color: null });
          if (isThemeColor(value) || (typeof value === "string" && value)) {
            return commands.setMark("textStyle", { color: value });
          }
          return false;
        },
      // Paragraph borders: value picks sides (bottom/top/left/right/all/outside);
      // "none" clears all. Merges with each paragraph's existing so other sides
      // stay. Default single 0.75pt, "auto" color (Word default).
      border:
        (side) =>
        ({ state, tr }) => {
          const blocks = selectedParagraphs(state);
          if (!blocks.length) return false;
          // The split button's main click carries no value — default bottom.
          const s = side ?? "bottom";
          if (s === "none") {
            for (const { pos, node } of blocks) {
              const attrs = node.attrs as Record<string, unknown>;
              tr.setNodeMarkup(pos, undefined, { ...attrs, border: null });
            }
            return true;
          }
          const sides =
            s === "all" || s === "outside"
              ? BORDER_SIDES
              : (BORDER_SIDES as readonly string[]).includes(s)
                ? [s]
                : null;
          if (!sides) return false;
          for (const { pos, node } of blocks) {
            const attrs = node.attrs as Record<string, unknown>;
            const current = (attrs.border ?? {}) as Record<string, unknown>;
            const border = { ...current };
            for (const side of sides) border[side] = { ...DEFAULT_BORDER };
            tr.setNodeMarkup(pos, undefined, { ...attrs, border });
          }
          return true;
        },

      // ── Lists / blocks ──
      // Flat list toggles: stamp/clear the selected paragraphs' list attrs.
      // The ribbon dropdown's variant picks the marker (●/○/■, decimal/alpha/
      // roman); clicking the current variant clears the list (Word).
      "bullet-list":
        (variant) =>
        ({ state, tr }) =>
          toggleList(state, tr, "bullet", variant && BULLET_GLYPHS[variant] ? variant : "bullet"),
      "ordered-list":
        (variant) =>
        ({ state, tr }) =>
          toggleList(
            state,
            tr,
            "ordered",
            variant && ORDERED_FORMATS[variant] ? variant : "decimal",
          ),
      // Quote: stamp/clear Word's built-in IntenseQuote paragraph style (a
      // blockquote is a styled paragraph in OOXML, not a wrapper node).
      blockquote:
        () =>
        ({ state, chain }) => {
          const block = formattableBlock(state);
          if (!block) return false;
          const quoted = block.attrs.style === "IntenseQuote";
          return chain()
            .updateAttributes(block.type, { style: quoted ? null : "IntenseQuote" })
            .run();
        },
      // OOXML has no HR element — a horizontal rule is a thematic-break
      // paragraph (rendered with a bottom border).
      "horizontal-rule":
        () =>
        ({ chain }) =>
          chain()
            .insertContent({ type: "paragraph", attrs: { thematicBreak: true } })
            .run(),
      // setPageBreak splits the paragraph so the paginator reflows the tail.
      "page-break":
        () =>
        ({ commands }) =>
          commands.setPageBreak(),
      "column-break":
        () =>
        ({ commands }) =>
          commands.setColumnBreak(),
      "section-break":
        () =>
        ({ commands }) =>
          commands.setSectionBreak(),
      // Insert a 3×3 table (Word's default Insert > Table preset). The header
      // row is the row-level tblHeader attr (w:tblHeader) — no header-cell
      // node type exists; every cell is a plain tableCell. Borders stamp
      // Word's "Table Grid" — 0.5pt single lines everywhere (w:sz is eighths
      // of a point, 4 = 0.5pt) — so the table is visible without a TableGrid
      // style in the document's styles.xml.
      "insert-table":
        () =>
        ({ state, dispatch }) => {
          const { table, tableRow, tableCell, paragraph } = state.schema.nodes;
          const cell = tableCell.createAndFill(null, [paragraph.create()]);
          if (!cell) return false;
          // The shape is structurally valid by construction (3 cells in a
          // "tableCell+" row), so the fill can only fail on a schema drift.
          const mkRow = (header: boolean): PMNode =>
            tableRow.createAndFill(header ? { tableHeader: true } : null, [cell, cell, cell])!;
          const node = table.createAndFill(
            {
              borders: {
                top: GRID_BORDER,
                bottom: GRID_BORDER,
                left: GRID_BORDER,
                right: GRID_BORDER,
                insideHorizontal: GRID_BORDER,
                insideVertical: GRID_BORDER,
              },
            },
            [mkRow(true), mkRow(false), mkRow(false)],
          );
          if (!node) return false;
          if (dispatch) {
            const pos = state.selection.from;
            const tr = state.tr.replaceSelectionWith(node);
            // Caret lands in the first cell, ready to type (Word behavior).
            tr.setSelection(TextSelection.near(tr.doc.resolve(pos + 2)));
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Delete the enclosing table (Word's right-click "Delete Table"). The
      // nearest ancestor table wins, so a table nested in a cell deletes
      // only itself.
      "delete-table":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          const { $from } = state.selection;
          return deleteTableAt(
            state,
            dispatch,
            $from.before(anchor.tableAt),
            $from.node(anchor.tableAt).nodeSize,
          );
        },
      // ── Table context commands (Word's Table Design / Layout tabs) ──

      // Insert a row copying the current one (formatting follows, like Word).
      "insert-row-above":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.rowAt < 0) return false;
          if (dispatch) {
            const { $from } = state.selection;
            const row = $from.node(anchor.rowAt);
            dispatch(
              state.tr.insert($from.before(anchor.rowAt), row.copy(row.content)).scrollIntoView(),
            );
          }
          return true;
        },
      "insert-row-below":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.rowAt < 0) return false;
          if (dispatch) {
            const { $from } = state.selection;
            const row = $from.node(anchor.rowAt);
            dispatch(
              state.tr.insert($from.after(anchor.rowAt), row.copy(row.content)).scrollIntoView(),
            );
          }
          return true;
        },
      // One cell per row, copied from each row's cell at the current column
      // index (rows may index differently once spans exist — span-aware grid
      // math is logged for a later batch). Bottom-up keeps positions valid as
      // earlier edits shift later ones.
      "insert-column-right":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.rowAt < 0) return false;
          if (dispatch) {
            const { $from } = state.selection;
            const tableNode = $from.node(anchor.tableAt);
            const tablePos = $from.before(anchor.tableAt);
            const cellIndex = $from.index(anchor.rowAt);
            const tr = state.tr;
            for (let r = tableNode.childCount - 1; r >= 0; r -= 1) {
              const rowNode = tableNode.child(r);
              let rowPos = tablePos + 1;
              for (let i = 0; i < r; i += 1) rowPos += tableNode.child(i).nodeSize;
              const idx = Math.min(cellIndex, rowNode.childCount - 1);
              let cellPos = rowPos + 1;
              for (let c = 0; c <= idx; c += 1) cellPos += rowNode.child(c).nodeSize;
              tr.insert(cellPos, rowNode.child(idx));
            }
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      "insert-column-left":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.rowAt < 0) return false;
          if (dispatch) {
            const { $from } = state.selection;
            const tableNode = $from.node(anchor.tableAt);
            const tablePos = $from.before(anchor.tableAt);
            const cellIndex = $from.index(anchor.rowAt);
            const tr = state.tr;
            for (let r = tableNode.childCount - 1; r >= 0; r -= 1) {
              const rowNode = tableNode.child(r);
              let rowPos = tablePos + 1;
              for (let i = 0; i < r; i += 1) rowPos += tableNode.child(i).nodeSize;
              const idx = Math.min(cellIndex, rowNode.childCount - 1);
              let cellPos = rowPos + 1;
              for (let c = 0; c < idx; c += 1) cellPos += rowNode.child(c).nodeSize;
              tr.insert(cellPos, rowNode.child(idx));
            }
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Deleting the last row/column deletes the whole table (Word behavior).
      "delete-row":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.rowAt < 0) return false;
          const { $from } = state.selection;
          const tableNode = $from.node(anchor.tableAt);
          if (tableNode.childCount === 1) {
            return deleteTableAt(state, dispatch, $from.before(anchor.tableAt), tableNode.nodeSize);
          }
          if (dispatch) {
            const rowPos = $from.before(anchor.rowAt);
            const row = $from.node(anchor.rowAt);
            dispatch(state.tr.delete(rowPos, rowPos + row.nodeSize).scrollIntoView());
          }
          return true;
        },
      "delete-column":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.rowAt < 0) return false;
          const { $from } = state.selection;
          const tableNode = $from.node(anchor.tableAt);
          const cellIndex = $from.index(anchor.rowAt);
          const minCells = Math.min(
            ...Array.from(
              { length: tableNode.childCount },
              (_, r) => tableNode.child(r).childCount,
            ),
          );
          if (minCells === 1) {
            return deleteTableAt(state, dispatch, $from.before(anchor.tableAt), tableNode.nodeSize);
          }
          if (dispatch) {
            const tablePos = $from.before(anchor.tableAt);
            const tr = state.tr;
            for (let r = tableNode.childCount - 1; r >= 0; r -= 1) {
              const rowNode = tableNode.child(r);
              let rowPos = tablePos + 1;
              for (let i = 0; i < r; i += 1) rowPos += tableNode.child(i).nodeSize;
              const idx = Math.min(cellIndex, rowNode.childCount - 1);
              let cellPos = rowPos + 1;
              for (let c = 0; c < idx; c += 1) cellPos += rowNode.child(c).nodeSize;
              tr.delete(cellPos, cellPos + rowNode.child(idx).nodeSize);
            }
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      "select-table":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          if (dispatch) {
            const pos = state.selection.$from.before(anchor.tableAt);
            dispatch(state.tr.setSelection(NodeSelection.create(state.doc, pos)).scrollIntoView());
          }
          return true;
        },
      "select-table-row":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.rowAt < 0) return false;
          if (dispatch) {
            const pos = state.selection.$from.before(anchor.rowAt);
            dispatch(state.tr.setSelection(NodeSelection.create(state.doc, pos)).scrollIntoView());
          }
          return true;
        },
      "select-table-cell":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.cellAt < 0) return false;
          if (dispatch) {
            const { $from } = state.selection;
            const cellPos = $from.before(anchor.cellAt);
            const cell = $from.node(anchor.cellAt);
            dispatch(
              state.tr
                .setSelection(
                  TextSelection.create(state.doc, cellPos + 1, cellPos + cell.nodeSize - 1),
                )
                .scrollIntoView(),
            );
          }
          return true;
        },
      // The caret's column across all rows: a TextSelection from the first
      // row's cell content to the last row's (the same index-per-row fallback
      // as insert-column — span-aware grid math is logged for a later batch).
      "select-table-column":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.rowAt < 0) return false;
          if (dispatch) {
            const { $from } = state.selection;
            const tableNode = $from.node(anchor.tableAt);
            const tablePos = $from.before(anchor.tableAt);
            const cellIndex = $from.index(anchor.rowAt);
            let first = -1;
            let lastEnd = -1;
            for (let r = 0; r < tableNode.childCount; r += 1) {
              const rowNode = tableNode.child(r);
              let rowPos = tablePos + 1;
              for (let i = 0; i < r; i += 1) rowPos += tableNode.child(i).nodeSize;
              const idx = Math.min(cellIndex, rowNode.childCount - 1);
              let cellPos = rowPos + 1;
              for (let c = 0; c < idx; c += 1) cellPos += rowNode.child(c).nodeSize;
              if (first < 0) first = cellPos + 1;
              lastEnd = cellPos + rowNode.child(idx).nodeSize - 1;
            }
            dispatch(
              state.tr
                .setSelection(TextSelection.create(state.doc, first, lastEnd))
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Word's 9-grid: the vertical half lands on the cell (verticalAlign),
      // the horizontal half on every paragraph in the cell (alignment).
      "align-cell":
        (value) =>
        ({ state, dispatch }) => {
          const spec = CELL_ALIGN[value ?? ""];
          if (!spec) return false;
          const anchor = tableAncestry(state);
          if (!anchor || anchor.cellAt < 0) return false;
          if (dispatch) {
            const { $from } = state.selection;
            const cellPos = $from.before(anchor.cellAt);
            const cell = $from.node(anchor.cellAt);
            const from = cellPos + 1;
            const to = cellPos + cell.nodeSize - 1;
            const { paragraph } = state.schema.nodes;
            const tr = state.tr.setNodeMarkup(cellPos, undefined, {
              ...cell.attrs,
              verticalAlign: spec.v,
            });
            state.doc.nodesBetween(from, to, (node, pos) => {
              if (node.type === paragraph && node.attrs.alignment !== spec.h) {
                tr.setNodeMarkup(pos, undefined, { ...node.attrs, alignment: spec.h });
              }
              return true;
            });
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Word's Repeat Header Rows — toggles the current row's tblHeader.
      "repeat-header-rows":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.rowAt < 0) return false;
          if (dispatch) {
            const { $from } = state.selection;
            const row = $from.node(anchor.rowAt);
            dispatch(
              state.tr
                .setNodeMarkup($from.before(anchor.rowAt), undefined, {
                  ...row.attrs,
                  tableHeader: !row.attrs.tableHeader,
                })
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Cell-level shading (tcPr shd) — the Home shading button stays at
      // paragraph level; Word's Table Design shading is the cell property.
      "cell-shading":
        (value) =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.cellAt < 0) return false;
          const stamp = shadingStamp(value);
          if (stamp === undefined) return false;
          if (dispatch) {
            const { $from } = state.selection;
            const cellPos = $from.before(anchor.cellAt);
            const cell = $from.node(anchor.cellAt);
            dispatch(
              state.tr
                .setNodeMarkup(cellPos, undefined, { ...cell.attrs, shading: stamp })
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Apply a Table Styles gallery preset: the border set on the table plus
      // the conditional fills baked onto the cells. Every cell's shading is
      // rewritten (fill or null), so switching presets never leaves the
      // previous style's bands behind.
      "table-style":
        (value) =>
        ({ state, dispatch }) => {
          const preset = value ? TABLE_STYLE_PRESETS[value] : undefined;
          if (!preset) return false;
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          if (dispatch) {
            const { $from } = state.selection;
            const tablePos = $from.before(anchor.tableAt);
            const tableNode = $from.node(anchor.tableAt);
            const tr = state.tr.setNodeMarkup(tablePos, undefined, {
              ...tableNode.attrs,
              borders: preset.borders,
            });
            for (let r = 0; r < tableNode.childCount; r += 1) {
              const rowNode = tableNode.child(r);
              let rowPos = tablePos + 1;
              for (let i = 0; i < r; i += 1) rowPos += tableNode.child(i).nodeSize;
              const isHeader = !!rowNode.attrs.tableHeader;
              // Word bands the odd body rows (1st, 3rd, …); with a header at
              // r=0 those are the even table indices ≥ 2.
              const isBand = !isHeader && preset.bandFill != null && r >= 2 && r % 2 === 0;
              const fill = isHeader ? preset.headerFill : isBand ? preset.bandFill : undefined;
              rowNode.forEach((cell: PMNode, offset: number) => {
                const cellPos = rowPos + 1 + offset;
                tr.setNodeMarkup(cellPos, undefined, {
                  ...cell.attrs,
                  shading: fill ? { fill, type: "clear" } : null,
                });
              });
            }
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Toggle one of the Table Style Options flags (Word's Header Row /
      // Total Row / Banded Rows / … checkboxes) — the table's tblLook.
      "toggle-table-look":
        (value) =>
        ({ state, dispatch }) => {
          const flags = ["firstRow", "lastRow", "firstCol", "lastCol", "bandRow", "bandCol"];
          if (typeof value !== "string" || !flags.includes(value)) return false;
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          if (dispatch) {
            const { $from } = state.selection;
            const tablePos = $from.before(anchor.tableAt);
            const table = $from.node(anchor.tableAt);
            const look = {
              ...((table.attrs.tableLook ?? {}) as Record<string, boolean>),
            };
            look[value] = !look[value];
            dispatch(
              state.tr
                .setNodeMarkup(tablePos, undefined, { ...table.attrs, tableLook: look })
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Border-side presets on the table (value space matches the Home
      // paragraph-border menu).
      "table-borders":
        (value) =>
        ({ state, dispatch }) => {
          if (typeof value !== "string") return false;
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          const current = (state.selection.$from.node(anchor.tableAt).attrs.borders ??
            null) as TableBordersLike | null;
          return stampTableBorders(state, dispatch, tableBordersStamp(value, current));
        },
      // Word's Text Direction button: toggles the cell's tcPr textDirection
      // (tbRl ↔ unset). The attr round-trips through the docx engine; the
      // canvas doesn't paint vertical cell text yet.
      "text-direction":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.cellAt < 0) return false;
          if (dispatch) {
            const { $from } = state.selection;
            const cellPos = $from.before(anchor.cellAt);
            const cell = $from.node(anchor.cellAt);
            const next = cell.attrs.textDirection ? null : "tbRl";
            dispatch(
              state.tr
                .setNodeMarkup(cellPos, undefined, { ...cell.attrs, textDirection: next })
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Word's "Convert to Text": each row becomes a paragraph, cells joined
      // by tabs (Word's default separator); the caret lands where the table
      // stood. A cell's multi-paragraph content collapses to its text — Word
      // keeps the paragraphs, our join is the honest simple form.
      "convert-to-text":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          if (dispatch) {
            const { $from } = state.selection;
            const tablePos = $from.before(anchor.tableAt);
            const tableNode = $from.node(anchor.tableAt);
            const { paragraph } = state.schema.nodes;
            const paras: PMNode[] = [];
            for (let r = 0; r < tableNode.childCount; r += 1) {
              const texts: string[] = [];
              tableNode.child(r).forEach((cell: PMNode) => texts.push(cell.textContent));
              const text = texts.join("\t");
              paras.push(
                text ? paragraph.create(null, state.schema.text(text)) : paragraph.create(),
              );
            }
            const tr = state.tr.replaceWith(tablePos, tablePos + tableNode.nodeSize, paras);
            tr.setSelection(TextSelection.near(tr.doc.resolve(tablePos + 1)));
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Word's Merge Cells over the selection's bounding rectangle. Each
      // spanned row folds its cells into one: the row's first cell takes
      // columnSpan = width, rows below the first take verticalMerge
      // "continue" (their content stays put — the layout folds continue
      // cells into the restart cell, so nothing is lost). Grid math is
      // cellIndex-approximate, so a span-mismatched row is left untouched.
      "merge-cells":
        () =>
        ({ state, dispatch }) => {
          const fromA = ancestryAt(state.selection.$from);
          const toA = ancestryAt(state.selection.$to);
          if (!fromA || !toA || fromA.rowAt < 0 || toA.rowAt < 0) return false;
          const { $from, $to } = state.selection;
          if ($from.before(fromA.tableAt) !== $to.before(toA.tableAt)) return false;
          const tableNode = $from.node(fromA.tableAt);
          const grid = (tableNode.attrs.columnWidths as number[] | null)?.length ?? 0;
          const c1 = $from.index(fromA.rowAt);
          const c2 = $to.index(toA.rowAt);
          if (fromA.rowAt === toA.rowAt && c1 === c2) return false;
          if (dispatch) {
            const tablePos = $from.before(fromA.tableAt);
            // Row indices into the table's children — the ancestry depths are
            // not indexes (a depth-2 rowAt would address the last row).
            const rowFrom = $from.index(fromA.tableAt);
            const rowTo = $to.index(toA.tableAt);
            const tr = state.tr;
            for (let r = rowTo; r >= rowFrom; r -= 1) {
              const rowNode = tableNode.child(r);
              // Bottom-up keeps positions valid as earlier deletions shift
              // later ones; a row that doesn't match the grid exactly (a
              // previously merged one) is skipped rather than corrupted.
              if (grid > 0 && rowNode.childCount !== grid) continue;
              let rowPos = tablePos + 1;
              for (let i = 0; i < r; i += 1) rowPos += tableNode.child(i).nodeSize;
              const last = Math.min(c2, rowNode.childCount - 1);
              if (c1 > last) continue;
              let basePos = rowPos + 1;
              for (let c = 0; c < c1; c += 1) basePos += rowNode.child(c).nodeSize;
              const base = rowNode.child(c1);
              tr.setNodeMarkup(basePos, undefined, {
                ...base.attrs,
                columnSpan: last > c1 ? last - c1 + 1 : null,
                verticalMerge: r > rowFrom ? "continue" : base.attrs.verticalMerge,
              });
              for (let c = last; c > c1; c -= 1) {
                let cellPos = rowPos + 1;
                for (let cc = 0; cc < c; cc += 1) cellPos += rowNode.child(cc).nodeSize;
                tr.delete(cellPos, cellPos + rowNode.child(c).nodeSize);
              }
            }
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Word's Split Cells without the dialog: a merged cell (columnSpan or
      // verticalMerge) returns to its own single grid cell, empty twins
      // taking the spanned columns. The dialog's rows×cols form is not built.
      "split-cell":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.cellAt < 0) return false;
          const { $from } = state.selection;
          const cell = $from.node(anchor.cellAt);
          const span = (cell.attrs.columnSpan as number | null) ?? 1;
          if (span < 2 && !cell.attrs.verticalMerge) return false;
          if (dispatch) {
            const cellPos = $from.before(anchor.cellAt);
            const tr = state.tr.setNodeMarkup(cellPos, undefined, {
              ...cell.attrs,
              columnSpan: null,
              verticalMerge: null,
            });
            const blank = cell.type.create(null, state.schema.nodes.paragraph.create());
            for (let i = 0; i < span - 1; i += 1) {
              tr.insert(cellPos + cell.nodeSize, blank);
            }
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Word's Split Table: the caret's row starts a second table with the
      // same formatting (attrs are shared — borders, grid, style).
      "split-table":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.rowAt < 0) return false;
          const { $from } = state.selection;
          const tableNode = $from.node(anchor.tableAt);
          const rowIdx = $from.index(anchor.tableAt);
          if (rowIdx === 0 || rowIdx >= tableNode.childCount) return false;
          if (dispatch) {
            const tablePos = $from.before(anchor.tableAt);
            const rowsA: PMNode[] = [];
            const rowsB: PMNode[] = [];
            for (let r = 0; r < tableNode.childCount; r += 1) {
              (r < rowIdx ? rowsA : rowsB).push(tableNode.child(r));
            }
            const create = state.schema.nodes.table.create.bind(state.schema.nodes.table);
            const tr = state.tr.replaceWith(tablePos, tablePos + tableNode.nodeSize, [
              create(tableNode.attrs, rowsA),
              create(tableNode.attrs, rowsB),
            ]);
            // The caret lands in the second table's first cell: the first
            // table's size is its rows plus the open/close tokens.
            const firstTableSize = rowsA.reduce((sum, r) => sum + r.nodeSize, 0) + 2;
            tr.setSelection(TextSelection.near(tr.doc.resolve(tablePos + firstTableSize + 1)));
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Word's AutoFit Contents: each column shrinks to its widest cell's
      // content (a character-count heuristic — see measureTextTwip) without
      // growing past the current grid. Span-free tables only.
      "autofit-contents":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          const { $from } = state.selection;
          const tableNode = $from.node(anchor.tableAt);
          const widths = tableNode.attrs.columnWidths as number[] | null;
          if (!widths || widths.length === 0) return false;
          const cols = widths.length;
          for (let r = 0; r < tableNode.childCount; r += 1) {
            const row = tableNode.child(r);
            if (row.childCount !== cols) return false;
            for (let c = 0; c < cols; c += 1) {
              const cell = row.child(c);
              if (cell.attrs.columnSpan || cell.attrs.verticalMerge) return false;
            }
          }
          if (dispatch) {
            const next = widths.map((w, c) => {
              let widest = 0;
              for (let r = 0; r < tableNode.childCount; r += 1) {
                widest = Math.max(widest, measureTextTwip(tableNode.child(r).child(c).textContent));
              }
              return Math.max(MIN_COL_TWIP, Math.min(w, widest));
            });
            dispatch(
              state.tr
                .setNodeMarkup($from.before(anchor.tableAt), undefined, {
                  ...tableNode.attrs,
                  columnWidths: next,
                  layout: null,
                })
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Word's AutoFit Window: the grid scales proportionally to the page's
      // text width (the host resolves that from the layout flow and passes it
      // as the twip value). A table without a grid starts from equal columns.
      "autofit-window":
        (value) =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          const total = Number(value);
          if (!Number.isFinite(total) || total <= 0) return false;
          const { $from } = state.selection;
          const tableNode = $from.node(anchor.tableAt);
          const widths =
            (tableNode.attrs.columnWidths as number[] | null)?.filter((w) => w > 0) ?? [];
          const cols = Math.max(widths.length, tableNode.child(0)?.childCount ?? 0);
          if (cols === 0) return false;
          if (dispatch) {
            const sum = widths.reduce((a, b) => a + b, 0);
            const next = Array.from({ length: cols }, (_, c) =>
              sum > 0 && c < widths.length
                ? Math.max(1, Math.round((widths[c]! / sum) * total))
                : Math.round(total / cols),
            );
            dispatch(
              state.tr
                .setNodeMarkup($from.before(anchor.tableAt), undefined, {
                  ...tableNode.attrs,
                  columnWidths: next,
                  layout: null,
                })
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Word's Fixed Column Width — toggles the tblLayout fixed flag (the
      // grid stops following content; the columns stay where the grid puts
      // them).
      "fixed-column-width":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          if (dispatch) {
            const { $from } = state.selection;
            const tableNode = $from.node(anchor.tableAt);
            dispatch(
              state.tr
                .setNodeMarkup($from.before(anchor.tableAt), undefined, {
                  ...tableNode.attrs,
                  layout: tableNode.attrs.layout === "fixed" ? null : "fixed",
                })
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Word's Distribute Columns: the grid splits its total evenly (the last
      // column absorbs the rounding remainder so the sum is exact).
      "distribute-columns":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          const { $from } = state.selection;
          const tableNode = $from.node(anchor.tableAt);
          const widths = tableNode.attrs.columnWidths as number[] | null;
          if (!widths || widths.length === 0) return false;
          if (dispatch) {
            const sum = widths.reduce((a, b) => a + b, 0);
            const even = Math.floor(sum / widths.length);
            const next = widths.map((_, c) =>
              c === widths.length - 1 ? sum - even * (widths.length - 1) : even,
            );
            dispatch(
              state.tr
                .setNodeMarkup($from.before(anchor.tableAt), undefined, {
                  ...tableNode.attrs,
                  columnWidths: next,
                })
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Cell width — the column width of the caret's cell (the grid is the
      // one width source the layout reads; Word's tcW maps onto it here).
      "cell-width":
        (value) =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.rowAt < 0) return false;
          const tw = parseMeasureTwip(value);
          if (tw == null || tw < MIN_COL_TWIP) return false;
          const { $from } = state.selection;
          const tableNode = $from.node(anchor.tableAt);
          const widths = tableNode.attrs.columnWidths as number[] | null;
          const col = $from.index(anchor.rowAt);
          if (!widths || col >= widths.length) return false;
          if (dispatch) {
            const next = [...widths];
            next[col] = Math.round(tw);
            dispatch(
              state.tr
                .setNodeMarkup($from.before(anchor.tableAt), undefined, {
                  ...tableNode.attrs,
                  columnWidths: next,
                })
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Row height — the caret row's trHeight (atLeast; "0"/auto clears it).
      "cell-height":
        (value) =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.rowAt < 0) return false;
          const tw = parseMeasureTwip(value);
          if (tw == null || tw < 0) return false;
          if (dispatch) {
            const { $from } = state.selection;
            const row = $from.node(anchor.rowAt);
            dispatch(
              state.tr
                .setNodeMarkup($from.before(anchor.rowAt), undefined, {
                  ...row.attrs,
                  height: tw > 0 ? { value: Math.round(tw), rule: "atLeast" } : null,
                })
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Wrap the selection in a link (empty selection → link around the URL text).
      // Word stamps inserted hyperlink runs with the "Hyperlink" character
      // style — that style (not the w:hyperlink element) paints links blue —
      // so the same chain stamps it here (one transaction, one undo step).
      link:
        (href) =>
        ({ chain }) => {
          const url = href || (typeof window !== "undefined" && window.prompt("Link URL")) || "";
          if (!url) return false;
          return chain()
            .extendMarkRange("link")
            .setLink({ href: url })
            .setMark("textStyle", { style: "Hyperlink" })
            .run();
        },

      // ── Style gallery (combobox-driven): value picks the block style ──
      // A HeadingLevel id stamps the paragraph's `heading` attr (a heading IS
      // a paragraph); everything else carries `style` so the injected document
      // CSS applies. The paragraph keeps `style` clear when a HeadingLevel
      // applies — office-open's single pStyle writer prefers `style`, so both
      // set would mask the heading.
      style:
        (styleId) =>
        ({ chain }) => {
          const id = (styleId ?? "").trim();
          if (HEADING_LEVEL_BY_STYLE[id]) {
            return chain().updateAttributes("paragraph", { heading: id, style: null }).run();
          }
          return chain()
            .updateAttributes("paragraph", { style: id || null, heading: null })
            .run();
        },
      // Word's References > Add Text: mark every selected paragraph as a TOC
      // level by stamping its heading pStyle; "none" returns it to body text.
      // The heading wins over a named style (the single pStyle writer prefers
      // `style`), so a level stamp clears it — the same rule the style
      // gallery applies in reverse.
      "add-text":
        (value) =>
        ({ state, tr }) => {
          const heading = ADD_TEXT_LEVELS[value ?? ""];
          if (heading === undefined) return false;
          const blocks = selectedParagraphs(state);
          if (!blocks.length) return false;
          for (const { pos, node } of blocks) {
            const attrs = node.attrs as Record<string, unknown>;
            tr.setNodeMarkup(pos, undefined, {
              ...attrs,
              heading,
              style: heading ? null : ((attrs.style as string | null) ?? null),
            });
          }
          return true;
        },

      // ── Editing — change case / sort / multilevel list level ──
      // Transform selected text to the requested case and replace the
      // selection, preserving the run's marks. No-op on an empty selection.
      "change-case":
        (mode) =>
        ({ state, chain }) => {
          const { from, to, empty } = state.selection;
          if (empty) return false;
          const text = state.doc.textBetween(from, to, "");
          if (!text) return false;
          const out = transformCase(text, mode);
          if (out === text) return false;
          const marks = state.selection.$from.marks();
          return chain()
            .command(({ tr }) => {
              tr.replaceWith(from, to, state.schema.text(out, marks));
              return true;
            })
            .run();
        },
      // Sort the sibling blocks covered by the selection in ascending text
      // order (locale-aware, numeric). Only same-parent block sequences are
      // reorderable — mirroring Word Sort on a paragraph/list range.
      sort:
        () =>
        ({ state, chain }) => {
          const { selection, doc } = state;
          const { from, to, empty } = selection;
          if (empty) return false;
          const $from = doc.resolve(from);
          const $to = doc.resolve(to);
          if ($from.depth !== $to.depth || $from.depth < 1 || $from.parent !== $to.parent)
            return false;
          const depth = $from.depth;
          const parent = $from.parent;
          const children: import("@tiptap/pm/model").Node[] = [];
          parent.forEach((child: import("@tiptap/pm/model").Node) => children.push(child));
          const startIndex = $from.index(depth);
          const endIndex = $to.indexAfter(depth);
          const range = children.slice(startIndex, endIndex);
          if (range.length < 2) return false;
          const sorted = [...range].sort((a, b) =>
            a.textContent.trim().localeCompare(b.textContent.trim(), undefined, { numeric: true }),
          );
          if (sorted.every((node, i) => node === range[i])) return false;
          let startPos = $from.start(depth);
          let endPos = startPos;
          for (const node of range) endPos += node.nodeSize;
          return chain()
            .command(({ tr }) => {
              tr.replaceWith(startPos, endPos, sorted);
              return true;
            })
            .run();
        },
      // Promote/demote the selected list paragraphs to a fixed multilevel
      // depth (level-1 = top, level-2/3 = one/two in), keeping each
      // paragraph's list kind and reference. "in"/"out" (the Bullets and
      // Numbering drop-downs' Change List Level item) step each paragraph
      // relative to its own level — the shared Tab semantics. The split's
      // main click carries no value — top level, not a demotion to level 2.
      // Plain paragraphs gain a fresh decimal multilevel list (Word's gallery
      // applies a list; a silent no-op reads as a broken button).
      "multilevel-list":
        (level) =>
        ({ state, tr }) => {
          const step = level === "in" ? 1 : level === "out" ? -1 : 0;
          const target =
            step === 0 ? (level === "level-3" ? 2 : level === "level-2" ? 1 : 0) : null;
          let touched = false;
          let freshRef: string | null = null;
          for (const { pos, node } of selectedParagraphs(state)) {
            const attrs = node.attrs as Record<string, unknown>;
            const cur = listStateOf(attrs);
            if (!cur.kind) {
              // One shared list for the whole selection (Word numbers the
              // applied gallery as one list).
              freshRef ??= nextOrderedReference(
                collectListReferences(state.doc),
                (state.doc.attrs as { numbering?: unknown }).numbering,
              );
              tr.setNodeMarkup(pos, undefined, {
                ...attrs,
                bullet: null,
                numbering: { reference: freshRef, level: step === 0 ? target! : 0 },
              });
              touched = true;
              continue;
            }
            const depth = step === 0 ? target! : Math.min(8, Math.max(0, cur.level + step));
            if (depth === cur.level) continue;
            const patch =
              cur.kind === "bullet" && cur.variant === "bullet"
                ? { bullet: { level: depth }, numbering: null }
                : { bullet: null, numbering: { reference: cur.reference, level: depth } };
            tr.setNodeMarkup(pos, undefined, { ...attrs, ...patch });
            touched = true;
          }
          return touched;
        },
      // Delete the currently selected image node (mirrors Office.js
      // InlinePicture.delete()). Only fires on an image NodeSelection.
      "delete-picture":
        () =>
        ({ state, commands }) => {
          const sel = state.selection;
          if (!(sel instanceof NodeSelection) || sel.node.type.name !== "image") return false;
          return commands.deleteSelection();
        },
      // Reposition a floating (wp:anchor wrapNone) image by writing new EMU
      // offsets into its floating attrs. value is JSON {hOffset, vOffset}.
      // The host image NodeView dispatches this on drag end.
      "position-picture":
        (value?) =>
        ({ state, tr }) => {
          if (!value) return false;
          const sel = state.selection;
          if (!(sel instanceof NodeSelection) || sel.node.type.name !== "image") return false;
          let parsed: { hOffset?: number; vOffset?: number };
          try {
            parsed = JSON.parse(value);
          } catch {
            return false;
          }
          const old = sel.node.attrs as ImageAttrs;
          if (!old.floating) return false;
          // align and offset are mutually exclusive in OOXML — writing offset
          // must clear align, or the serializer ignores the offset. Preserve
          // relative (default would otherwise become "page").
          const h = old.floating.horizontalPosition;
          const v = old.floating.verticalPosition;
          tr.setNodeMarkup(sel.from, undefined, {
            ...old,
            floating: {
              ...old.floating,
              horizontalPosition: { relative: h.relative, offset: parsed.hOffset ?? h.offset },
              verticalPosition: { relative: v.relative, offset: parsed.vOffset ?? v.offset },
            },
          });
          // Suppress scrollIntoView — for a position drag the user is already
          // looking at the image and a scroll jump would feel jumpy.
          tr.setMeta("scrollIntoView", false);
          return true;
        },
      // ── Arrange — floating drawings (the Layout tab's Arrange group) ──
      // Every command targets the selected floating drawing (a floating
      // image or a wps shape); on any other selection they decline, so the
      // ribbon greys them out through editor.can().

      // Word's Bring Forward / Send Backward: step w:relativeHeight within
      // the drawing's behind/in-front band; the painter stacks same-band
      // floats by it (ties keep document order).
      "bring-forward":
        () =>
        ({ state, tr }) => {
          const target = floatingDrawingAt(state);
          if (!target) return false;
          const floating = floatingOf(target);
          return stampFloating(tr, target, {
            ...floating,
            zIndex: (typeof floating.zIndex === "number" ? floating.zIndex : 0) + 1,
          });
        },
      "send-backward":
        () =>
        ({ state, tr }) => {
          const target = floatingDrawingAt(state);
          if (!target) return false;
          const floating = floatingOf(target);
          return stampFloating(tr, target, {
            ...floating,
            zIndex: Math.max(0, (typeof floating.zIndex === "number" ? floating.zIndex : 0) - 1),
          });
        },
      // Word's Wrap Text menu: In Front of Text / Behind Text clear the wrap
      // (wrapNone) and set behindDoc; the four wrap styles stamp the type
      // and drop behindDoc (Word 2013+ honors it for wrapNone anchors only).
      wrap:
        (value) =>
        ({ state, tr }) => {
          const target = floatingDrawingAt(state);
          if (!target) return false;
          const floating = { ...floatingOf(target) };
          if (value === "front" || value === "behind") {
            delete floating.wrap;
            floating.behindDocument = value === "behind";
          } else if (value === "square" || value === "tight" || value === "through") {
            floating.wrap = { type: value };
            floating.behindDocument = false;
          } else if (value === "top-bottom") {
            floating.wrap = { type: "topAndBottom" };
            floating.behindDocument = false;
          } else {
            return false;
          }
          return stampFloating(tr, target, floating);
        },
      // Word's Rotate menu: right/left step the rotation 90° (OOXML rot is
      // clockwise-positive); the flips toggle the mirror flags. The attrs
      // live in two places — an image carries rotation/flipH/flipV on its
      // top level (a tri-state: null omits, true/false emit explicit bytes),
      // a shape mirrors them inside its transformation.
      rotate:
        (value) =>
        ({ state, tr }) => {
          const target = floatingDrawingAt(state);
          if (!target) return false;
          const step = value === "right" ? 90 : value === "left" ? -90 : 0;
          if (target.kind === "image") {
            const attrs = { ...target.attrs };
            if (step !== 0) {
              const rotation = typeof attrs.rotation === "number" ? attrs.rotation : 0;
              attrs.rotation = (((rotation + step) % 360) + 360) % 360;
            } else if (value === "flip-h") {
              attrs.flipH = attrs.flipH !== true;
            } else if (value === "flip-v") {
              attrs.flipV = attrs.flipV !== true;
            } else {
              return false;
            }
            return stampAttrs(tr, target, attrs);
          }
          const shape = { ...(target.attrs.wpsShape as Record<string, unknown>) };
          const t = { ...((shape.transformation ?? {}) as Record<string, unknown>) };
          if (step !== 0) {
            const rotation = typeof t.rotation === "number" ? t.rotation : 0;
            t.rotation = (((rotation + step) % 360) + 360) % 360;
          } else if (value === "flip-h") {
            t.flipHorizontal = t.flipHorizontal !== true;
          } else if (value === "flip-v") {
            t.flipVertical = t.flipVertical !== true;
          } else {
            return false;
          }
          shape.transformation = t;
          return stampAttrs(tr, target, { ...target.attrs, wpsShape: shape });
        },
      // Word's Position gallery: the nine-cell grid stamps margin-relative
      // align tokens on both axes. A fresh position object per stamp — align
      // and offset are mutually exclusive, so a stale offset must not
      // survive next to the new align.
      position:
        (value) =>
        ({ state, tr }) => {
          const spec = POSITION_ALIGN[value ?? ""];
          if (!spec) return false;
          const target = floatingDrawingAt(state);
          if (!target) return false;
          return stampFloating(tr, target, {
            ...floatingOf(target),
            horizontalPosition: { relative: "margin", align: spec.h },
            verticalPosition: { relative: "margin", align: spec.v },
          });
        },
      // The Align menu: horizontal alignment within the margins (the single
      // axis of the position gallery).
      "align-objects":
        (value) =>
        ({ state, tr }) => {
          const align = value === "left" || value === "center" || value === "right" ? value : null;
          if (!align) return false;
          const target = floatingDrawingAt(state);
          if (!target) return false;
          return stampFloating(tr, target, {
            ...floatingOf(target),
            horizontalPosition: { relative: "margin", align },
          });
        },
    };
  },
});
