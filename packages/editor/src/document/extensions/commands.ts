import type { ImageAttrs, TableFloatOptions } from "@docen/docx";
import {
  BULLET_GLYPHS,
  HIGHLIGHT_PALETTE_RGB,
  nextMultilevelReference,
  nextOrderedReference,
  ORDERED_FORMATS,
} from "@docen/docx";
import { Extension } from "@docen/docx/core";
import type { Node as PMNode, ResolvedPos } from "@tiptap/pm/model";
import type { Mark } from "@tiptap/pm/model";
import type { EditorState } from "@tiptap/pm/state";
import type { Transaction } from "@tiptap/pm/state";
import { NodeSelection, Selection, TextSelection } from "@tiptap/pm/state";
import { DocAttrStep } from "@tiptap/pm/transform";

import { CellSelection, cellsInRect, gridColOf, spanOf } from "../canvas/cell-selection";

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
      "underline-style": (style?: string, color?: string | null) => ReturnType;
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
      "justify-distribute": () => ReturnType;
      "indent-increase": () => ReturnType;
      "indent-decrease": () => ReturnType;
      "line-spacing": (mult?: string) => ReturnType;
      "paragraph-dialog-apply": (patch?: ParagraphDialogPatch) => ReturnType;
      shading: (value?: unknown) => ReturnType;
      "font-color": (value?: unknown) => ReturnType;
      border: (side?: string) => ReturnType;
      "borders-apply": (patch?: BordersDialogPatch) => ReturnType;
      // Lists / blocks
      "bullet-list": (variant?: string) => ReturnType;
      "ordered-list": (variant?: string) => ReturnType;
      blockquote: () => ReturnType;
      "horizontal-rule": () => ReturnType;
      "page-break": () => ReturnType;
      "column-break": () => ReturnType;
      "section-break": () => ReturnType;
      "section-break-next": () => ReturnType;
      "section-break-continuous": () => ReturnType;
      "insert-table": (options?: InsertTableOptions) => ReturnType;
      "delete-table": () => ReturnType;
      // Table context commands (the Table Design / Layout contextual tabs).
      "insert-row-above": () => ReturnType;
      "insert-row-below": () => ReturnType;
      "insert-row-at": (index: number, tablePos?: number) => ReturnType;
      "insert-column-left": () => ReturnType;
      "insert-column-right": () => ReturnType;
      "insert-column-at": (index: number, tablePos?: number) => ReturnType;
      "set-table-column-widths": (widths: number[], tablePos?: number) => ReturnType;
      "set-table-row-height": (
        rowIndex: number,
        height: { rule: "atLeast" | "exact"; value: number } | null,
        tablePos?: number,
      ) => ReturnType;
      "delete-row": () => ReturnType;
      "delete-column": () => ReturnType;
      "select-table": () => ReturnType;
      "select-table-row": () => ReturnType;
      "select-table-cell": () => ReturnType;
      "select-table-column": () => ReturnType;
      "align-cell": (value?: string) => ReturnType;
      "repeat-header-rows": (value?: boolean) => ReturnType;
      "cant-split": (value?: boolean) => ReturnType;
      "cell-shading": (value?: unknown) => ReturnType;
      "cell-borders": (value?: unknown) => ReturnType;
      "set-cell-insets": (value?: unknown) => ReturnType;
      "set-cell-vertical-align": (value: "top" | "center" | "bottom" | null) => ReturnType;
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
      "distribute-rows": () => ReturnType;
      "cell-margins": (value?: string) => ReturnType;
      "cell-width": (value?: string) => ReturnType;
      "cell-height": (value?: string) => ReturnType;
      "table-alignment": (value: "left" | "center" | "right") => ReturnType;
      "table-text-wrapping": (
        value: "none" | "around" | boolean | TableFloatOptions | null,
      ) => ReturnType;
      "table-properties-apply": (patch?: TablePropertiesPatch) => ReturnType;
      "convert-text-to-table": (options?: ConvertTextToTableOptions | string) => ReturnType;
      "sort-table": (options?: SortTableOptions | "asc" | "desc") => ReturnType;
      "table-formula": (options?: TableFormulaOptions | string) => ReturnType;
      "table-cell-spacing": (
        value?: number | string | { size: number; type?: string } | null,
      ) => ReturnType;
      link: (href?: string) => ReturnType;
      style: (styleId?: string) => ReturnType;
      "modify-style": (patch?: ModifyStylePatch) => ReturnType;
      "style-set": (value?: string) => ReturnType;
      "add-text": (value?: string) => ReturnType;
      // Editing
      "change-case": (mode?: string) => ReturnType;
      sort: () => ReturnType;
      "multilevel-list": (level?: string) => ReturnType;
      // Picture — names align to Office.js InlinePicture (delete / left / top).
      "delete-picture": () => ReturnType;
      "position-picture": (value?: string) => ReturnType;
      "move-drawing": (value?: string) => ReturnType;
      "place-drawing": (value?: string) => ReturnType;
      "rotate-drawing": (value?: string) => ReturnType;
      "drawing-properties-apply": (patch?: DrawingPropertiesPatch) => ReturnType;
      "drawing-crop-apply": (patch?: DrawingCropPatch) => ReturnType;
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
  "underline-style",
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
  "justify-distribute",
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
  "section-break-next",
  "section-break-continuous",
  "insert-table",
  "delete-table",
  "insert-row-above",
  "insert-row-below",
  "insert-row-at",
  "insert-column-left",
  "insert-column-right",
  "insert-column-at",
  "set-table-column-widths",
  "set-table-row-height",
  "delete-row",
  "delete-column",
  "select-table",
  "select-table-row",
  "select-table-cell",
  "select-table-column",
  "align-cell",
  "repeat-header-rows",
  "cant-split",
  "cell-shading",
  "cell-borders",
  "set-cell-insets",
  "set-cell-vertical-align",
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
  "distribute-rows",
  "cell-margins",
  "cell-width",
  "cell-height",
  "table-alignment",
  "table-text-wrapping",
  "table-properties-apply",
  "text-direction",
  "convert-to-text",
  "convert-text-to-table",
  "sort-table",
  "table-formula",
  "table-cell-spacing",
  "link",
  "style",
  "modify-style",
  "style-set",
  "add-text",
  "undo",
  "redo",
  "change-case",
  "sort",
  "multilevel-list",
  "delete-picture",
  "position-picture",
  "move-drawing",
  "place-drawing",
  "rotate-drawing",
  "drawing-properties-apply",
  "drawing-crop-apply",
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

/**
 * The full attr patch the Paragraph dialog commits on OK — every field always
 * present (Office commits the dialog atomically): "special: none" is an
 * explicit clear of firstLine/hanging, "body text" an explicit clear of
 * outlineLevel. Indent/spacing values are OOXML twips; lineRule the w:spacing
 * tokens. Stamped onto every selected paragraph by
 * {@link documentCommands.paragraph-dialog-apply}.
 */
/** Options for the insert-table command (all fields fall back to Word's
 *  3×3 default preset; rows/cols are clamped to the schema-safe range). */
export interface InsertTableOptions {
  /** Row count (1-50, default 3 — the first row is the header row). */
  rows?: number;
  /** Column count (1-10, default 3). */
  cols?: number;
}

/**
 * What the Modify Style dialog commits on OK — the style's chain pointers
 * plus the run formatting the dialog edits, stamped onto the named style by
 * {@link documentCommands.modify-style}. `null` clears the field (the style
 * inherits from its basedOn chain again); the run booleans are absolute
 * states, matching the dialog's checkboxes.
 */
export interface ModifyStylePatch {
  /** The styleId to modify (e.g. "Normal", "Heading1", a custom id). */
  id: string;
  basedOn: string | null;
  /** The style applied to the next paragraph typed after this one. */
  next: string | null;
  font: string | null;
  /** Font size in points. */
  size: number | null;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  /** Hex without "#", or null for the automatic (text) color. */
  color: string | null;
}

const ALIGN_VALUES = ["left", "center", "right", "both", "distribute"] as const;

// CSS/OOXML synonyms fold onto the paragraph attr's vocabulary.
const ALIGN_ALIASES: Record<string, string> = { start: "left", end: "right", justify: "both" };

/** Normalize an incoming paragraph alignment (w:jc val, CSS text-align, or a
 *  dialog string) onto the paragraph attr's vocabulary; unknown → "left". */
export function normalizeParagraphAlignment(raw: unknown): string {
  const alignment = typeof raw === "string" ? (ALIGN_ALIASES[raw] ?? raw) : "left";
  return (ALIGN_VALUES as readonly string[]).includes(alignment) ? alignment : "left";
}

export interface ParagraphDialogPatch {
  alignment: string;
  outlineLevel: number | null;
  indent: {
    left?: number;
    right?: number;
    firstLine?: number;
    hanging?: number;
  };
  spacing: {
    before?: number;
    after?: number;
    line?: number;
    lineRule?: "auto" | "atLeast" | "exact";
  };
  widowControl: boolean;
  keepNext: boolean;
  keepLines: boolean;
  pageBreakBefore: boolean;
  // The Line-and-Page-Breaks tab's formatting exceptions.
  suppressLineNumbers: boolean;
  suppressAutoHyphens: boolean;
  // The Chinese-Layout tab: kinsoku/wordWrap/overflowPunct are the line-break
  // group, the two autoSpace flags the character-spacing group (autoSpaceDN
  // rides the engine's autoSpaceEastAsianText attr), textAlignment the
  // vertical alignment drop-down (w:textAlignment tokens).
  kinsoku: boolean;
  wordWrap: boolean;
  overflowPunct: boolean;
  autoSpaceDE: boolean;
  autoSpaceDN: boolean;
  textAlignment: string;
}

/**
 * What the Table Properties dialog or programmatic caller commits on apply —
 * the table tab's geometry (alignment, left indent, float/wrapping) and row
 * properties (cantSplit, tableHeader), stamped onto the caret's table by
 * {@link documentCommands.table-properties-apply}.
 */
export interface TablePropertiesPatch {
  /** w:jc table alignment — "left" commits null (OOXML's default). */
  alignment?: "left" | "center" | "right";
  /** w:tblInd left indent in twips; 0 commits null. */
  indent?: number;
  /** w:cantSplit row protection flag on targeted rows. */
  cantSplit?: boolean;
  /** w:tblHeader repeat header rows flag on targeted rows. */
  tableHeader?: boolean;
  /** w:tblpPr text wrapping ("none" clears float, "around" enables float). */
  textWrapping?: "none" | "around";
  /** w:tblpPr floating options directly. */
  float?: TableFloatOptions | null;
  /** w:tblCellSpacing cell spacing in twips, universal measure string, or object. */
  cellSpacing?: number | string | { size: number; type?: string } | null;
}

/** Options for {@link documentCommands.convert-text-to-table}. */
export interface ConvertTextToTableOptions {
  /** Delimiter to split columns by. If omitted, auto-detected (\t, ,, ;, |). */
  delimiter?: string | RegExp;
  /** If true, the first row is marked as tableHeader. */
  hasHeader?: boolean;
}

/** Options for {@link documentCommands.sort-table}. */
export interface SortTableOptions {
  /** 0-based column index to sort by. Defaults to active cell's column. */
  column?: number;
  /** Sort order. Defaults to "asc". */
  order?: "asc" | "desc";
  /** Whether row 0 is treated as a table header (not sorted).
   * Defaults to whether row 0 has tableHeader: true. */
  hasHeader?: boolean;
}

/** Options for {@link documentCommands.table-formula}. */
export interface TableFormulaOptions {
  /** The formula string, e.g. "=SUM(ABOVE)", "=AVERAGE(LEFT)", "=COUNT(A1:B3)", etc.
   * If omitted, auto-picks =SUM(ABOVE) or =SUM(LEFT) based on adjacent numbers. */
  formula?: string;
  /** Optional number format or prefix/suffix. */
  numberFormat?: string;
}

/**
 * What the Size-and-Position dialog commits on OK — the selected floating
 * drawing's geometry in centimeters (the dialog's display unit), stamped by
 * {@link documentCommands.drawing-properties-apply}. Absent fields keep the
 * current value.
 */
export interface DrawingPropertiesPatch {
  /** The drawing box width in cm (px for an image, EMU for a shape payload). */
  widthCm: number;
  /** The drawing box height in cm. */
  heightCm: number;
  /** Clockwise rotation about the box center, degrees. */
  rotationDeg: number;
  /** Horizontal offset from the anchor's horizontal base, cm → EMU. */
  offsetHCm: number;
  /** Vertical offset from the anchor's vertical base, cm → EMU. */
  offsetVCm: number;
}

/**
 * What the crop mode commits — the selected image's a:srcRect insets as
 * fractions of the source (0.1 = 10% off that edge), stamped by {@link
 * documentCommands.drawing-crop-apply}. An all-zero set clears the crop.
 */
export interface DrawingCropPatch {
  /** Left inset as a source fraction (0.1 = 10% cropped off the left). */
  left: number;
  /** Top inset as a source fraction. */
  top: number;
  /** Right inset as a source fraction. */
  right: number;
  /** Bottom inset as a source fraction. */
  bottom: number;
}

/** One w:pBdr edge as the dialog stages it: the ST_Border style token, the
 *  width in eighths of a point, and the hex color (null = auto ink). */
export interface BorderSideState {
  style: string;
  size: number;
  color: string | null;
}

/** What the Borders and Shading dialog commits on OK: the stamp target tab,
 *  the per-edge borders (border/page tabs — every edge present, null clears),
 *  or the paragraph fill (shading tab, null clears). */
export interface BordersDialogPatch {
  tab: "border" | "page" | "shading";
  sides?: Partial<Record<"top" | "bottom" | "left" | "right", BorderSideState | null>>;
  /** Hex RRGGBB paragraph fill; null clears the shading. */
  fill?: string | null;
}

// ── Pure helpers (take EditorState, return data; never touch the chain) ──

/** The Design tab's built-in style sets — body + heading font families written
 *  onto the document defaults and the built-in heading styles (Word's Style
 *  Set gallery swaps the theme fonts the same way). Keys are the
 *  DefaultStylesOptions slots ("document" = the docDefaults run defaults). */
const STYLE_SET_PRESETS: Readonly<Record<string, Readonly<Record<string, unknown>>>> = {
  modern: {
    document: { font: "Calibri" },
    title: { font: "Calibri Light" },
    heading1: { font: "Calibri Light" },
    heading2: { font: "Calibri Light" },
    heading3: { font: "Calibri Light" },
  },
  classic: {
    document: { font: "Times New Roman" },
    title: { font: "Cambria" },
    heading1: { font: "Cambria" },
    heading2: { font: "Cambria" },
    heading3: { font: "Cambria" },
  },
  elegant: {
    document: { font: "Georgia" },
    title: { font: "Georgia" },
    heading1: { font: "Georgia" },
    heading2: { font: "Georgia" },
    heading3: { font: "Georgia" },
  },
};

/** Copy a style entry with the Modify Style dialog's chain pointers and run
 *  formatting applied. `null` fields are cleared (inherit again); the JSON
 *  round-trip drops the undefined holes the clearing leaves behind and keeps
 *  the stamped model structured-cloneable. */
function withModifyStylePatch(
  entry: Record<string, unknown>,
  patch: ModifyStylePatch,
): Record<string, unknown> {
  const run = { ...((entry.run ?? {}) as Record<string, unknown>) };
  run.font = patch.font ?? undefined;
  // w:sz travels with w:szCs (Word writes the pair together) — one patch
  // field stamps both.
  run.size = patch.size ?? undefined;
  run.sizeComplexScript = patch.size ?? undefined;
  run.bold = patch.bold;
  run.italic = patch.italic;
  run.underline = patch.underline ? { type: "single" } : undefined;
  run.color = patch.color ?? undefined;
  const out: Record<string, unknown> = { ...entry, run };
  if (patch.basedOn) out.basedOn = patch.basedOn;
  else delete out.basedOn;
  if (patch.next) out.next = patch.next;
  else delete out.next;
  return JSON.parse(JSON.stringify(out)) as Record<string, unknown>;
}

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

/** The paragraphs the selection covers, with their positions. Supports
 *  both regular selections and rectangular CellSelection ranges. */
function selectedParagraphs(state: EditorState): { pos: number; node: PMNode }[] {
  if (state.selection instanceof CellSelection) {
    const targets = tableTargets(state);
    if (targets && targets.cells.length > 0) {
      const out: { pos: number; node: PMNode }[] = [];
      for (const { pos: cellPos, node: cellNode } of targets.cells) {
        cellNode.descendants((node, offset) => {
          if (node.type.name === "paragraph") {
            out.push({ pos: cellPos + 1 + offset, node });
          }
          return true;
        });
      }
      return out;
    }
  }
  const { from, to } = state.selection;
  const out: { pos: number; node: PMNode }[] = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name === "paragraph") out.push({ pos, node });
    return true;
  });
  return out;
}

/** Stamp the alignment onto every selected paragraph directly — PM's
 *  updateAttributes would walk the CellSelection's bounding range and bleed
 *  onto cells the rectangular selection skips. */
function setParagraphAlignment(state: EditorState, tr: Transaction, alignment: string): boolean {
  const paras = selectedParagraphs(state);
  if (!paras.length) return false;
  for (const { pos, node } of paras) {
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, alignment });
  }
  return true;
}

/** Every numbering reference the doc's list paragraphs carry — feeds the
 *  fresh-reference allocator so a new list never collides with an existing
 *  one's numbering. */
export function collectListReferences(doc: PMNode): string[] {
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
  if (value === "none" || value === null) return null;
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
  if (typeof value === "object" && value !== null && "fill" in value) {
    const obj = value as Record<string, unknown>;
    return {
      type: "clear",
      ...obj,
    };
  }
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
export function ancestryAt($pos: ResolvedPos): {
  tableAt: number;
  rowAt: number;
  cellAt: number;
} | null {
  const { table, tableRow, tableCell } = $pos.doc.type.schema.nodes;
  let tableAt = -1;
  let rowAt = -1;
  let cellAt = -1;
  for (let d = $pos.depth; d > 0; d -= 1) {
    const t = $pos.node(d).type;
    if (t === tableCell && cellAt < 0) cellAt = d;
    else if (t === tableRow && rowAt < 0) rowAt = d;
    else if (t === table && tableAt < 0) tableAt = d;
  }
  return tableAt >= 0 ? { tableAt, rowAt, cellAt } : null;
}

/** The cell targets of a table selection: when a {@link CellSelection} is
 *  active, all cells inside its rectangular bounds; otherwise the single cell
 *  holding the text selection. Both carry the enclosing table's node and
 *  document position, plus the distinct rows and columns the selection touches.
 *
 *  `tableAncestry` only looks up the `$from` path — on a CellSelection `depth -
 *  1` there and every cell-level command would decline — this is the one
 *  resolver the cell/row/column-level commands share. Cell stamps carry their
 *  table-child row index and grid column; the row/column commands read
 *  `rows`/`cols` (Word: a whole-pick height lands on every picked row, a
 *  width on every picked column). Null outside a table. */
function tableTargets(state: EditorState): {
  tablePos: number;
  tableNode: PMNode;
  cells: { pos: number; node: PMNode; row: number; col: number }[];
  rows: Set<number>;
  cols: Set<number>;
} | null {
  const { selection } = state;
  const isCell = selection instanceof CellSelection;
  let anchorCell: number;
  let tablePos: number;
  if (isCell) {
    anchorCell = selection.anchorCell;
    const $t = selection.$from;
    tablePos = $t.before($t.depth - 1);
  } else {
    const anchor = ancestryAt(selection.$from);
    if (!anchor || anchor.cellAt < 0) return null;
    anchorCell = selection.$from.before(anchor.cellAt);
    tablePos = selection.$from.before(anchor.tableAt);
  }
  const cells: { pos: number; node: PMNode; row: number; col: number }[] = [];
  const rows = new Set<number>();
  const cols = new Set<number>();
  cellsInRect(
    state.doc,
    anchorCell,
    isCell ? selection.headCell : anchorCell,
    (node, pos, row, col) => {
      cells.push({ pos, node, row, col });
      rows.add(row);
      const span = spanOf(node);
      for (let c = col; c < col + span; c += 1) cols.add(c);
    },
  );
  const tableNode = state.doc.nodeAt(tablePos);
  if (!cells.length || !tableNode) return null;
  return { tablePos, tableNode, cells, rows, cols };
}

/** The selected rows stamped through one forward walk (markup writes keep
 *  positions stable, so row positions stay valid as they're written). */
function stampRows(
  tr: Transaction,
  targets: NonNullable<ReturnType<typeof tableTargets>>,
  patch: (row: PMNode) => Record<string, unknown>,
): void {
  let rowPos = targets.tablePos + 1;
  for (let r = 0; r < targets.tableNode.childCount; r += 1) {
    const row = targets.tableNode.child(r)!;
    if (targets.rows.has(r)) tr.setNodeMarkup(rowPos, undefined, patch(row));
    rowPos += row.nodeSize;
  }
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

/** The Cell Margins menu presets (Word's Table Layout → Alignment group):
 *  null clears the cell's tcMar so the table's default applies; the named
 *  presets stamp Word's twip values (0 top/bottom, narrow 0.075" / wide 0.2"
 *  left/right). */
const CELL_MARGIN_PRESETS: Readonly<Record<string, Record<string, unknown> | null>> = {
  default: null,
  none: {
    top: { size: 0, type: "twips" },
    right: { size: 0, type: "twips" },
    bottom: { size: 0, type: "twips" },
    left: { size: 0, type: "twips" },
  },
  narrow: {
    top: { size: 0, type: "twips" },
    right: { size: 108, type: "twips" },
    bottom: { size: 0, type: "twips" },
    left: { size: 108, type: "twips" },
  },
  wide: {
    top: { size: 0, type: "twips" },
    right: { size: 288, type: "twips" },
    bottom: { size: 0, type: "twips" },
    left: { size: 288, type: "twips" },
  },
};

/** Parse cell insets / margins from a preset name, direct numbers, or side spec. */
function parseCellInsets(value: unknown): Record<string, unknown> | null | undefined {
  if (value === null || value === "default") return null;
  if (typeof value === "string") {
    return CELL_MARGIN_PRESETS.hasOwnProperty(value) ? CELL_MARGIN_PRESETS[value] : undefined;
  }
  if (typeof value === "number") {
    const side = { size: Math.round(value), type: "twips" };
    return { top: side, bottom: side, left: side, right: side };
  }
  if (typeof value === "object" && value !== null) {
    const raw = value as Record<string, unknown>;
    const toMarginSide = (v: unknown) => {
      if (typeof v === "number") return { size: Math.round(v), type: "twips" };
      if (typeof v === "object" && v !== null && "size" in v) return v;
      return undefined;
    };
    const margins: Record<string, unknown> = {};
    if ("top" in raw) margins.top = toMarginSide(raw.top);
    if ("bottom" in raw) margins.bottom = toMarginSide(raw.bottom);
    if ("left" in raw) margins.left = toMarginSide(raw.left);
    if ("right" in raw) margins.right = toMarginSide(raw.right);
    return margins;
  }
  return undefined;
}

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

/** Total number of grid columns across all rows in tableNode (span-aware). */
function tableGridColumns(tableNode: PMNode): number {
  if (Array.isArray(tableNode.attrs.columnWidths) && tableNode.attrs.columnWidths.length > 0) {
    return tableNode.attrs.columnWidths.length;
  }
  let maxCols = 0;
  for (let r = 0; r < tableNode.childCount; r += 1) {
    const row = tableNode.child(r);
    let cols = 0;
    for (let c = 0; c < row.childCount; c += 1) {
      cols += spanOf(row.child(c));
    }
    maxCols = Math.max(maxCols, cols);
  }
  return maxCols;
}

/** Computes the available text body width in twips from the document's sectionProperties. */
function availableTextWidthTwips(state: EditorState): number {
  let sectPr: Record<string, unknown> | undefined;
  const from = state.selection.from;
  state.doc.descendants((node, nodePos) => {
    if (nodePos + node.nodeSize <= from) return true;
    if (node.type.name === "paragraph" && node.attrs.sectionProperties) {
      sectPr = node.attrs.sectionProperties as Record<string, unknown>;
      return false;
    }
    return true;
  });
  if (!sectPr) {
    sectPr = (state.doc.attrs as { sectionProperties?: Record<string, unknown> })
      ?.sectionProperties;
  }
  const pageSize = (
    sectPr?.pageSize && typeof sectPr.pageSize === "object" ? sectPr.pageSize : {}
  ) as Record<string, unknown>;
  const pageMargin = (
    sectPr?.pageMargin && typeof sectPr.pageMargin === "object" ? sectPr.pageMargin : {}
  ) as Record<string, unknown>;
  const pageWidth = typeof pageSize.width === "number" ? pageSize.width : 12240;
  const marginLeft = typeof pageMargin.left === "number" ? pageMargin.left : 1440;
  const marginRight = typeof pageMargin.right === "number" ? pageMargin.right : 1440;
  const textWidth = pageWidth - (marginLeft + marginRight);
  return textWidth > 0 ? textWidth : 9360;
}

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

/** Apply borders preset or custom borders to targeted cells. */
function applyCellBorders(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  value: unknown,
): boolean {
  const targets = tableTargets(state);
  if (!targets) return false;

  if (value === "none" || value === null) {
    if (dispatch) {
      const tr = state.tr;
      for (const { pos, node: cell } of targets.cells) {
        tr.setNodeMarkup(pos, undefined, { ...cell.attrs, borders: null });
      }
      dispatch(tr.scrollIntoView());
    }
    return true;
  }

  let edge: { style: string; size: number; color: string } = GRID_BORDER;
  let preset: string | undefined;
  let directBorders: Record<string, unknown> | undefined;

  if (typeof value === "string") {
    preset = value;
  } else if (typeof value === "object" && value !== null) {
    const obj = value as Record<string, unknown>;
    if ("border" in obj && typeof obj.border === "object" && obj.border !== null) {
      edge = { ...GRID_BORDER, ...(obj.border as Record<string, unknown>) };
    }
    if ("preset" in obj && typeof obj.preset === "string") {
      preset = obj.preset;
    } else if ("side" in obj && typeof obj.side === "string") {
      preset = obj.side;
    } else if ("top" in obj || "bottom" in obj || "left" in obj || "right" in obj) {
      directBorders = obj;
    }
  }

  if (!preset && !directBorders) return false;

  if (dispatch) {
    const minRow = Math.min(...targets.rows);
    const maxRow = Math.max(...targets.rows);
    const minCol = Math.min(...targets.cols);
    const maxCol = Math.max(...targets.cols);
    const tr = state.tr;

    for (const { pos, node: cell, row, col } of targets.cells) {
      const span = spanOf(cell);
      const cellEndCol = col + span - 1;
      const current = (cell.attrs.borders ?? {}) as Record<string, unknown>;
      let next: Record<string, unknown>;

      if (directBorders) {
        next = { ...current, ...directBorders };
      } else {
        next = { ...current };
        switch (preset) {
          case "all":
            next.top = edge;
            next.bottom = edge;
            next.left = edge;
            next.right = edge;
            break;
          case "outside":
            if (row === minRow) next.top = edge;
            if (row === maxRow) next.bottom = edge;
            if (col === minCol) next.left = edge;
            if (cellEndCol === maxCol) next.right = edge;
            break;
          case "inside":
            if (row > minRow) next.top = edge;
            if (row < maxRow) next.bottom = edge;
            if (col > minCol) next.left = edge;
            if (cellEndCol < maxCol) next.right = edge;
            break;
          case "insideHorizontal":
            if (row > minRow) next.top = edge;
            if (row < maxRow) next.bottom = edge;
            break;
          case "insideVertical":
            if (col > minCol) next.left = edge;
            if (cellEndCol < maxCol) next.right = edge;
            break;
          case "top":
            if (row === minRow) next.top = edge;
            break;
          case "bottom":
            if (row === maxRow) next.bottom = edge;
            break;
          case "left":
            if (col === minCol) next.left = edge;
            break;
          case "right":
            if (cellEndCol === maxCol) next.right = edge;
            break;
          default:
            return false;
        }
      }

      tr.setNodeMarkup(pos, undefined, { ...cell.attrs, borders: next });
    }
    dispatch(tr.scrollIntoView());
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
      // The underline split's style pick: "none" clears, a token applies the
      // pattern — merging over the current mark so the color survives. The
      // current mark comes from any run in the selection ($from.marks() is
      // empty across a whole-document selection).
      "underline-style":
        (style, color) =>
        ({ state, commands }) => {
          if (!style || style === "none") return commands.unsetMark("underline");
          let current: Mark | undefined;
          state.doc.nodesBetween(state.selection.from, state.selection.to, (node) => {
            current ??= node.marks.find((m) => m.type.name === "underline");
            return !current;
          });
          return commands.setMark("underline", {
            ...((current?.attrs ?? {}) as Record<string, unknown>),
            style,
            // The Font dialog passes an explicit color ("automatic" = null);
            // the ribbon menu omits it and keeps whatever the mark carried.
            ...(color !== undefined ? { color } : {}),
          });
        },
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
          // button's main click) applies Word's default yellow. A value that is
          // already an ST_HighlightColor token (the color picker's highlight
          // palette emits tokens verbatim) passes straight through.
          if (value === "none") return commands.unsetMark("highlight");
          const token =
            HIGHLIGHT_TOKENS[value ?? ""] ??
            (value && value in HIGHLIGHT_PALETTE_RGB ? value : "yellow");
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
        ({ state, tr }) =>
          setParagraphAlignment(state, tr, "left"),
      "align-center":
        () =>
        ({ state, tr }) =>
          setParagraphAlignment(state, tr, "center"),
      "align-right":
        () =>
        ({ state, tr }) =>
          setParagraphAlignment(state, tr, "right"),
      justify:
        () =>
        ({ state, tr }) =>
          setParagraphAlignment(state, tr, "both"),
      "justify-distribute":
        () =>
        ({ state, tr }) =>
          setParagraphAlignment(state, tr, "distribute"),

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
      // The Paragraph dialog's OK — stamp its full patch onto every selected
      // paragraph. firstLine/hanging arrive mutually exclusive (the unchosen
      // key is undefined and clears), so the spread over the current indent
      // commits the switch; the booleans always write (the dialog commits
      // atomically, Word-style).
      "paragraph-dialog-apply":
        (patch) =>
        ({ state, tr }) => {
          if (!patch) return false;
          let touched = false;
          for (const { pos, node } of selectedParagraphs(state)) {
            const attrs = { ...(node.attrs as Record<string, unknown>) };
            attrs.alignment = patch.alignment;
            attrs.outlineLevel = patch.outlineLevel ?? undefined;
            attrs.indent = { ...((attrs.indent ?? {}) as object), ...patch.indent };
            attrs.spacing = { ...((attrs.spacing ?? {}) as object), ...patch.spacing };
            attrs.widowControl = patch.widowControl;
            attrs.keepNext = patch.keepNext;
            attrs.keepLines = patch.keepLines;
            attrs.pageBreakBefore = patch.pageBreakBefore;
            attrs.suppressLineNumbers = patch.suppressLineNumbers;
            attrs.suppressAutoHyphens = patch.suppressAutoHyphens;
            attrs.kinsoku = patch.kinsoku;
            attrs.wordWrap = patch.wordWrap;
            attrs.overflowPunctuation = patch.overflowPunct;
            attrs.autoSpaceDE = patch.autoSpaceDE;
            attrs.autoSpaceEastAsianText = patch.autoSpaceDN;
            attrs.textAlignment = patch.textAlignment;
            tr.setNodeMarkup(pos, undefined, attrs);
            touched = true;
          }
          return touched;
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
      // The Borders and Shading dialog's OK (border tab) — replaces each
      // selected paragraph's w:pBdr wholesale with the staged sides (a null
      // edge clears that side; every edge null drops the border).
      "borders-apply":
        (patch) =>
        ({ state, tr }) => {
          if (!patch?.sides) return false;
          const blocks = selectedParagraphs(state);
          if (!blocks.length) return false;
          for (const { pos, node } of blocks) {
            const attrs = node.attrs as Record<string, unknown>;
            const current = { ...((attrs.border ?? {}) as Record<string, unknown>) };
            for (const side of BORDER_SIDES) {
              const edge = patch.sides[side];
              if (!edge) delete current[side];
              else
                current[side] = {
                  style: edge.style,
                  size: Math.max(2, Math.round(edge.size)),
                  color: edge.color ?? "auto",
                  space: 0,
                };
            }
            const border = Object.keys(current).length ? current : null;
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
        ({ state, commands }) => {
          const anchor = tableAncestry(state);
          if (anchor && anchor.tableAt >= 0) {
            return (
              (commands as unknown as Record<string, () => boolean>)["split-table"]?.() ?? false
            );
          }
          return commands.setColumnBreak();
        },
      "section-break":
        () =>
        ({ commands }) =>
          commands.setSectionBreak(),
      // The Breaks menu's typed section breaks — Next Page (the plain
      // section-break command) and Continuous (flows on the same page).
      "section-break-next":
        () =>
        ({ commands }) =>
          commands.setSectionBreak(),
      "section-break-continuous":
        () =>
        ({ commands }) =>
          commands.setSectionBreak({ type: "continuous" }),
      // Insert a 3×3 table (Word's default Insert > Table preset). The header
      // row is the row-level tblHeader attr (w:tblHeader) — no header-cell
      // node type exists; every cell is a plain tableCell. Borders stamp
      // Word's "Table Grid" — 0.5pt single lines everywhere (w:sz is eighths
      // of a point, 4 = 0.5pt) — so the table is visible without a TableGrid
      // style in the document's styles.xml.
      "insert-table":
        (options) =>
        ({ state, dispatch }) => {
          const rows = Math.max(1, Math.min(50, Math.trunc(options?.rows ?? 3)));
          const cols = Math.max(1, Math.min(10, Math.trunc(options?.cols ?? 3)));
          const { table, tableRow, tableCell, paragraph } = state.schema.nodes;
          const cell = tableCell.createAndFill(null, [paragraph.create()]);
          if (!cell) return false;
          // The shape is structurally valid by construction (cols cells in a
          // "tableCell+" row), so the fill can only fail on a schema drift.
          // Cells/rows are immutable PM nodes — one instance is shared.
          const headerRow = tableRow.createAndFill({ tableHeader: true }, Array(cols).fill(cell))!;
          const dataRow = tableRow.createAndFill(null, Array(cols).fill(cell))!;
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
            [headerRow, ...Array(rows - 1).fill(dataRow)],
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

      // Insert an empty row copying the current row's structure and cell formatting.
      "insert-row-above":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.rowAt < 0) return false;
          if (dispatch) {
            const { $from } = state.selection;
            const row = $from.node(anchor.rowAt);
            const emptyCells: PMNode[] = [];
            row.forEach((cell) => {
              const para = state.schema.nodes.paragraph.create();
              emptyCells.push(cell.type.createAndFill(cell.attrs, [para])!);
            });
            const newRow = row.type.create(row.attrs, emptyCells);
            const insertPos = $from.before(anchor.rowAt);
            const tr = state.tr.insert(insertPos, newRow);
            tr.setSelection(TextSelection.near(tr.doc.resolve(insertPos + 2)));
            dispatch(tr.scrollIntoView());
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
            const emptyCells: PMNode[] = [];
            row.forEach((cell) => {
              const para = state.schema.nodes.paragraph.create();
              emptyCells.push(cell.type.createAndFill(cell.attrs, [para])!);
            });
            const newRow = row.type.create(row.attrs, emptyCells);
            const insertPos = $from.after(anchor.rowAt);
            const tr = state.tr.insert(insertPos, newRow);
            tr.setSelection(TextSelection.near(tr.doc.resolve(insertPos + 2)));
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Insert an empty row at a specific row index (0..nRows).
      "insert-row-at":
        (rowIndex: number, targetTablePos?: number) =>
        ({ state, dispatch }) => {
          let tablePos = targetTablePos;
          let tableNode: PMNode | null = null;
          if (tablePos != null) {
            tableNode = state.doc.nodeAt(tablePos);
          } else {
            const anchor = tableAncestry(state);
            if (!anchor) return false;
            tablePos = state.selection.$from.before(anchor.tableAt);
            tableNode = state.selection.$from.node(anchor.tableAt);
          }
          if (!tableNode || tablePos == null || tableNode.childCount === 0) return false;
          if (dispatch) {
            const tr = state.tr;
            const templateIdx = Math.min(rowIndex > 0 ? rowIndex - 1 : 0, tableNode.childCount - 1);
            const templateRow = tableNode.child(templateIdx);
            const emptyCells: PMNode[] = [];
            templateRow.forEach((cell) => {
              const para = state.schema.nodes.paragraph.create();
              emptyCells.push(cell.type.createAndFill(cell.attrs, [para])!);
            });
            const newRow = templateRow.type.create(templateRow.attrs, emptyCells);
            let insertPos = tablePos + 1;
            for (let r = 0; r < Math.min(rowIndex, tableNode.childCount); r += 1) {
              insertPos += tableNode.child(r).nodeSize;
            }
            tr.insert(insertPos, newRow);
            tr.setSelection(TextSelection.near(tr.doc.resolve(insertPos + 2)));
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // One empty cell per row, copied from each row's cell attrs at the current column
      // index. Bottom-up keeps positions valid as earlier edits shift later ones.
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
            let targetCellPos = -1;
            for (let r = tableNode.childCount - 1; r >= 0; r -= 1) {
              const rowNode = tableNode.child(r);
              let rowPos = tablePos + 1;
              for (let i = 0; i < r; i += 1) rowPos += tableNode.child(i).nodeSize;
              const idx = Math.min(cellIndex, rowNode.childCount - 1);
              let cellPos = rowPos + 1;
              for (let c = 0; c <= idx; c += 1) cellPos += rowNode.child(c).nodeSize;
              const template = rowNode.child(idx);
              const para = state.schema.nodes.paragraph.create();
              const emptyCell = template.type.createAndFill(template.attrs, [para])!;
              tr.insert(cellPos, emptyCell);
              if (r === $from.index(anchor.rowAt)) {
                targetCellPos = cellPos;
              }
            }
            if (Array.isArray(tableNode.attrs.columnWidths)) {
              const widths = [...(tableNode.attrs.columnWidths as number[])];
              const w = widths[cellIndex] ?? 2880;
              widths.splice(cellIndex + 1, 0, w);
              tr.setNodeMarkup(tablePos, undefined, { ...tableNode.attrs, columnWidths: widths });
            }
            if (targetCellPos > 0) {
              const sel =
                Selection.findFrom(tr.doc.resolve(targetCellPos + 1), 1, true) ??
                TextSelection.near(tr.doc.resolve(targetCellPos + 2));
              tr.setSelection(sel);
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
            let targetCellPos = -1;
            for (let r = tableNode.childCount - 1; r >= 0; r -= 1) {
              const rowNode = tableNode.child(r);
              let rowPos = tablePos + 1;
              for (let i = 0; i < r; i += 1) rowPos += tableNode.child(i).nodeSize;
              const idx = Math.min(cellIndex, rowNode.childCount - 1);
              let cellPos = rowPos + 1;
              for (let c = 0; c < idx; c += 1) cellPos += rowNode.child(c).nodeSize;
              const template = rowNode.child(idx);
              const para = state.schema.nodes.paragraph.create();
              const emptyCell = template.type.createAndFill(template.attrs, [para])!;
              tr.insert(cellPos, emptyCell);
              if (r === $from.index(anchor.rowAt)) {
                targetCellPos = cellPos;
              }
            }
            if (Array.isArray(tableNode.attrs.columnWidths)) {
              const widths = [...(tableNode.attrs.columnWidths as number[])];
              const w = widths[cellIndex] ?? 2880;
              widths.splice(cellIndex, 0, w);
              tr.setNodeMarkup(tablePos, undefined, { ...tableNode.attrs, columnWidths: widths });
            }
            if (targetCellPos > 0) {
              const sel =
                Selection.findFrom(tr.doc.resolve(targetCellPos + 1), 1, true) ??
                TextSelection.near(tr.doc.resolve(targetCellPos + 2));
              tr.setSelection(sel);
            }
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Insert a column at a specific column index (0..nCols).
      "insert-column-at":
        (colIndex: number, targetTablePos?: number) =>
        ({ state, dispatch }) => {
          let tablePos = targetTablePos;
          let tableNode: PMNode | null = null;
          if (tablePos != null) {
            tableNode = state.doc.nodeAt(tablePos);
          } else {
            const anchor = tableAncestry(state);
            if (!anchor) return false;
            tablePos = state.selection.$from.before(anchor.tableAt);
            tableNode = state.selection.$from.node(anchor.tableAt);
          }
          if (!tableNode || tablePos == null || tableNode.childCount === 0) return false;
          if (dispatch) {
            const tr = state.tr;
            let targetCellPos = -1;
            for (let r = tableNode.childCount - 1; r >= 0; r -= 1) {
              const rowNode = tableNode.child(r);
              let rowPos = tablePos + 1;
              for (let i = 0; i < r; i += 1) rowPos += tableNode.child(i).nodeSize;
              const idx = Math.min(colIndex, rowNode.childCount);
              let cellPos = rowPos + 1;
              for (let c = 0; c < idx; c += 1) cellPos += rowNode.child(c).nodeSize;
              const templateIdx = Math.min(colIndex > 0 ? colIndex - 1 : 0, rowNode.childCount - 1);
              const template = rowNode.child(templateIdx);
              const para = state.schema.nodes.paragraph.create();
              const emptyCell = template.type.createAndFill(template.attrs, [para])!;
              tr.insert(cellPos, emptyCell);
              if (r === 0) targetCellPos = cellPos;
            }
            if (Array.isArray(tableNode.attrs.columnWidths)) {
              const widths = [...(tableNode.attrs.columnWidths as number[])];
              const w = widths[colIndex - 1] ?? widths[colIndex] ?? 2880;
              widths.splice(colIndex, 0, w);
              tr.setNodeMarkup(tablePos, undefined, { ...tableNode.attrs, columnWidths: widths });
            }
            if (targetCellPos > 0) {
              const sel =
                Selection.findFrom(tr.doc.resolve(targetCellPos + 1), 1, true) ??
                TextSelection.near(tr.doc.resolve(targetCellPos + 2));
              tr.setSelection(sel);
            }
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Set table column widths directly on the table node.
      "set-table-column-widths":
        (widths: number[], targetTablePos?: number) =>
        ({ state, dispatch }) => {
          let tablePos = targetTablePos;
          let tableNode: PMNode | null = null;
          if (tablePos != null) {
            tableNode = state.doc.nodeAt(tablePos);
          } else {
            const anchor = tableAncestry(state);
            if (!anchor) return false;
            tablePos = state.selection.$from.before(anchor.tableAt);
            tableNode = state.selection.$from.node(anchor.tableAt);
          }
          if (!tableNode || tablePos == null) return false;
          if (dispatch) {
            dispatch(
              state.tr
                .setNodeMarkup(tablePos, undefined, {
                  ...tableNode.attrs,
                  columnWidths: widths,
                })
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Set height on a specific table row.
      "set-table-row-height":
        (
          rowIndex: number,
          height: { rule: "atLeast" | "exact"; value: number } | null,
          targetTablePos?: number,
        ) =>
        ({ state, dispatch }) => {
          let tablePos = targetTablePos;
          let tableNode: PMNode | null = null;
          if (tablePos != null) {
            tableNode = state.doc.nodeAt(tablePos);
          } else {
            const anchor = tableAncestry(state);
            if (!anchor) return false;
            tablePos = state.selection.$from.before(anchor.tableAt);
            tableNode = state.selection.$from.node(anchor.tableAt);
          }
          if (!tableNode || tablePos == null || rowIndex >= tableNode.childCount) return false;
          if (dispatch) {
            let rowPos = tablePos + 1;
            for (let r = 0; r < rowIndex; r += 1) rowPos += tableNode.child(r).nodeSize;
            const rowNode = state.doc.nodeAt(rowPos)!;
            dispatch(
              state.tr
                .setNodeMarkup(rowPos, undefined, {
                  ...rowNode.attrs,
                  height,
                })
                .scrollIntoView(),
            );
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
            if (Array.isArray(tableNode.attrs.columnWidths)) {
              const widths = [...(tableNode.attrs.columnWidths as number[])];
              if (cellIndex < widths.length) {
                widths.splice(cellIndex, 1);
                tr.setNodeMarkup(tablePos, undefined, { ...tableNode.attrs, columnWidths: widths });
              }
            }
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      "select-table":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.cellAt < 0) return false;
          if (dispatch) {
            // Every cell whole (Word's corner-handle pick) — a NodeSelection
            // over the table node would make Backspace erase the table.
            const cellPos = state.selection.$from.before(anchor.cellAt);
            dispatch(
              state.tr
                .setSelection(CellSelection.tableSelection(state.doc.resolve(cellPos)) as never)
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Word's row pick: a cell selection over the caret's whole row — the
      // same shape a bar-arrow click or a cross-cell drag produces, so
      // delete-row / merge-cells downstream see one selection model.
      "select-table-row":
        () =>
        ({ state, dispatch }) => {
          let cellPos: number;
          if (state.selection instanceof CellSelection) {
            cellPos = state.selection.anchorCell;
          } else {
            const anchor = tableAncestry(state);
            if (!anchor || anchor.cellAt < 0) return false;
            cellPos = state.selection.$from.before(anchor.cellAt);
          }
          if (dispatch) {
            dispatch(
              state.tr
                .setSelection(CellSelection.rowSelection(state.doc.resolve(cellPos)) as never)
                .scrollIntoView(),
            );
          }
          return true;
        },
      "select-table-cell":
        () =>
        ({ state, dispatch }) => {
          let cellPos: number;
          if (state.selection instanceof CellSelection) {
            cellPos = state.selection.anchorCell;
          } else {
            const anchor = tableAncestry(state);
            if (!anchor || anchor.cellAt < 0) return false;
            cellPos = state.selection.$from.before(anchor.cellAt);
          }
          if (dispatch) {
            dispatch(
              state.tr
                .setSelection(new CellSelection(state.doc.resolve(cellPos)) as never)
                .scrollIntoView(),
            );
          }
          return true;
        },
      // The caret's column across all rows: a cell selection over the whole
      // column (span-aware grid math comes free from prosemirror-tables).
      "select-table-column":
        () =>
        ({ state, dispatch }) => {
          let cellPos: number;
          if (state.selection instanceof CellSelection) {
            cellPos = state.selection.anchorCell;
          } else {
            const anchor = tableAncestry(state);
            if (!anchor || anchor.cellAt < 0) return false;
            cellPos = state.selection.$from.before(anchor.cellAt);
          }
          if (dispatch) {
            dispatch(
              state.tr
                .setSelection(CellSelection.colSelection(state.doc.resolve(cellPos)) as never)
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
          // No value = the split's primary face — Word defaults it to
          // middle-center (what the button's icon shows).
          const spec = CELL_ALIGN[value ?? "mc"];
          if (!spec) return false;
          const targets = tableTargets(state);
          if (!targets) return false;
          if (dispatch) {
            const { paragraph } = state.schema.nodes;
            const tr = state.tr;
            for (const { pos, node: cell } of targets.cells) {
              tr.setNodeMarkup(pos, undefined, {
                ...cell.attrs,
                verticalAlign: spec.v,
              });
              state.doc.nodesBetween(pos + 1, pos + cell.nodeSize - 1, (node, npos) => {
                if (node.type === paragraph && node.attrs.alignment !== spec.h) {
                  tr.setNodeMarkup(npos, undefined, { ...node.attrs, alignment: spec.h });
                }
                return true;
              });
            }
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Word's Cell Margins presets (Table Layout): "default" clears the cell's
      // tcMar so the table default applies again; the named presets stamp
      // their twip insets on the caret's cell.
      "cell-margins":
        (value) =>
        ({ state, dispatch }) => {
          if (typeof value !== "string" || !CELL_MARGIN_PRESETS.hasOwnProperty(value)) return false;
          const targets = tableTargets(state);
          if (!targets) return false;
          if (dispatch) {
            const tr = state.tr;
            for (const { pos, node: cell } of targets.cells) {
              tr.setNodeMarkup(pos, undefined, {
                ...cell.attrs,
                margins: CELL_MARGIN_PRESETS[value],
              });
            }
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      "set-cell-insets":
        (value) =>
        ({ state, dispatch }) => {
          const parsed = parseCellInsets(value);
          if (parsed === undefined) return false;
          const targets = tableTargets(state);
          if (!targets) return false;
          if (dispatch) {
            const tr = state.tr;
            for (const { pos, node: cell } of targets.cells) {
              tr.setNodeMarkup(pos, undefined, {
                ...cell.attrs,
                margins: parsed,
              });
            }
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      "set-cell-vertical-align":
        (value) =>
        ({ state, dispatch }) => {
          if (value !== "top" && value !== "center" && value !== "bottom" && value !== null) {
            return false;
          }
          const targets = tableTargets(state);
          if (!targets) return false;
          if (dispatch) {
            const tr = state.tr;
            for (const { pos, node: cell } of targets.cells) {
              tr.setNodeMarkup(pos, undefined, {
                ...cell.attrs,
                verticalAlign: value,
              });
            }
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Word's Repeat Header Rows — marks the whole pick (the selected rows
      // gain or lose tblHeader together, either to an explicit boolean or
      // toggling based on the anchor row's current state).
      "repeat-header-rows":
        (value) =>
        ({ state, dispatch }) => {
          const targets = tableTargets(state);
          if (!targets) return false;
          if (dispatch) {
            const firstRow = Math.min(...targets.rows);
            const current = !!(targets.tableNode.child(firstRow)!.attrs.tableHeader as boolean);
            const next = typeof value === "boolean" ? value : !current;
            const tr = state.tr;
            stampRows(tr, targets, (row) => ({
              ...row.attrs,
              tableHeader: next,
            }));
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Word's CantSplit / Row pagination protection — marks the whole pick
      // (the selected rows gain or lose cantSplit together, either to an
      // explicit boolean or toggling based on the anchor row's current state).
      "cant-split":
        (value) =>
        ({ state, dispatch }) => {
          const targets = tableTargets(state);
          if (!targets) return false;
          if (dispatch) {
            const firstRow = Math.min(...targets.rows);
            const current = !!(targets.tableNode.child(firstRow)!.attrs.cantSplit as boolean);
            const next = typeof value === "boolean" ? value : !current;
            const tr = state.tr;
            stampRows(tr, targets, (row) => ({
              ...row.attrs,
              cantSplit: next ? true : null,
            }));
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Cell-level shading (tcPr shd) — the Home shading button stays at
      // paragraph level; Word's Table Design shading is the cell property.
      "cell-shading":
        (value) =>
        ({ state, dispatch }) => {
          const stamp = shadingStamp(value);
          if (stamp === undefined) return false;
          const targets = tableTargets(state);
          if (!targets) return false;
          if (dispatch) {
            const tr = state.tr;
            for (const { pos, node: cell } of targets.cells) {
              tr.setNodeMarkup(pos, undefined, { ...cell.attrs, shading: stamp });
            }
            dispatch(tr.scrollIntoView());
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
      "cell-borders":
        (value) =>
        ({ state, dispatch }) => {
          return applyCellBorders(state, dispatch, value);
        },
      // Border-side presets on the table (value space matches the Home
      // paragraph-border menu). When CellSelection is active, delegates to
      // applyCellBorders to style the selected cells.
      "table-borders":
        (value) =>
        ({ state, dispatch }) => {
          if (typeof value !== "string") return false;
          if (state.selection instanceof CellSelection) {
            return applyCellBorders(state, dispatch, value);
          }
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          const current = (state.selection.$from.node(anchor.tableAt).attrs.borders ??
            null) as TableBordersLike | null;
          return stampTableBorders(state, dispatch, tableBordersStamp(value, current));
        },
      // Word's Text Direction button: one press turns the whole pick to one
      // direction (tbRl ↔ unset, the anchor cell's state deciding). The attr
      // round-trips through the docx engine; the canvas doesn't paint
      // vertical cell text yet.
      "text-direction":
        () =>
        ({ state, dispatch }) => {
          const targets = tableTargets(state);
          if (!targets) return false;
          if (dispatch) {
            const next = targets.cells[0]!.node.attrs.textDirection ? null : "tbRl";
            const tr = state.tr;
            for (const { pos, node: cell } of targets.cells) {
              tr.setNodeMarkup(pos, undefined, { ...cell.attrs, textDirection: next });
            }
            dispatch(tr.scrollIntoView());
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
      // Word's "Convert Text to Table": converts delimited text into a structured
      // table with automatic delimiter detection (tabs, commas, semicolons, pipes).
      "convert-text-to-table":
        (options) =>
        ({ state, dispatch }) => {
          if (tableAncestry(state)) return false;

          const opts: ConvertTextToTableOptions =
            typeof options === "string" ? { delimiter: options } : (options ?? {});

          const { from, to } = state.selection;
          const $from = state.doc.resolve(from);
          const $to = state.doc.resolve(to);

          const lines: string[] = [];
          let startPos: number;
          let endPos: number;

          if ($from.sameParent($to) && $from.parent.isTextblock) {
            startPos = $from.before();
            endPos = $from.after();
            const text = state.selection.empty
              ? $from.parent.textContent
              : state.doc.textBetween(from, to, "\n");
            for (const l of text.split(/\r?\n/)) {
              lines.push(l);
            }
          } else {
            const startIndex = $from.index(0);
            const endIndex = Math.min(
              state.doc.childCount,
              Math.max(startIndex + 1, $to.index(0) + ($to.parentOffset > 0 ? 1 : 0)),
            );

            let pos = 0;
            for (let i = 0; i < startIndex; i += 1) {
              pos += state.doc.child(i).nodeSize;
            }
            startPos = pos;
            for (let i = startIndex; i < endIndex; i += 1) {
              const child = state.doc.child(i);
              pos += child.nodeSize;
              const text = child.textContent;
              for (const l of text.split(/\r?\n/)) {
                lines.push(l);
              }
            }
            endPos = pos;
          }

          while (lines.length > 1 && !lines[lines.length - 1]!.trim()) {
            lines.pop();
          }
          if (!lines.length || (lines.length === 1 && !lines[0]!.trim())) {
            return false;
          }

          let delimiter = opts.delimiter;
          if (!delimiter) {
            if (lines.some((l) => l.includes("\t"))) {
              delimiter = "\t";
            } else if (lines.some((l) => l.includes(","))) {
              delimiter = ",";
            } else if (lines.some((l) => l.includes(";"))) {
              delimiter = ";";
            } else if (lines.some((l) => l.includes("|"))) {
              delimiter = "|";
            } else {
              delimiter = "\t";
            }
          }

          const rawRows = lines.map((line) => {
            if (delimiter === "|") {
              let trimmed = line.trim();
              if (trimmed.startsWith("|")) trimmed = trimmed.slice(1);
              if (trimmed.endsWith("|")) trimmed = trimmed.slice(0, -1);
              return trimmed.split("|").map((p) => p.trim());
            }
            if (typeof delimiter === "string") {
              return line.split(delimiter).map((p) => p.trim());
            }
            return line.split(delimiter).map((p) => p.trim());
          });

          const maxCols = Math.max(1, ...rawRows.map((r) => r.length));
          const { table, tableRow, tableCell, paragraph } = state.schema.nodes;
          const rows: PMNode[] = [];

          for (let r = 0; r < rawRows.length; r += 1) {
            const parts = rawRows[r]!;
            const cells: PMNode[] = [];
            for (let c = 0; c < maxCols; c += 1) {
              const text = parts[c] ?? "";
              const p = text ? paragraph.create(null, state.schema.text(text)) : paragraph.create();
              cells.push(tableCell.create(null, [p]));
            }
            const isHeader = Boolean(opts.hasHeader && r === 0);
            rows.push(tableRow.create(isHeader ? { tableHeader: true } : null, cells));
          }

          const tableNode = table.create(
            {
              borders: TABLE_GRID_BORDERS,
            },
            rows,
          );

          if (dispatch) {
            const tr = state.tr.replaceWith(startPos, endPos, tableNode);
            tr.setSelection(TextSelection.near(tr.doc.resolve(startPos + 2)));
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Word's Sort inside Table: sorts table rows by active or specified column,
      // preserving header rows (tableHeader: true) with natural locale and numeric comparison.
      "sort-table":
        (options) =>
        ({ state, dispatch }) => {
          const targets = tableTargets(state);
          if (!targets) return false;
          const { tableNode, tablePos } = targets;
          if (tableNode.childCount < 2) return false;

          const colIndex =
            typeof options === "object" && typeof options?.column === "number"
              ? options.column
              : (targets.cols.values().next().value ?? 0);
          const order =
            (typeof options === "string" ? options : options?.order) === "desc" ? "desc" : "asc";

          const row0 = tableNode.child(0);
          const hasHeader =
            typeof options === "object" && options?.hasHeader !== undefined
              ? Boolean(options.hasHeader)
              : Boolean(row0.attrs.tableHeader);

          const startRow = hasHeader ? 1 : 0;
          if (tableNode.childCount - startRow < 2) return false;

          const allRows: PMNode[] = [];
          tableNode.forEach((r) => allRows.push(r));

          const headerRows = allRows.slice(0, startRow);
          const dataRows = allRows.slice(startRow);

          const getCellText = (row: PMNode, targetCol: number): string => {
            let col = 0;
            for (let c = 0; c < row.childCount; c += 1) {
              const cell = row.child(c);
              const span = spanOf(cell);
              if (targetCol >= col && targetCol < col + span) {
                return cell.textContent.trim();
              }
              col += span;
            }
            return "";
          };

          const sortedRows = [...dataRows].sort((a, b) => {
            const textA = getCellText(a, colIndex);
            const textB = getCellText(b, colIndex);
            const cmp = textA.localeCompare(textB, undefined, {
              numeric: true,
              sensitivity: "base",
            });
            return order === "desc" ? -cmp : cmp;
          });

          const newRows = [...headerRows, ...sortedRows];
          if (dispatch) {
            const newTable = tableNode.type.create(tableNode.attrs, newRows);
            const tr = state.tr.replaceWith(tablePos, tablePos + tableNode.nodeSize, newTable);
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Word's Formula in Table: evaluates =SUM(ABOVE), =SUM(LEFT), =AVERAGE(...),
      // =COUNT(...), =MIN(...), =MAX(...), =PRODUCT(...) and writes formatted result into active cell.
      "table-formula":
        (options) =>
        ({ state, dispatch }) => {
          const targets = tableTargets(state);
          if (!targets) return false;
          const { tableNode, cells } = targets;
          const activeCellTarget = cells[0];
          if (!activeCellTarget) return false;

          const activeRow = activeCellTarget.row;
          const activeCol = activeCellTarget.col;
          const numRows = tableNode.childCount;

          const grid: (PMNode | null)[][] = [];
          for (let r = 0; r < numRows; r += 1) {
            const rowNode = tableNode.child(r);
            grid[r] = [];
            let col = 0;
            for (let c = 0; c < rowNode.childCount; c += 1) {
              const cellNode = rowNode.child(c);
              const span = spanOf(cellNode);
              for (let s = 0; s < span; s += 1) {
                grid[r]![col + s] = cellNode;
              }
              col += span;
            }
          }

          const extractNumber = (cell: PMNode | null | undefined): number | null => {
            if (!cell) return null;
            const text = cell.textContent.trim();
            if (!text) return null;
            const cleaned = text.replace(/[$€£¥, ]/g, "");
            if (cleaned.endsWith("%")) {
              const n = Number(cleaned.slice(0, -1));
              return Number.isFinite(n) ? n / 100 : null;
            }
            if (cleaned.startsWith("(") && cleaned.endsWith(")")) {
              const n = Number(cleaned.slice(1, -1));
              return Number.isFinite(n) ? -n : null;
            }
            const n = Number(cleaned);
            return Number.isFinite(n) ? n : null;
          };

          let formulaStr: string;
          if (typeof options === "string") {
            formulaStr = options;
          } else if (options?.formula) {
            formulaStr = options.formula;
          } else {
            let hasAbove = false;
            if (activeRow > 0 && extractNumber(grid[activeRow - 1]?.[activeCol]) !== null) {
              hasAbove = true;
            }
            let hasLeft = false;
            if (activeCol > 0 && extractNumber(grid[activeRow]?.[activeCol - 1]) !== null) {
              hasLeft = true;
            }
            if (hasAbove) {
              formulaStr = "=SUM(ABOVE)";
            } else if (hasLeft) {
              formulaStr = "=SUM(LEFT)";
            } else {
              formulaStr = "=SUM(ABOVE)";
            }
          }

          const rawExpr = formulaStr.trim().replace(/^=/, "").trim();

          const parseCoord = (coord: string): { r: number; c: number } | null => {
            const m = /^([A-Za-z]+)(\d+)$/.exec(coord.trim());
            if (!m) return null;
            const colLetters = m[1]!.toUpperCase();
            let c = 0;
            for (let i = 0; i < colLetters.length; i += 1) {
              c = c * 26 + (colLetters.charCodeAt(i) - 64);
            }
            c -= 1;
            const r = parseInt(m[2]!, 10) - 1;
            return { r, c };
          };

          const parseRange = (rangeStr: string): number[] => {
            const nums: number[] = [];
            const parts = rangeStr.split(",").map((p) => p.trim());
            for (const part of parts) {
              const upper = part.toUpperCase();
              if (upper === "ABOVE") {
                for (let r = activeRow - 1; r >= 0; r -= 1) {
                  const val = extractNumber(grid[r]?.[activeCol]);
                  if (val !== null) nums.push(val);
                  else break;
                }
              } else if (upper === "LEFT") {
                for (let c = activeCol - 1; c >= 0; c -= 1) {
                  const val = extractNumber(grid[activeRow]?.[c]);
                  if (val !== null) nums.push(val);
                  else break;
                }
              } else if (upper === "BELOW") {
                for (let r = activeRow + 1; r < numRows; r += 1) {
                  const val = extractNumber(grid[r]?.[activeCol]);
                  if (val !== null) nums.push(val);
                  else break;
                }
              } else if (upper === "RIGHT") {
                const maxC = grid[activeRow]?.length ?? 0;
                for (let c = activeCol + 1; c < maxC; c += 1) {
                  const val = extractNumber(grid[activeRow]?.[c]);
                  if (val !== null) nums.push(val);
                  else break;
                }
              } else if (upper.includes(":")) {
                const [fromStr, toStr] = upper.split(":");
                const fromCoord = parseCoord(fromStr!);
                const toCoord = parseCoord(toStr!);
                if (fromCoord && toCoord) {
                  const rMin = Math.min(fromCoord.r, toCoord.r);
                  const rMax = Math.max(fromCoord.r, toCoord.r);
                  const cMin = Math.min(fromCoord.c, toCoord.c);
                  const cMax = Math.max(fromCoord.c, toCoord.c);
                  for (let r = rMin; r <= rMax; r += 1) {
                    for (let c = cMin; c <= cMax; c += 1) {
                      const val = extractNumber(grid[r]?.[c]);
                      if (val !== null) nums.push(val);
                    }
                  }
                }
              } else {
                const coord = parseCoord(upper);
                if (coord) {
                  const val = extractNumber(grid[coord.r]?.[coord.c]);
                  if (val !== null) nums.push(val);
                }
              }
            }
            return nums;
          };

          let calculatedVal = 0;
          const fnMatch = /^(SUM|AVERAGE|COUNT|MIN|MAX|PRODUCT)\s*\((.*?)\)$/i.exec(rawExpr);
          if (fnMatch) {
            const fnName = fnMatch[1]!.toUpperCase();
            const args = fnMatch[2]!;
            const nums = parseRange(args);
            switch (fnName) {
              case "SUM":
                calculatedVal = nums.reduce((acc, v) => acc + v, 0);
                break;
              case "AVERAGE":
                calculatedVal = nums.length ? nums.reduce((acc, v) => acc + v, 0) / nums.length : 0;
                break;
              case "COUNT":
                calculatedVal = nums.length;
                break;
              case "MIN":
                calculatedVal = nums.length ? Math.min(...nums) : 0;
                break;
              case "MAX":
                calculatedVal = nums.length ? Math.max(...nums) : 0;
                break;
              case "PRODUCT":
                calculatedVal = nums.length ? nums.reduce((acc, v) => acc * v, 1) : 0;
                break;
            }
          } else {
            const evalExpr = rawExpr.replace(/[A-Za-z]+\d+/g, (match) => {
              const coord = parseCoord(match);
              if (!coord) return "0";
              const val = extractNumber(grid[coord.r]?.[coord.c]);
              return String(val ?? 0);
            });
            if (!/^[\d+\-*/.() ]+$/.test(evalExpr)) {
              return false;
            }
            const tokens = evalExpr.match(/\d+(?:\.\d+)?|[+\-*/()]/g) || [];
            let pos = 0;
            const parsePrimary = (): number => {
              const token = tokens[pos];
              if (token === "(") {
                pos += 1;
                const val = parseExpr();
                if (tokens[pos] === ")") pos += 1;
                return val;
              }
              if (token === "-") {
                pos += 1;
                return -parsePrimary();
              }
              if (token === "+") {
                pos += 1;
                return parsePrimary();
              }
              pos += 1;
              return Number(token) || 0;
            };
            const parseFactor = (): number => {
              let left = parsePrimary();
              while (pos < tokens.length && (tokens[pos] === "*" || tokens[pos] === "/")) {
                const op = tokens[pos++];
                const right = parsePrimary();
                if (op === "*") left *= right;
                else if (op === "/") left = right !== 0 ? left / right : 0;
              }
              return left;
            };
            const parseExpr = (): number => {
              let left = parseFactor();
              while (pos < tokens.length && (tokens[pos] === "+" || tokens[pos] === "-")) {
                const op = tokens[pos++];
                const right = parseFactor();
                if (op === "+") left += right;
                else if (op === "-") left -= right;
              }
              return left;
            };
            calculatedVal = parseExpr();
          }

          let resultText = String(calculatedVal);
          if (Number.isFinite(calculatedVal)) {
            if (Number.isInteger(calculatedVal)) {
              resultText = calculatedVal.toString();
            } else {
              resultText = Number(calculatedVal.toFixed(4)).toString();
            }
          }

          if (dispatch) {
            const cellNode = activeCellTarget.node;
            const p = state.schema.nodes.paragraph.create(
              null,
              resultText ? state.schema.text(resultText) : undefined,
            );
            const tr = state.tr.replaceWith(
              activeCellTarget.pos + 1,
              activeCellTarget.pos + cellNode.nodeSize - 1,
              p,
            );
            tr.setSelection(TextSelection.near(tr.doc.resolve(activeCellTarget.pos + 2)));
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
      // Word's Merge Cells over the selection's bounding rectangle.
      // Rebuilds merged rows with proper OOXML columnSpan and verticalMerge ("restart" / "continue"),
      // while preserving all paragraphs and blocks from every merged cell in the master cell.
      "merge-cells":
        () =>
        ({ state, dispatch }) => {
          let $from = state.selection.$from;
          let $to = state.selection.$to;
          if (state.selection instanceof CellSelection) {
            const sel = state.selection as unknown as CellSelection;
            const $a = state.doc.resolve(sel.anchorCell + 2);
            const $h = state.doc.resolve(sel.headCell + 2);
            $from = $a.pos <= $h.pos ? $a : $h;
            $to = $a.pos <= $h.pos ? $h : $a;
          }
          const fromA = ancestryAt($from);
          const toA = ancestryAt($to);
          if (!fromA || !toA || fromA.rowAt < 0 || toA.rowAt < 0) return false;
          if ($from.before(fromA.tableAt) !== $to.before(toA.tableAt)) return false;

          const tablePos = $from.before(fromA.tableAt);
          const tableNode = $from.node(fromA.tableAt);

          const rA = { row: $from.index(fromA.tableAt), cell: $from.index(fromA.rowAt) };
          const rH = { row: $to.index(toA.tableAt), cell: $to.index(toA.rowAt) };
          const rowFrom = Math.min(rA.row, rH.row);
          const rowTo = Math.max(rA.row, rH.row);
          const c1 = Math.min(rA.cell, rH.cell);
          const c2 = Math.max(rA.cell, rH.cell);
          if (rowFrom === rowTo && c1 === c2) return false;

          const anchorRowNode = tableNode.child(rA.row);
          const headRowNode = tableNode.child(rH.row);
          const colFrom = Math.min(
            gridColOf(anchorRowNode, rA.cell),
            gridColOf(headRowNode, rH.cell),
          );
          const colTo = Math.max(
            gridColOf(anchorRowNode, rA.cell) + spanOf(anchorRowNode.child(rA.cell)!),
            gridColOf(headRowNode, rH.cell) + spanOf(headRowNode.child(rH.cell)!),
          );

          const totalCols = colTo - colFrom;
          const totalRows = rowTo - rowFrom + 1;
          if (totalCols <= 1 && totalRows <= 1) return false;

          if (dispatch) {
            // 1. Collect all non-empty block content from all merged cells in row-major order
            const collectedBlocks: PMNode[] = [];
            for (let r = rowFrom; r <= rowTo; r += 1) {
              const rowNode = tableNode.child(r);
              let cCol = 0;
              for (let c = 0; c < rowNode.childCount; c += 1) {
                const cell = rowNode.child(c);
                const span = spanOf(cell);
                const cEnd = cCol + span;
                if (cCol < colTo && cEnd > colFrom) {
                  cell.forEach((child) => {
                    if (child.isTextblock && child.content.size === 0) return;
                    collectedBlocks.push(child);
                  });
                }
                cCol = cEnd;
              }
            }
            if (collectedBlocks.length === 0) {
              collectedBlocks.push(state.schema.nodes.paragraph.create());
            }

            // 2. Rebuild rows bottom-up so document positions remain stable
            const tr = state.tr;
            for (let r = rowTo; r >= rowFrom; r -= 1) {
              const rowNode = tableNode.child(r);
              let rowPos = tablePos + 1;
              for (let i = 0; i < r; i += 1) rowPos += tableNode.child(i).nodeSize;

              const newCells: PMNode[] = [];
              let cCol = 0;
              let mergedInserted = false;

              for (let c = 0; c < rowNode.childCount; c += 1) {
                const cell = rowNode.child(c);
                const span = spanOf(cell);
                const cEnd = cCol + span;

                if (cEnd <= colFrom || cCol >= colTo) {
                  newCells.push(cell);
                } else {
                  if (!mergedInserted) {
                    mergedInserted = true;
                    if (r === rowFrom) {
                      const masterAttrs = {
                        ...cell.attrs,
                        columnSpan: totalCols > 1 ? totalCols : null,
                        verticalMerge: null,
                      };
                      newCells.push(cell.type.create(masterAttrs, collectedBlocks));
                    } else {
                      const contAttrs = {
                        ...cell.attrs,
                        columnSpan: totalCols > 1 ? totalCols : null,
                        verticalMerge: "continue",
                      };
                      newCells.push(
                        cell.type.create(contAttrs, [state.schema.nodes.paragraph.create()]),
                      );
                    }
                  }
                }
                cCol = cEnd;
              }

              const newRow = rowNode.type.create(rowNode.attrs, newCells);
              tr.replaceWith(rowPos, rowPos + rowNode.nodeSize, newRow);
            }

            // 3. Set selection in the master cell
            let masterPos = tablePos + 1;
            for (let i = 0; i < rowFrom; i += 1) {
              masterPos += tr.doc.nodeAt(tablePos)!.child(i).nodeSize;
            }
            let cellPos = masterPos + 1;
            const updatedRow = tr.doc.nodeAt(tablePos)!.child(rowFrom);
            let curCol = 0;
            for (let c = 0; c < updatedRow.childCount; c += 1) {
              const cell = updatedRow.child(c);
              if (curCol === colFrom) break;
              curCol += spanOf(cell);
              cellPos += cell.nodeSize;
            }
            const sel = Selection.findFrom(tr.doc.resolve(cellPos + 1), 1, true);
            if (sel) tr.setSelection(sel);
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Word's Split Cells: splits horizontally merged (columnSpan > 1) and/or
      // vertically merged ("restart" / "continue") cells back to individual grid cells.
      "split-cell":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.cellAt < 0 || anchor.rowAt < 0 || anchor.tableAt < 0) return false;
          const { $from } = state.selection;
          const tableNode = $from.node(anchor.tableAt);
          const tablePos = $from.before(anchor.tableAt);
          const cell = $from.node(anchor.cellAt);
          const span = spanOf(cell);
          const vMerge = cell.attrs.verticalMerge as string | null | undefined;

          const curRowIdx = $from.index(anchor.tableAt);
          const curCellIdx = $from.index(anchor.rowAt);
          const targetCol = gridColOf(tableNode.child(curRowIdx), curCellIdx);

          let hasContinueBelow = false;
          if (curRowIdx + 1 < tableNode.childCount) {
            const nextRow = tableNode.child(curRowIdx + 1);
            let cCol = 0;
            for (let c = 0; c < nextRow.childCount; c += 1) {
              const cNode = nextRow.child(c);
              if (cCol === targetCol && cNode.attrs.verticalMerge === "continue") {
                hasContinueBelow = true;
                break;
              }
              cCol += spanOf(cNode);
            }
          }

          const isVerticalMerge = vMerge === "restart" || vMerge === "continue" || hasContinueBelow;

          if (span <= 1 && !isVerticalMerge) return false;

          if (dispatch) {
            const tr = state.tr;

            if (isVerticalMerge) {
              let restartRowIdx = curRowIdx;
              if (vMerge === "continue") {
                for (let r = curRowIdx - 1; r >= 0; r -= 1) {
                  const row = tableNode.child(r);
                  let col = 0;
                  for (let c = 0; c < row.childCount; c += 1) {
                    const cNode = row.child(c);
                    if (col === targetCol) {
                      if (cNode.attrs.verticalMerge !== "continue") {
                        restartRowIdx = r;
                      }
                      break;
                    }
                    col += spanOf(cNode);
                  }
                  if (restartRowIdx !== curRowIdx) break;
                }
              }

              const affectedRowIndices: number[] = [restartRowIdx];
              for (let r = restartRowIdx + 1; r < tableNode.childCount; r += 1) {
                const row = tableNode.child(r);
                let col = 0;
                let isContinue = false;
                for (let c = 0; c < row.childCount; c += 1) {
                  const cNode = row.child(c);
                  if (col === targetCol) {
                    if (cNode.attrs.verticalMerge === "continue") {
                      isContinue = true;
                    }
                    break;
                  }
                  col += spanOf(cNode);
                }
                if (isContinue) {
                  affectedRowIndices.push(r);
                } else {
                  break;
                }
              }

              for (let i = affectedRowIndices.length - 1; i >= 0; i -= 1) {
                const r = affectedRowIndices[i]!;
                const rowNode = tableNode.child(r);
                let rowPos = tablePos + 1;
                for (let j = 0; j < r; j += 1) rowPos += tableNode.child(j).nodeSize;

                const newCells: PMNode[] = [];
                let col = 0;
                for (let c = 0; c < rowNode.childCount; c += 1) {
                  const cNode = rowNode.child(c);
                  const cSpan = spanOf(cNode);
                  if (col === targetCol) {
                    const unmergedAttrs = {
                      ...cNode.attrs,
                      columnSpan: null,
                      verticalMerge: null,
                    };
                    newCells.push(cNode.type.create(unmergedAttrs, cNode.content));
                    for (let s = 1; s < cSpan; s += 1) {
                      newCells.push(
                        cNode.type.create(
                          { ...cNode.attrs, columnSpan: null, verticalMerge: null },
                          [state.schema.nodes.paragraph.create()],
                        ),
                      );
                    }
                  } else {
                    newCells.push(cNode);
                  }
                  col += cSpan;
                }
                const newRow = rowNode.type.create(rowNode.attrs, newCells);
                tr.replaceWith(rowPos, rowPos + rowNode.nodeSize, newRow);
              }
            } else if (span > 1) {
              const cellPos = $from.before(anchor.cellAt);
              tr.setNodeMarkup(cellPos, undefined, {
                ...cell.attrs,
                columnSpan: null,
                verticalMerge: null,
              });
              const blank = cell.type.create(null, [state.schema.nodes.paragraph.create()]);
              for (let i = 0; i < span - 1; i += 1) {
                tr.insert(cellPos + cell.nodeSize, blank);
              }
            }
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Word's Split Table: the caret's row starts a second table with the
      // same formatting (attrs are shared — borders, grid, style), separated
      // by a blank paragraph so the two tables don't merge back in Word.
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
            const sep = state.schema.nodes.paragraph.create();
            const tr = state.tr.replaceWith(tablePos, tablePos + tableNode.nodeSize, [
              create(tableNode.attrs, rowsA),
              sep,
              create(tableNode.attrs, rowsB),
            ]);
            // The caret lands in the second table's first cell: the first
            // table's size is its rows plus the open/close tokens, plus the separator paragraph.
            const firstTableSize = rowsA.reduce((sum, r) => sum + r.nodeSize, 0) + 2;
            tr.setSelection(
              TextSelection.near(tr.doc.resolve(tablePos + firstTableSize + sep.nodeSize + 1)),
            );
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Word's AutoFit Contents: each column shrinks to its widest cell's
      // content (a character-count heuristic — see measureTextTwip) without
      // growing past the current grid. Supports tables with merged cells.
      "autofit-contents":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          const { $from } = state.selection;
          const tableNode = $from.node(anchor.tableAt);
          const cols = tableGridColumns(tableNode);
          if (cols === 0) return false;

          const existingWidths = (tableNode.attrs.columnWidths as number[] | null) ?? [];

          // Compute widest content for each grid column
          const colWidest = Array.from<number>({ length: cols }).fill(0);
          for (let r = 0; r < tableNode.childCount; r += 1) {
            const row = tableNode.child(r);
            let col = 0;
            for (let c = 0; c < row.childCount; c += 1) {
              const cell = row.child(c);
              const span = spanOf(cell);
              if (cell.attrs.verticalMerge === "continue") {
                col += span;
                continue;
              }
              const textLen = measureTextTwip(cell.textContent);
              if (span === 1) {
                if (col < cols) {
                  colWidest[col] = Math.max(colWidest[col], textLen);
                }
              } else {
                const perCol = Math.round(textLen / span);
                for (let s = 0; s < span && col + s < cols; s += 1) {
                  colWidest[col + s] = Math.max(colWidest[col + s], perCol);
                }
              }
              col += span;
            }
          }

          if (dispatch) {
            const next = colWidest.map((widest, c) => {
              const currentW = existingWidths[c] ?? 2880;
              return Math.max(MIN_COL_TWIP, Math.min(currentW, widest));
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
      // text width (or an explicitly provided twip value). A table without a
      // grid starts from equal columns.
      "autofit-window":
        (value) =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          let total: number;
          if (value !== undefined) {
            total = Number(value);
            if (!Number.isFinite(total) || total <= 0) return false;
          } else {
            total = availableTextWidthTwips(state);
          }
          const { $from } = state.selection;
          const tableNode = $from.node(anchor.tableAt);
          const widths =
            (tableNode.attrs.columnWidths as number[] | null)?.filter((w) => w > 0) ?? [];
          const cols = Math.max(widths.length, tableGridColumns(tableNode));
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
      // Word's Distribute Columns: the grid splits its width evenly. When a
      // CellSelection spans multiple columns, distributes only the selected
      // columns; otherwise distributes all columns table-wide.
      "distribute-columns":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          const { $from } = state.selection;
          const tableNode = $from.node(anchor.tableAt);
          const totalCols = tableGridColumns(tableNode);
          if (totalCols === 0) return false;

          let widths = tableNode.attrs.columnWidths as number[] | null;
          if (!widths || widths.length !== totalCols) {
            widths = Array.from({ length: totalCols }, () => 2880);
          }

          const targets = tableTargets(state);
          const isCell = state.selection instanceof CellSelection;
          const hasSubsetCols =
            isCell && targets && targets.cols.size > 1 && targets.cols.size < totalCols;

          if (dispatch) {
            const next = [...widths];
            if (hasSubsetCols && targets) {
              const selectedCols = Array.from(targets.cols).sort((a, b) => a - b);
              const sum = selectedCols.reduce((acc, c) => acc + (widths![c] ?? 2880), 0);
              const even = Math.floor(sum / selectedCols.length);
              selectedCols.forEach((c, i) => {
                next[c] =
                  i === selectedCols.length - 1 ? sum - even * (selectedCols.length - 1) : even;
              });
            } else {
              const sum = widths.reduce((a, b) => a + b, 0);
              const even = Math.floor(sum / totalCols);
              for (let c = 0; c < totalCols; c += 1) {
                next[c] = c === totalCols - 1 ? sum - even * (totalCols - 1) : even;
              }
            }
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
      // Word's Distribute Rows: the targeted rows' heights split their total evenly.
      // If a CellSelection spans multiple rows, distributes only those rows;
      // otherwise distributes all rows in the table. Rows without a declared
      // height use an appropriate default (~400 twips).
      "distribute-rows":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          const { $from } = state.selection;
          const tableNode = $from.node(anchor.tableAt);
          const tablePos = $from.before(anchor.tableAt);
          const totalRows = tableNode.childCount;
          if (totalRows < 2) return false;

          const targets = tableTargets(state);
          const isCell = state.selection instanceof CellSelection;
          const selectedRowIndices =
            isCell && targets && targets.rows.size > 1
              ? Array.from(targets.rows).sort((a, b) => a - b)
              : Array.from({ length: totalRows }, (_, i) => i);

          if (selectedRowIndices.length < 2) return false;

          const DEFAULT_ROW_HEIGHT = 400;
          let declaredCount = 0;
          let declaredSum = 0;

          const rowInfos: { pos: number; rowIdx: number; height: number }[] = [];
          let curPos = tablePos + 1;
          for (let r = 0; r < totalRows; r += 1) {
            const row = tableNode.child(r);
            if (selectedRowIndices.includes(r)) {
              const h = row.attrs.height as { value?: unknown } | null;
              const hasVal = h && typeof h.value === "number" && h.value > 0;
              const val = hasVal ? (h.value as number) : DEFAULT_ROW_HEIGHT;
              if (hasVal) {
                declaredCount += 1;
                declaredSum += val;
              }
              rowInfos.push({ pos: curPos, rowIdx: r, height: val });
            }
            curPos += row.nodeSize;
          }

          if (dispatch) {
            const totalSum =
              declaredCount > 0
                ? declaredSum + (rowInfos.length - declaredCount) * DEFAULT_ROW_HEIGHT
                : rowInfos.length * DEFAULT_ROW_HEIGHT;
            const even = Math.floor(totalSum / rowInfos.length);
            const tr = state.tr;

            rowInfos.forEach(({ pos }, i) => {
              const row = tr.doc.nodeAt(pos)!;
              const hValue =
                i === rowInfos.length - 1 ? totalSum - even * (rowInfos.length - 1) : even;
              tr.setNodeMarkup(pos, undefined, {
                ...row.attrs,
                height: {
                  value: hValue,
                  rule: "atLeast",
                },
              });
            });
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Cell width — the width of every picked column (the grid is the one
      // width source the layout reads; Word's tcW maps onto it here).
      "cell-width":
        (value) =>
        ({ state, dispatch }) => {
          const tw = parseMeasureTwip(value);
          if (tw == null || tw < MIN_COL_TWIP) return false;
          const targets = tableTargets(state);
          if (!targets) return false;
          const widths = targets.tableNode.attrs.columnWidths as number[] | null;
          if (!widths || [...targets.cols].some((col) => col >= widths.length)) return false;
          if (dispatch) {
            const next = [...widths];
            for (const col of targets.cols) next[col] = Math.round(tw);
            dispatch(
              state.tr
                .setNodeMarkup(targets.tablePos, undefined, {
                  ...targets.tableNode.attrs,
                  columnWidths: next,
                })
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Word's Table Alignment: left (null in OOXML), center, right
      "table-alignment":
        (alignment) =>
        ({ state, dispatch }) => {
          if (alignment !== "left" && alignment !== "center" && alignment !== "right") {
            return false;
          }
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          const { $from } = state.selection;
          const tableNode = $from.node(anchor.tableAt);
          if (dispatch) {
            dispatch(
              state.tr
                .setNodeMarkup($from.before(anchor.tableAt), undefined, {
                  ...tableNode.attrs,
                  alignment: alignment === "left" ? null : alignment,
                })
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Word's Table Text Wrapping: none (inline / float null) vs around (floating / float object)
      "table-text-wrapping":
        (wrapping) =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          const { $from } = state.selection;
          const tableNode = $from.node(anchor.tableAt);

          let nextFloat: Record<string, unknown> | null = null;
          if (wrapping === "around" || wrapping === true) {
            nextFloat = (tableNode.attrs.float as Record<string, unknown> | null) ?? {
              horizontalAnchor: "margin",
              verticalAnchor: "paragraph",
            };
          } else if (wrapping === "none" || wrapping === false || wrapping === null) {
            nextFloat = null;
          } else if (wrapping !== undefined && typeof wrapping === "object") {
            nextFloat = wrapping as Record<string, unknown>;
          } else {
            return false;
          }

          if (dispatch) {
            dispatch(
              state.tr
                .setNodeMarkup($from.before(anchor.tableAt), undefined, {
                  ...tableNode.attrs,
                  float: nextFloat,
                })
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Table Properties dialog's OK / programmatic table properties patch —
      // writes table-level (alignment, indent, text wrapping/float) and row-level
      // (cantSplit, tableHeader) in one clean transaction.
      "table-properties-apply":
        (patch) =>
        ({ state, dispatch }) => {
          if (!patch || typeof patch !== "object") return false;
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          const { $from } = state.selection;
          const tableNode = $from.node(anchor.tableAt);

          const nextAttrs: Record<string, unknown> = { ...tableNode.attrs };
          let changedTable = false;

          if (patch.alignment !== undefined) {
            if (
              patch.alignment === "left" ||
              patch.alignment === "center" ||
              patch.alignment === "right"
            ) {
              nextAttrs.alignment = patch.alignment === "left" ? null : patch.alignment;
              changedTable = true;
            } else {
              return false;
            }
          }

          if (patch.indent !== undefined) {
            if (typeof patch.indent === "number" && patch.indent >= 0) {
              nextAttrs.indent = patch.indent > 0 ? Math.round(patch.indent) : null;
              changedTable = true;
            } else {
              return false;
            }
          }

          if (patch.textWrapping !== undefined) {
            if (patch.textWrapping === "around") {
              nextAttrs.float = (tableNode.attrs.float as Record<string, unknown> | null) ?? {
                horizontalAnchor: "margin",
                verticalAnchor: "paragraph",
              };
              changedTable = true;
            } else if (patch.textWrapping === "none") {
              nextAttrs.float = null;
              changedTable = true;
            } else {
              return false;
            }
          }

          if (patch.float !== undefined) {
            nextAttrs.float = patch.float ?? null;
            changedTable = true;
          }

          if (patch.cellSpacing !== undefined) {
            if (
              patch.cellSpacing === null ||
              patch.cellSpacing === 0 ||
              patch.cellSpacing === "0"
            ) {
              nextAttrs.cellSpacing = null;
              changedTable = true;
            } else if (typeof patch.cellSpacing === "number") {
              nextAttrs.cellSpacing =
                patch.cellSpacing > 0
                  ? { size: Math.round(patch.cellSpacing), type: "twips" }
                  : null;
              changedTable = true;
            } else if (typeof patch.cellSpacing === "string") {
              const parsed = parseMeasureTwip(patch.cellSpacing);
              nextAttrs.cellSpacing =
                parsed !== null && parsed > 0 ? { size: Math.round(parsed), type: "twips" } : null;
              changedTable = true;
            } else if (typeof patch.cellSpacing === "object") {
              nextAttrs.cellSpacing = patch.cellSpacing;
              changedTable = true;
            }
          }

          if (dispatch) {
            const tr = state.tr;
            if (changedTable) {
              tr.setNodeMarkup($from.before(anchor.tableAt), undefined, nextAttrs);
            }
            const targets = tableTargets(state);
            if (targets) {
              if (patch.cantSplit !== undefined) {
                stampRows(tr, targets, (row) => ({
                  ...row.attrs,
                  cantSplit: patch.cantSplit ? true : null,
                }));
              }
              if (patch.tableHeader !== undefined) {
                stampRows(tr, targets, (row) => ({
                  ...row.attrs,
                  tableHeader: patch.tableHeader ? true : null,
                }));
              }
            }
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
      // Table cell spacing (w:tblCellSpacing) — sets cell spacing twips on the enclosing table node.
      "table-cell-spacing":
        (value) =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          let spacing: { size: number; type?: string } | null = null;
          if (value === null || value === undefined || value === 0 || value === "0") {
            spacing = null;
          } else if (typeof value === "number") {
            spacing = value > 0 ? { size: Math.round(value), type: "twips" } : null;
          } else if (typeof value === "string") {
            const parsed = parseMeasureTwip(value);
            spacing =
              parsed !== null && parsed > 0 ? { size: Math.round(parsed), type: "twips" } : null;
          } else if (typeof value === "object" && "size" in (value as object)) {
            spacing = value as { size: number; type?: string };
          } else {
            return false;
          }
          if (dispatch) {
            const { $from } = state.selection;
            const tablePos = $from.before(anchor.tableAt);
            const tableNode = $from.node(anchor.tableAt);
            dispatch(
              state.tr
                .setNodeMarkup(tablePos, undefined, { ...tableNode.attrs, cellSpacing: spacing })
                .scrollIntoView(),
            );
          }
          return true;
        },
      // Row height — every picked row's trHeight (atLeast; "0"/auto clears
      // them all), so a whole-pick height lines the rows up (Word).
      "cell-height":
        (value) =>
        ({ state, dispatch }) => {
          const tw = parseMeasureTwip(value);
          if (tw == null || tw < 0) return false;
          const targets = tableTargets(state);
          if (!targets) return false;
          if (dispatch) {
            const height = tw > 0 ? { value: Math.round(tw), rule: "atLeast" } : null;
            const tr = state.tr;
            stampRows(tr, targets, (row) => ({ ...row.attrs, height }));
            dispatch(tr.scrollIntoView());
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
      // ── Styles system: redefine an existing style (the Modify Style dialog)
      // and the Design tab's style sets. Both stamp doc.attrs.styles in one
      // DocAttrStep, so the change rides the undo history and every consumer
      // (layout projection, compile) re-reads the same model.
      "modify-style":
        (patch) =>
        ({ tr }) => {
          if (!patch?.id) return false;
          const styles = { ...((tr.doc.attrs.styles ?? {}) as Record<string, unknown>) };
          const list = ((styles.paragraphStyles ?? []) as Record<string, unknown>[]).slice();
          const at = list.findIndex((s) => s.id === patch.id);
          // A built-in style may ALSO live under default.<key> (style sets
          // write there) — the pStyle id is the key with its first letter
          // upper-cased (Heading1 → heading1). The style index lets an
          // explicit paragraphStyles entry shadow the built-in slot, so both
          // sides must agree: once the style has an explicit definition the
          // built-in slot is dropped (Word does the same — a modified built-in
          // style becomes an explicit w:style). Copy slots before writing
          // (attrs may be aliased by a snapshot).
          const key = patch.id.charAt(0).toLowerCase() + patch.id.slice(1);
          const defaults = { ...((styles.default ?? {}) as Record<string, unknown>) };
          if (at >= 0) {
            list[at] = withModifyStylePatch(list[at], patch);
            styles.paragraphStyles = list;
            delete defaults[key];
            styles.default = defaults;
          } else {
            const entry = { ...((defaults[key] ?? {}) as Record<string, unknown>) };
            if (entry.name === undefined) entry.name = patch.id;
            defaults[key] = withModifyStylePatch(entry, patch);
            styles.default = defaults;
          }
          tr.step(new DocAttrStep("styles", styles));
          return true;
        },
      "style-set":
        (value) =>
        ({ tr }) => {
          const preset = STYLE_SET_PRESETS[value ?? ""];
          if (!preset) return false;
          const styles = { ...((tr.doc.attrs.styles ?? {}) as Record<string, unknown>) };
          const defaults = { ...((styles.default ?? {}) as Record<string, unknown>) };
          for (const [key, runPatch] of Object.entries(preset)) {
            const entry = { ...((defaults[key] ?? {}) as Record<string, unknown>) };
            entry.run = {
              ...((entry.run ?? {}) as Record<string, unknown>),
              ...(runPatch as Record<string, unknown>),
            };
            defaults[key] = entry;
          }
          styles.default = defaults;
          tr.step(new DocAttrStep("styles", styles));
          return true;
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
      // order (locale-aware, numeric). When selection is inside a table,
      // delegates to Word table sorting (sort-table).
      sort:
        () =>
        ({ state, chain, commands }) => {
          if (tableAncestry(state)) {
            return commands["sort-table"]();
          }
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
      // "preset:<id>" applies a List Library style (Word's gallery): plain
      // and bullet paragraphs gain one shared fresh reference of the preset,
      // numbered paragraphs keep their level and re-style wholesale. Plain
      // paragraphs with no value gain a fresh decimal multilevel list (a
      // gallery applies a list; a silent no-op reads as a broken button).
      "multilevel-list":
        (level) =>
        ({ state, tr }) => {
          const preset = level?.startsWith("preset:") ? level.slice(7) : null;
          const demote = level === "in" ? 1 : level === "out" ? -1 : 0;
          const target = level === "level-3" ? 2 : level === "level-2" ? 1 : 0;
          let touched = false;
          let freshRef: string | null = null;
          const freshFor = (base: string): string =>
            base === "ordered"
              ? nextOrderedReference(
                  collectListReferences(state.doc),
                  (state.doc.attrs as { numbering?: unknown }).numbering,
                )
              : nextMultilevelReference(
                  collectListReferences(state.doc),
                  (state.doc.attrs as { numbering?: unknown }).numbering,
                  preset!,
                );
          for (const { pos, node } of selectedParagraphs(state)) {
            const attrs = node.attrs as Record<string, unknown>;
            const cur = listStateOf(attrs);
            if (preset) {
              // The preset applies to every selected paragraph: numbered
              // ones re-style in place (level kept), bullets and plain
              // paragraphs convert into one shared list of the preset.
              const reference =
                cur.kind === "ordered" ? cur.reference! : (freshRef ??= freshFor("multilevel"));
              tr.setNodeMarkup(pos, undefined, {
                ...attrs,
                bullet: null,
                numbering: {
                  reference,
                  level: cur.kind === "ordered" ? cur.level : 0,
                },
              });
              touched = true;
              continue;
            }
            if (!cur.kind) {
              // One shared list for the whole selection (Word numbers the
              // applied gallery as one list).
              freshRef ??= freshFor("ordered");
              tr.setNodeMarkup(pos, undefined, {
                ...attrs,
                bullet: null,
                numbering: { reference: freshRef, level: target },
              });
              touched = true;
              continue;
            }
            const depth = demote === 0 ? target : Math.min(8, Math.max(0, cur.level + demote));
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
      // Move a selected floating drawing (image or wps shape) by a pointer-
      // drag delta: value is JSON {h, v} in EMU, added to the drawing's
      // current offsets. Align-anchored floats decline — the bridge commits
      // those through place-drawing instead.
      "move-drawing":
        (value?) =>
        ({ state, tr }) => {
          if (!value) return false;
          const target = floatingDrawingAt(state);
          if (!target) return false;
          let parsed: { h?: number; v?: number };
          try {
            parsed = JSON.parse(value) as { h?: number; v?: number };
          } catch {
            return false;
          }
          const floating = floatingOf(target);
          const h = floating.horizontalPosition as Record<string, unknown> | undefined;
          const v = floating.verticalPosition as Record<string, unknown> | undefined;
          if (!h || !v || typeof h.offset !== "number" || typeof v.offset !== "number")
            return false;
          return stampFloating(tr, target, {
            ...floating,
            horizontalPosition: { ...h, offset: h.offset + (parsed.h ?? 0) },
            verticalPosition: { ...v, offset: v.offset + (parsed.v ?? 0) },
          });
        },
      // Drop a dragged drawing at an absolute page position: value is JSON
      // {h, v} in EMU, page-local. An align-anchored float has no offset for
      // move-drawing to add to, so the drag lands as page-anchored offsets —
      // the painted position IS the value (Word converts an alignment to an
      // offset on drag; the drawn spot doesn't shift).
      "place-drawing":
        (value?) =>
        ({ state, tr }) => {
          if (!value) return false;
          const target = floatingDrawingAt(state);
          if (!target) return false;
          let parsed: { h?: number; v?: number };
          try {
            parsed = JSON.parse(value) as { h?: number; v?: number };
          } catch {
            return false;
          }
          if (typeof parsed.h !== "number" || typeof parsed.v !== "number") return false;
          return stampFloating(tr, target, {
            ...floatingOf(target),
            horizontalPosition: { relative: "page", offset: parsed.h },
            verticalPosition: { relative: "page", offset: parsed.v },
          });
        },
      // Rotate the selected drawing by a handle-swept delta: value is the
      // degrees to add to the drawing's current rotation (clockwise; image:
      // the flat rotation attr, floating or inline alike — the a:xfrm rot
      // spins the extent box either way; shape: its payload's transformation,
      // floating only).
      "rotate-drawing":
        (value?) =>
        ({ state, tr }) => {
          const delta = value == null ? Number.NaN : Number(value);
          if (!Number.isFinite(delta) || delta === 0) return false;
          const sel = state.selection;
          if (!(sel instanceof NodeSelection)) return false;
          if (sel.node.type.name === "image") {
            const attrs = sel.node.attrs as Record<string, unknown>;
            const target = { pos: sel.from, attrs, kind: "image" as const };
            const current = attrs.rotation;
            return stampAttrs(tr, target, {
              ...attrs,
              rotation: (typeof current === "number" ? current : 0) + delta,
            });
          }
          const target = floatingDrawingAt(state);
          if (!target) return false;
          const shape = target.attrs.wpsShape as Record<string, unknown>;
          const transformation = {
            ...(shape.transformation as Record<string, unknown> | undefined),
          };
          const current = transformation.rotation;
          transformation.rotation = (typeof current === "number" ? current : 0) + delta;
          return stampAttrs(tr, target, {
            ...target.attrs,
            wpsShape: { ...shape, transformation },
          });
        },
      // The Size-and-Position dialog's OK: absolute geometry in centimeters
      // (the dialog's display unit), converted to each carrier's native unit —
      // an image sizes in px and offsets in EMU, a shape payload in EMU.
      "drawing-properties-apply":
        (patch?) =>
        ({ state, tr }) => {
          if (!patch || typeof patch !== "object") return false;
          const target = floatingDrawingAt(state);
          if (!target) return false;
          const cmTo = (v: number, factor: number): number => Math.round(v * factor);
          const PX_PER_CM = 96 / 2.54;
          const EMU_PER_CM = 360000;
          const num = (v: unknown): number | null =>
            typeof v === "number" && Number.isFinite(v) ? v : null;
          const widthCm = num(patch.widthCm);
          const heightCm = num(patch.heightCm);
          const rotationDeg = num(patch.rotationDeg);
          const offsetHCm = num(patch.offsetHCm);
          const offsetVCm = num(patch.offsetVCm);
          if (target.kind === "image") {
            const attrs = { ...target.attrs };
            if (widthCm != null) attrs.width = cmTo(widthCm, PX_PER_CM);
            if (heightCm != null) attrs.height = cmTo(heightCm, PX_PER_CM);
            if (rotationDeg != null) attrs.rotation = rotationDeg;
            const floating = { ...floatingOf(target) };
            const hPos = {
              ...(floating.horizontalPosition as Record<string, unknown> | undefined),
            };
            const vPos = {
              ...(floating.verticalPosition as Record<string, unknown> | undefined),
            };
            if (offsetHCm != null) hPos.offset = cmTo(offsetHCm, EMU_PER_CM);
            if (offsetVCm != null) vPos.offset = cmTo(offsetVCm, EMU_PER_CM);
            floating.horizontalPosition = hPos;
            floating.verticalPosition = vPos;
            return stampAttrs(tr, target, { ...attrs, floating });
          }
          const shape = { ...(target.attrs.wpsShape as Record<string, unknown>) };
          const t = { ...((shape.transformation ?? {}) as Record<string, unknown>) };
          if (widthCm != null) t.width = cmTo(widthCm, EMU_PER_CM);
          if (heightCm != null) t.height = cmTo(heightCm, EMU_PER_CM);
          if (rotationDeg != null) t.rotation = rotationDeg;
          shape.transformation = t;
          // The shape's offsets ride the same Floating object as an image's.
          if (offsetHCm == null && offsetVCm == null)
            return stampAttrs(tr, target, { ...target.attrs, wpsShape: shape });
          const floating = { ...floatingOf(target) };
          const hPos = { ...(floating.horizontalPosition as Record<string, unknown> | undefined) };
          const vPos = { ...(floating.verticalPosition as Record<string, unknown> | undefined) };
          if (offsetHCm != null) hPos.offset = cmTo(offsetHCm, EMU_PER_CM);
          if (offsetVCm != null) vPos.offset = cmTo(offsetVCm, EMU_PER_CM);
          floating.horizontalPosition = hPos;
          floating.verticalPosition = vPos;
          shape.floating = floating;
          return stampAttrs(tr, target, { ...target.attrs, wpsShape: shape });
        },
      // The crop overlay's commit: the selected image's new a:srcRect insets
      // as source fractions, stored as the raw ST_Percentage ints the attrs
      // carry (office-open's parse emits raw ints despite the documented
      // integer percent — mirror cropOf's /100000 read side). All zero clears.
      "drawing-crop-apply":
        (patch?) =>
        ({ state, tr }) => {
          if (!patch || typeof patch !== "object") return false;
          const sel = state.selection;
          if (!(sel instanceof NodeSelection) || sel.node.type.name !== "image") return false;
          const num = (v: unknown): number | null =>
            typeof v === "number" && Number.isFinite(v) ? v : null;
          const left = num(patch.left);
          const top = num(patch.top);
          const right = num(patch.right);
          const bottom = num(patch.bottom);
          if (left == null || top == null || right == null || bottom == null) return false;
          const raw = (fraction: number): number => Math.round(fraction * 100000);
          const crop = { left: raw(left), top: raw(top), right: raw(right), bottom: raw(bottom) };
          const attrs = { ...sel.node.attrs };
          if (crop.left === 0 && crop.top === 0 && crop.right === 0 && crop.bottom === 0)
            delete attrs.crop;
          else attrs.crop = crop;
          tr.setNodeMarkup(sel.from, undefined, attrs);
          tr.setSelection(NodeSelection.create(tr.doc, sel.from) as never);
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
          // No value = the split's primary face — Word defaults it to left
          // (the button's icon); an unknown value declines.
          const align =
            value === "center" || value === "right"
              ? value
              : value === "left" || value == null || value === ""
                ? "left"
                : null;
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
