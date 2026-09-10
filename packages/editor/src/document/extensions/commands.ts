import type { ChartOptions, ChartType, ImageAttrs, LegendPosition } from "@docen/docx";
import {
  BULLET_GLYPHS,
  HIGHLIGHT_PALETTE_RGB,
  nextMultilevelReference,
  nextOrderedReference,
  ORDERED_FORMATS,
} from "@docen/docx";
import { Extension } from "@docen/docx/core";
import { EMU_PER_PX } from "@docen/layout";
import type { Node as PMNode, ResolvedPos } from "@tiptap/pm/model";
import type { Mark } from "@tiptap/pm/model";
import type { EditorState } from "@tiptap/pm/state";
import type { Transaction } from "@tiptap/pm/state";
import { NodeSelection, TextSelection } from "@tiptap/pm/state";
import { DocAttrStep } from "@tiptap/pm/transform";

import { freshChildEmu, memberEmuOf, unionBox, type Box } from "../../drawing";
import { CellSelection, cellsInRect } from "../canvas/cell-selection";

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
      chart: () => ReturnType;
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
      "distribute-rows": () => ReturnType;
      "cell-margins": (value?: string) => ReturnType;
      "cell-width": (value?: string) => ReturnType;
      "cell-height": (value?: string) => ReturnType;
      "table-properties-apply": (patch?: TablePropertiesPatch) => ReturnType;
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
      "reanchor-drawing": (value?: string) => ReturnType;
      "rotate-drawing": (value?: string) => ReturnType;
      "drawing-properties-apply": (patch?: DrawingPropertiesPatch) => ReturnType;
      "drawing-crop-apply": (patch?: DrawingCropPatch) => ReturnType;
      "drawing-crop-reset": () => ReturnType;
      // Word's Crop → Aspect Ratio presets — a "W:H" ratio string.
      "drawing-crop-aspect": (value?: string) => ReturnType;
      // Picture Format tab's Size boxes — a measure string ("5cm"/"2in"/…).
      "drawing-width": (value?: string) => ReturnType;
      "drawing-height": (value?: string) => ReturnType;
      // Picture Format > Adjust — the pixel-adjustment presets. The values
      // are the ribbon menu's value strings ("bright:40", "saturation:66",
      // "50", "color:FF0000" / "width:2.25" / "dash:sysDot" / "none").
      "picture-correction": (value?: string) => ReturnType;
      "picture-color": (value?: string) => ReturnType;
      "picture-transparency": (value?: string) => ReturnType;
      "picture-border": (value?: string) => ReturnType;
      "reset-picture": () => ReturnType;
      "reset-picture-size": (natural?: { width: number; height: number }) => ReturnType;
      "change-picture": (src?: string) => ReturnType;
      // The pixel tools' commit (Compress Pictures / Set Transparent Color):
      // a JSON `{ src, dropCrop? }` — the pixels are re-encoded at the UI
      // layer, this only swaps the source (and drops the srcRect when the
      // crop was baked in).
      "picture-pixels": (value?: string) => ReturnType;
      "shape-fill": (value?: string) => ReturnType;
      "shape-outline": (value?: string) => ReturnType;
      "shape-effects": (value?: string) => ReturnType;
      "shape-text-direction": (value?: string) => ReturnType;
      // Chart Design tab — the type token (column/bar/line/area/pie/doughnut/
      // scatter), the legend placement ("none" or a LegendPosition), and the
      // Edit Data dialog's commit (JSON {title?, categories?, series?}).
      "chart-type": (value?: string) => ReturnType;
      "chart-legend": (value?: string) => ReturnType;
      "chart-data-apply": (value?: string) => ReturnType;
      // value is the series index to remove (the plot's sub-selected series).
      "chart-series-delete": (value?: string) => ReturnType;
      // The plot's value-drag commit — JSON {series, point, value} writing one
      // data point (Excel's drag-a-point editing).
      "chart-value-apply": (value?: string) => ReturnType;
      // Arrange — floating drawings (z-order, wrap, rotation, position).
      "bring-forward": () => ReturnType;
      "send-backward": () => ReturnType;
      "bring-to-front": () => ReturnType;
      "send-to-back": () => ReturnType;
      wrap: (value?: string) => ReturnType;
      rotate: (value?: string) => ReturnType;
      position: (value?: string) => ReturnType;
      "align-objects": (value?: string) => ReturnType;
      // Group/Ungroup/Distribute — the Arrange group's multi-selection
      // actions. payload is JSON {members: [{pos, box}]} with the member page
      // boxes in px (the multi-selection overlay's hit boxes); distribute's
      // value is "h" | "v".
      "drawing-group": (payload?: string) => ReturnType;
      "drawing-ungroup": () => ReturnType;
      "drawing-distribute": (value?: string, payload?: string) => ReturnType;
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
  "chart",
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
  "distribute-rows",
  "cell-margins",
  "cell-width",
  "cell-height",
  "text-direction",
  "convert-to-text",
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
  "reanchor-drawing",
  "rotate-drawing",
  "drawing-properties-apply",
  "drawing-crop-apply",
  "drawing-crop-reset",
  "drawing-crop-aspect",
  "drawing-width",
  "drawing-height",
  "picture-correction",
  "picture-color",
  "picture-transparency",
  "picture-border",
  "reset-picture",
  "reset-picture-size",
  "change-picture",
  "picture-pixels",
  "shape-fill",
  "shape-outline",
  "shape-effects",
  "shape-text-direction",
  "chart-type",
  "chart-legend",
  "chart-data-apply",
  "chart-series-delete",
  "chart-value-apply",
  "bring-forward",
  "send-backward",
  "bring-to-front",
  "send-to-back",
  "wrap",
  "rotate",
  "position",
  "align-objects",
  "drawing-group",
  "drawing-ungroup",
  "drawing-distribute",
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
 * What the Table Properties dialog commits on OK — the table tab's geometry
 * (alignment and left indent in twips), stamped onto the caret's table by
 * {@link documentCommands.table-properties-apply}.
 */
export interface TablePropertiesPatch {
  /** w:jc table alignment — "left" commits null (OOXML's default). */
  alignment: "left" | "center" | "right";
  /** w:tblInd left indent in twips; 0 commits null. */
  indent: number;
}

/**
 * What the Size-and-Position dialog commits on OK — the selected floating
 * drawing's geometry in centimeters (the dialog's display unit), stamped by
 * {@link documentCommands.drawing-properties-apply}. Absent fields keep the
 * current value.
 */
export interface DrawingPropertiesPatch {
  /** The drawing box width in cm (px for an image, EMU for a shape payload). */
  widthCm?: number;
  /** The drawing box height in cm. */
  heightCm?: number;
  /** Clockwise rotation about the box center, degrees. */
  rotationDeg?: number;
  /** Horizontal offset from the anchor's horizontal base, cm → EMU. */
  offsetHCm?: number;
  /** Vertical offset from the anchor's vertical base, cm → EMU. */
  offsetVCm?: number;
  /** The replacement text (Word's Alt Text); empty clears it. Images only —
   *  a shape payload carries no alt-text field today. */
  altText?: string;
  /** The horizontal position base (an ST_RelFromH token: column/margin/page). */
  relativeH?: string;
  /** The vertical position base (an ST_RelFromV token: paragraph/line/margin/page). */
  relativeV?: string;
  /** Whether other floating drawings may overlap this one. */
  allowOverlap?: boolean;
  /** Whether the drawing keeps its table-cell layout behavior. */
  layoutInCell?: boolean;
  /** Whether the anchor moves with its paragraph (Word's Lock anchor). */
  lockAnchor?: boolean;
  /** The wrap distances (Word's Distance from text), cm → EMU. */
  distanceCm?: { top: number; bottom: number; left: number; right: number };
}

/**
 * What the Edit Data dialog commits on OK — the grid's title/category/series
 * values, stamped onto the selected chart by the chart-data-apply command.
 * The chart payload's other fields (type, legend, anchor) ride through
 * untouched.
 */
export interface ChartDataPatch {
  /** The chart title text; empty clears it. */
  title?: string;
  /** The category axis labels, one per plotted row. */
  categories: string[];
  /** One entry per series — the legend name and its values. */
  series: { name: string; values: number[] }[];
}

/**
 * What the crop mode commits — the selected image's a:srcRect insets as
 * fractions of the source (0.1 = 10% off that edge), stamped by {@link
 * documentCommands.drawing-crop-apply}. The extent resizes to the kept
 * region at the unchanged source scale. An all-zero set clears the crop
 * (and grows the extent back to the full source).
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

/** The selection's table targets — every cell a CellSelection crosses, or the
 *  caret's enclosing cell. A CellSelection's `$from` sits at a cell's START
 *  (inside the row, not the cell), so {@link tableAncestry} reads `cellAt:
 *  -1` there and every cell-level command would decline — this is the one
 *  resolver the cell/row/column-level commands share. Cell stamps carry their
 *  table-child row index and grid column; the row/column commands read
 *  `rows`/`cols` (Word: a whole-pick height lands on every picked row, a
 *  width on every picked column). Null outside a table. */
function tableTargets(state: EditorState): {
  tablePos: number;
  tableNode: PMNode;
  cells: { pos: number; node: PMNode }[];
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
  const cells: { pos: number; node: PMNode }[] = [];
  const rows = new Set<number>();
  const cols = new Set<number>();
  cellsInRect(
    state.doc,
    anchorCell,
    isCell ? selection.headCell : anchorCell,
    (node, pos, row, col) => {
      cells.push({ pos, node });
      rows.add(row);
      cols.add(col);
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
 *  `floating` attr set), a wps shape (floating inside its `wpsShape` payload),
 *  a wpg group (inside `wpgGroup`), or a chart (inside `chart`); the stage's
 *  hit boxes produce exactly these. The ribbon's dynamic pass greys the
 *  Arrange controls against this (and the inline twins below). */
export type FloatingDrawing = {
  pos: number;
  attrs: Record<string, unknown>;
  kind: "image" | "shape" | "group" | "chart";
};

export function floatingDrawingAt(state: EditorState): FloatingDrawing | null {
  const sel = state.selection;
  if (!(sel instanceof NodeSelection)) return null;
  return drawingAtPos(state.doc, sel.from);
}

/** The floating drawing at a document position — floatingDrawingAt's by-pos
 *  form, for the multi-selection payload (no NodeSelection involved). */
function drawingAtPos(doc: PMNode, pos: number): FloatingDrawing | null {
  const node = doc.nodeAt(pos);
  if (!node) return null;
  const attrs = node.attrs as Record<string, unknown>;
  if (node.type.name === "image") {
    return attrs.floating ? { pos, attrs, kind: "image" } : null;
  }
  if (node.type.name === "wpsShape") {
    const shape = attrs.wpsShape as Record<string, unknown> | null;
    return shape?.floating ? { pos, attrs, kind: "shape" } : null;
  }
  if (node.type.name === "wpgGroup") {
    const group = attrs.wpgGroup as Record<string, unknown> | null;
    return group?.floating ? { pos, attrs, kind: "group" } : null;
  }
  if (node.type.name === "chart") {
    const chart = attrs.chart as Record<string, unknown> | null;
    return chart?.floating ? { pos, attrs, kind: "chart" } : null;
  }
  return null;
}

/** Where each non-image kind carries its Floating object (the chart's rides
 *  its own payload). */
const FLOATING_CARRIER = { shape: "wpsShape", group: "wpgGroup", chart: "chart" } as const;

/** The drawing's Floating object (image: a flat attr; shape/group/chart:
 *  inside their payload). */
function floatingOf(target: FloatingDrawing): Record<string, unknown> {
  const carrier =
    target.kind === "image"
      ? target.attrs.floating
      : (target.attrs[FLOATING_CARRIER[target.kind]] as Record<string, unknown>).floating;
  return carrier as Record<string, unknown>;
}

/** The inline drawing under a NodeSelection — any of the four kinds whose
 *  payload lacks a Floating object. Word's Wrap Text and Position galleries
 *  also serve these: taking one converts the drawing to a floating one
 *  anchored to its own paragraph (keep-position). */
export function inlineDrawingAt(state: EditorState): FloatingDrawing | null {
  const sel = state.selection;
  if (!(sel instanceof NodeSelection)) return null;
  const attrs = sel.node.attrs as Record<string, unknown>;
  const pos = sel.from;
  switch (sel.node.type.name) {
    case "image":
      return attrs.floating ? null : { pos, attrs, kind: "image" };
    case "wpsShape": {
      const shape = attrs.wpsShape as Record<string, unknown> | null;
      return shape && !shape.floating ? { pos, attrs, kind: "shape" } : null;
    }
    case "wpgGroup": {
      const group = attrs.wpgGroup as Record<string, unknown> | null;
      return group && !group.floating ? { pos, attrs, kind: "group" } : null;
    }
    case "chart": {
      const chart = attrs.chart as Record<string, unknown> | null;
      return chart && !chart.floating ? { pos, attrs, kind: "chart" } : null;
    }
    default:
      return null;
  }
}

/** The inline picture under a NodeSelection — Word's Rotate menu works on
 *  embedded images too (shapes and charts don't rotate inline). */
export function inlineImageAt(state: EditorState): FloatingDrawing | null {
  const inline = inlineDrawingAt(state);
  return inline?.kind === "image" ? inline : null;
}

/** The selected wps shape — standalone or a group member. Style commands
 *  (fill/outline) write its attrs wherever it sits, unlike the Arrange
 *  commands which need a floating carrier. */
function shapeAt(
  state: EditorState,
): { pos: number; attrs: Record<string, unknown>; kind: "shape" } | null {
  const sel = state.selection;
  if (!(sel instanceof NodeSelection) || sel.node.type.name !== "wpsShape") return null;
  const attrs = sel.node.attrs as Record<string, unknown>;
  return attrs.wpsShape ? { pos: sel.from, attrs, kind: "shape" } : null;
}

/** The chart types whose series carry a grouping (w:bar/line/area families +
 *  stock's OHLC lanes) — the rest have none to hand down. */
const GROUPING_CHART_TYPES: readonly ChartType[] = ["column", "bar", "line", "area", "stock"];

/** The selected chart node — inline or floating alike: the type/legend/data
 *  edits write the chart payload wherever the chart sits. attrs.chart is a
 *  ChartOptions verbatim (the chart node's contract), typed here so every
 *  command below edits real fields instead of Record lookups. */
function chartAt(state: EditorState): {
  pos: number;
  attrs: Record<string, unknown>;
  chart: ChartOptions;
} | null {
  const sel = state.selection;
  if (!(sel instanceof NodeSelection) || sel.node.type.name !== "chart") return null;
  const attrs = sel.node.attrs as Record<string, unknown>;
  return attrs.chart ? { pos: sel.from, attrs, chart: attrs.chart as ChartOptions } : null;
}

/** Stamp the patched chart payload back onto the chart node, restoring the
 *  NodeSelection the way the Arrange commands do (the markup write collapses
 *  it — the tab must stay so the edit can repeat). */
function stampChart(
  tr: Transaction,
  target: NonNullable<ReturnType<typeof chartAt>>,
  chart: ChartOptions,
): boolean {
  tr.setNodeMarkup(target.pos, undefined, { ...target.attrs, chart });
  tr.setSelection(NodeSelection.create(tr.doc, target.pos));
  return true;
}

/** Write a Floating back onto the drawing, shallow-copying the carrier the
 *  way PM immutability requires (image: flat; shape/group/chart: the payload). */
function withFloating(
  target: NonNullable<ReturnType<typeof floatingDrawingAt>>,
  floating: Record<string, unknown>,
): Record<string, unknown> {
  if (target.kind === "image") return { ...target.attrs, floating };
  const key = FLOATING_CARRIER[target.kind];
  return {
    ...target.attrs,
    [key]: { ...(target.attrs[key] as Record<string, unknown>), floating },
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

/** A floating whose offsets moved by (dx, dy) EMU — group/ungroup/distribute's
 *  shared write. The caller guarantees both offsets are numbers. */
function offsetFloating(
  floating: Record<string, unknown>,
  dx: number,
  dy: number,
): Record<string, unknown> {
  const h = { ...(floating.horizontalPosition as Record<string, unknown>) };
  const v = { ...(floating.verticalPosition as Record<string, unknown>) };
  h.offset = (h.offset as number) + dx;
  v.offset = (v.offset as number) + dy;
  return { ...floating, horizontalPosition: h, verticalPosition: v };
}

// ── Group/Ungroup member mapping (Word's Group/Ungroup on a selection) ──────

/** One member's attrs re-homed into a fresh group's child space: the image's
 *  EMU box moves to groupXfrm (its floating dropped), the shape's
 *  transformation becomes the child-space offset form, a nested group's
 *  payload flips to the GroupMediaData child-space transformation. A chart
 *  has no standalone in-group carrier yet — null declines the grouping.
 *  Everything else (src/crop/fill/body) rides along untouched. */
function memberInGroupAttrs(
  kind: "image" | "shape" | "group" | "chart",
  attrs: Record<string, unknown>,
  child: { x: number; y: number; cx: number; cy: number },
): Record<string, unknown> | null {
  if (kind === "chart") return null;
  if (kind === "image") {
    const next: Record<string, unknown> = { ...attrs, groupXfrm: child };
    delete next.floating;
    return next;
  }
  if (kind === "shape") {
    const shape = { ...(attrs.wpsShape as Record<string, unknown>) };
    delete shape.floating;
    shape.transformation = {
      ...(shape.transformation as Record<string, unknown> | undefined),
      offset: { left: child.x, top: child.y },
      width: child.cx,
      height: child.cy,
    };
    return { ...attrs, wpsShape: shape };
  }
  // A nested group: GroupOptions form (transformation {width,height} +
  // floating) → GroupMediaData child form (offset.emus + emus), its own
  // chOff/chExt preserved — the member's inner space is untouched.
  const px = (emu: number): number => Math.round(emu / EMU_PER_PX);
  const { transformation: _t, floating: _f, ...rest } = attrs.wpgGroup as Record<string, unknown>;
  return {
    ...attrs,
    wpgGroup: {
      ...rest,
      transformation: {
        offset: { emus: { x: child.x, y: child.y }, pixels: { x: px(child.x), y: px(child.y) } },
        emus: { x: child.cx, y: child.cy },
        pixels: { x: px(child.cx), y: px(child.cy) },
      },
    },
  };
}

/** A member node's child-space EMU box — the group-child geometry each member
 *  kind carries (image: groupXfrm; shape: transformation.offset; nested group:
 *  the MediaDataTransformation pair). Null for members whose payload cannot
 *  carry geometry (passthrough atoms). */
function childBoxOf(member: PMNode): { x: number; y: number; cx: number; cy: number } | null {
  const attrs = member.attrs as Record<string, unknown>;
  if (member.type.name === "image") {
    const xfrm = attrs.groupXfrm as { x: number; y: number; cx: number; cy: number } | undefined;
    return xfrm && Number.isFinite(xfrm.cx) && Number.isFinite(xfrm.cy) ? xfrm : null;
  }
  if (member.type.name === "wpsShape") {
    const t = (attrs.wpsShape as Record<string, unknown>)?.transformation as
      | { offset?: { left?: number; top?: number }; width?: unknown; height?: unknown }
      | undefined;
    const cx = parseMeasureEmu(t?.width);
    const cy = parseMeasureEmu(t?.height);
    if (!t || cx == null || cy == null) return null;
    return { x: t.offset?.left ?? 0, y: t.offset?.top ?? 0, cx, cy };
  }
  if (member.type.name === "wpgGroup") {
    const t = (attrs.wpgGroup as Record<string, unknown>)?.transformation as
      | { offset?: { emus?: { x?: number; y?: number } }; emus?: { x?: number; y?: number } }
      | undefined;
    if (typeof t?.emus?.x !== "number" || typeof t.emus.y !== "number") return null;
    return { x: t.offset?.emus?.x ?? 0, y: t.offset?.emus?.y ?? 0, cx: t.emus.x, cy: t.emus.y };
  }
  return null;
}

/** Ungroup's reverse of memberInGroupAttrs: one member node back to a
 *  standalone floating node, sized `back.cx × back.cy` EMU and anchored
 *  `back.(dx,dy)` EMU from the group's own floating (it lands where the group
 *  drew it — Word's Ungroup keeps positions). Null for a member with no
 *  standalone carrier (chart/content-part passthrough atoms) — the command
 *  declines rather than dropping members. */
function memberOutNode(
  member: PMNode,
  groupFloating: Record<string, unknown>,
  back: { dx: number; dy: number; cx: number; cy: number },
): PMNode | null {
  const attrs = member.attrs as Record<string, unknown>;
  const floating = offsetFloating(groupFloating, back.dx, back.dy);
  if (member.type.name === "image") {
    if (!attrs.groupXfrm) return null;
    const next: Record<string, unknown> = {
      ...attrs,
      width: Math.round(back.cx / EMU_PER_PX),
      height: Math.round(back.cy / EMU_PER_PX),
      floating,
    };
    delete next.groupXfrm;
    return member.type.create(next, member.content, member.marks);
  }
  if (member.type.name === "wpsShape") {
    const shape = { ...(attrs.wpsShape as Record<string, unknown>) };
    const t = { ...(shape.transformation as Record<string, unknown>) };
    delete t.offset;
    t.width = back.cx;
    t.height = back.cy;
    shape.transformation = t;
    shape.floating = floating;
    return member.type.create({ ...attrs, wpsShape: shape }, member.content, member.marks);
  }
  if (member.type.name === "wpgGroup") {
    // Back to the GroupOptions form: extent from the mapped size, floating
    // from the group's, its own chOff/chExt preserved.
    const { transformation: _t, ...rest } = attrs.wpgGroup as Record<string, unknown>;
    return member.type.create(
      {
        ...attrs,
        wpgGroup: { ...rest, transformation: { width: back.cx, height: back.cy }, floating },
      },
      member.content,
      member.marks,
    );
  }
  return null;
}

/** The multi-selection payload group and distribute share: JSON
 *  {members: [{pos, box}]} with page-px boxes (the multi-selection overlay's
 *  hit boxes), every member a floating drawing. Sorted by document position —
 *  group replaces at the first; distribute re-sorts per axis. */
function multiMembersOf(
  state: EditorState,
  payload: string | undefined,
): { target: NonNullable<ReturnType<typeof drawingAtPos>>; box: Box }[] | null {
  if (!payload) return null;
  let parsed: unknown;
  try {
    parsed = JSON.parse(payload);
  } catch {
    return null;
  }
  const members = (parsed as { members?: unknown } | null)?.members;
  if (!Array.isArray(members) || members.length < 2) return null;
  const out: { target: NonNullable<ReturnType<typeof drawingAtPos>>; box: Box }[] = [];
  for (const m of members) {
    const { pos, box } = (m ?? {}) as { pos?: unknown; box?: Partial<Box> };
    const dims = box && [box.x, box.y, box.width, box.height];
    if (
      typeof pos !== "number" ||
      !dims ||
      dims.some((n) => typeof n !== "number" || !Number.isFinite(n))
    )
      return null;
    const target = drawingAtPos(state.doc, pos);
    if (!target) return null;
    out.push({ target, box: box as Box });
  }
  out.sort((a, b) => a.target.pos - b.target.pos);
  return out;
}

/** The Size-and-Position dialog's Floating half — position bases, layout
 *  flags, wrap distances, and the anchor offsets — folded onto a fresh copy
 *  of the drawing's Floating (the image and shape branches share it). */
function applyFloatingExtras(
  floating: Record<string, unknown>,
  patch: DrawingPropertiesPatch,
  offsetHCm: number | null,
  offsetVCm: number | null,
): Record<string, unknown> {
  const EMU_PER_CM = 360000;
  const next = { ...floating };
  const hPos = { ...(floating.horizontalPosition as Record<string, unknown> | undefined) };
  const vPos = { ...(floating.verticalPosition as Record<string, unknown> | undefined) };
  if (offsetHCm != null) hPos.offset = Math.round(offsetHCm * EMU_PER_CM);
  if (offsetVCm != null) vPos.offset = Math.round(offsetVCm * EMU_PER_CM);
  if (typeof patch.relativeH === "string") hPos.relative = patch.relativeH;
  if (typeof patch.relativeV === "string") vPos.relative = patch.relativeV;
  next.horizontalPosition = hPos;
  next.verticalPosition = vPos;
  if (typeof patch.allowOverlap === "boolean") next.allowOverlap = patch.allowOverlap;
  if (typeof patch.layoutInCell === "boolean") next.layoutInCell = patch.layoutInCell;
  if (typeof patch.lockAnchor === "boolean") next.lockAnchor = patch.lockAnchor;
  if (patch.distanceCm && typeof patch.distanceCm === "object") {
    const emu = (v: number): number => Math.round(v * EMU_PER_CM);
    next.margins = {
      top: emu(patch.distanceCm.top),
      bottom: emu(patch.distanceCm.bottom),
      left: emu(patch.distanceCm.left),
      right: emu(patch.distanceCm.right),
    };
  }
  return next;
}

/** Trade the picture's extent by a crop patch's kept fractions (new = old ×
 *  fNew/fOld per axis — Word's crop keeps the source scale, so the extent
 *  tracks the kept region) and stamp/delete attrs.crop: the all-zero patch
 *  grows the frame back to the full source, which is both Reset Crop and
 *  Reset Picture's crop discard. A degenerate fraction (a fOld of 0 from a
 *  hostile file, a non-positive fNew from a raw command call) leaves that
 *  axis alone. */
function tradeCropExtent(
  attrs: Record<string, unknown>,
  left: number,
  top: number,
  right: number,
  bottom: number,
): void {
  const num = (v: unknown): number | null =>
    typeof v === "number" && Number.isFinite(v) ? v : null;
  const raw = (fraction: number): number => Math.round(fraction * 100000);
  const crop = { left: raw(left), top: raw(top), right: raw(right), bottom: raw(bottom) };
  const keptOf = (a: unknown, b: unknown): number =>
    1 - (num(a) ?? 0) / 100000 - (num(b) ?? 0) / 100000;
  const prev = (attrs.crop ?? {}) as Record<string, unknown>;
  const fOldW = keptOf(prev.left, prev.right);
  const fOldH = keptOf(prev.top, prev.bottom);
  const fNewW = 1 - left - right;
  const fNewH = 1 - top - bottom;
  if (fOldW > 0 && fNewW > 0 && typeof attrs.width === "number")
    attrs.width = Math.round((attrs.width * fNewW) / fOldW);
  if (fOldH > 0 && fNewH > 0 && typeof attrs.height === "number")
    attrs.height = Math.round((attrs.height * fNewH) / fOldH);
  if (crop.left === 0 && crop.top === 0 && crop.right === 0 && crop.bottom === 0) delete attrs.crop;
  else attrs.crop = crop;
}

/** Stamp the crop patch onto the selected image — the shared body of
 *  drawing-crop-apply (the crop overlay's commit) and drawing-crop-reset. */ function applyCropPatch(
  state: EditorState,
  tr: Transaction,
  patch: DrawingCropPatch,
): boolean {
  const sel = state.selection;
  if (!(sel instanceof NodeSelection) || sel.node.type.name !== "image") return false;
  const num = (v: unknown): number | null =>
    typeof v === "number" && Number.isFinite(v) ? v : null;
  const left = num(patch.left);
  const top = num(patch.top);
  const right = num(patch.right);
  const bottom = num(patch.bottom);
  if (left == null || top == null || right == null || bottom == null) return false;
  const attrs = { ...sel.node.attrs };
  tradeCropExtent(attrs, left, top, right, bottom);
  tr.setNodeMarkup(sel.from, undefined, attrs);
  tr.setSelection(NodeSelection.create(tr.doc, sel.from) as never);
  return true;
}

/** Word's Crop → Aspect Ratio preset — the shared body of
 *  drawing-crop-aspect. Fits the largest centered rect of the "W:H" ratio
 *  inside the picture's kept region: the kept fractions are source-space but
 *  the ratio applies to the content's pixel proportions, so the frame's
 *  display extent (which follows the kept region at the unchanged source
 *  scale) supplies the proportion bridge between the two spaces. */ function applyCropAspect(
  state: EditorState,
  tr: Transaction,
  value?: string,
): boolean {
  const sel = state.selection;
  if (!(sel instanceof NodeSelection) || sel.node.type.name !== "image") return false;
  const m = /^(\d+(?:\.\d+)?):(\d+(?:\.\d+)?)$/.exec((value ?? "").trim());
  if (!m) return false;
  const target = Number(m[1]) / Number(m[2]);
  const attrs = sel.node.attrs as Record<string, unknown>;
  if (typeof attrs.width !== "number" || typeof attrs.height !== "number") return false;
  const frame = attrs.width / attrs.height;
  if (!(frame > 0) || !Number.isFinite(frame) || !(target > 0)) return false;
  const frac = (v: unknown): number =>
    typeof v === "number" && Number.isFinite(v) ? v / 100000 : 0;
  const prev = (attrs.crop ?? {}) as Record<string, unknown>;
  const cl = frac(prev.left);
  const ct = frac(prev.top);
  const cr = frac(prev.right);
  const cb = frac(prev.bottom);
  const keptW = 1 - cl - cr;
  const keptH = 1 - ct - cb;
  if (keptW <= 0 || keptH <= 0) return false;
  // A wider-than-target frame trims the sides; a taller one trims top/bottom.
  const newW = frame > target ? (keptW * target) / frame : keptW;
  const newH = frame > target ? keptH : (keptH * frame) / target;
  return applyCropPatch(state, tr, {
    left: cl + (keptW - newW) / 2,
    right: cr + (keptW - newW) / 2,
    top: ct + (keptH - newH) / 2,
    bottom: cb + (keptH - newH) / 2,
  });
}

/** Stamp one dimension onto the selected drawing — the shared body of the
 *  Picture/Shape Format tabs' Height/Width boxes (a measure string converted
 *  at 96 DPI). An image (inline or floating) sizes in px attrs; a shape/group
 *  payload in EMU — a group's extent change scales its members through the
 *  child coordinate space (Word's group resize). */ function applyImageSize(
  state: EditorState,
  tr: Transaction,
  axis: "width" | "height",
  value: unknown,
): boolean {
  const tw = parseMeasureTwip(value);
  if (tw == null || tw <= 0) return false;
  const sel = state.selection;
  if (!(sel instanceof NodeSelection)) return false;
  // 96-DPI px is the finest granularity an extent stores — a measure under
  // half a px would round to a zero extent (an invisible drawing).
  const px = Math.round(tw / 15); // 15 twips to the px at 96 DPI
  if (px <= 0) return false;
  const emu = px * 9525; // 9525 to the EMU
  if (sel.node.type.name === "image") {
    const attrs = { ...sel.node.attrs };
    attrs[axis] = px;
    tr.setNodeMarkup(sel.from, undefined, attrs);
    tr.setSelection(NodeSelection.create(tr.doc, sel.from) as never);
    return true;
  }
  if (sel.node.type.name !== "wpsShape" && sel.node.type.name !== "wpgGroup") {
    if (sel.node.type.name !== "chart") return false;
    const attrs = { ...sel.node.attrs } as Record<string, unknown>;
    const chart = { ...(attrs.chart as Record<string, unknown>) };
    const t = { ...((chart.transformation ?? {}) as Record<string, unknown>) };
    t[axis] = emu;
    chart.transformation = t;
    attrs.chart = chart;
    tr.setNodeMarkup(sel.from, undefined, attrs);
    tr.setSelection(NodeSelection.create(tr.doc, sel.from) as never);
    return true;
  }
  const key = sel.node.type.name === "wpsShape" ? "wpsShape" : "wpgGroup";
  const attrs = { ...sel.node.attrs } as Record<string, unknown>;
  const payload = { ...(attrs[key] as Record<string, unknown>) };
  const t = { ...((payload.transformation ?? {}) as Record<string, unknown>) };
  t[axis] = emu;
  payload.transformation = t;
  attrs[key] = payload;
  tr.setNodeMarkup(sel.from, undefined, attrs);
  tr.setSelection(NodeSelection.create(tr.doc, sel.from) as never);
  return true;
}

/** ── Picture Format > Adjust ──
 *  The pixel-adjustment commands retarget the selected image, merging into
 *  the attrs.blipEffects / attrs.outline office-open objects the projection
 *  reads (blipEffects → the pixel filter and opacity, outline → the border
 *  stroke). A preset equal to "no adjustment" deletes its field and an
 *  emptied effects object drops off the attrs, so flipping between Word's
 *  presets never leaves husk groups behind. */

/** Stamp the selected image's attrs through `mutate`, keeping the
 *  NodeSelection. False when the selection is not an image. */
function patchPicture(
  state: EditorState,
  tr: Transaction,
  mutate: (attrs: Record<string, unknown>) => void,
): boolean {
  const sel = state.selection;
  if (!(sel instanceof NodeSelection) || sel.node.type.name !== "image") return false;
  const attrs = { ...sel.node.attrs };
  mutate(attrs);
  tr.setNodeMarkup(sel.from, undefined, attrs);
  tr.setSelection(NodeSelection.create(tr.doc, sel.from) as never);
  return true;
}

/** The changes Reset Picture discards: the blip adjustments, the border, the
 *  effect list, and the crop (traded back to the full source — the extent
 *  math reads the old fractions, so this must run before the field is
 *  gone, which tradeCropExtent itself arranges). */
function resetPictureChanges(attrs: Record<string, unknown>): void {
  delete attrs.blipEffects;
  delete attrs.outline;
  delete attrs.effects;
  tradeCropExtent(attrs, 0, 0, 0, 0);
}

/** Merge a blipEffects mutation into the image's attrs, pruning emptied
 *  effect groups and a fully-cleared effects object (absent reads as
 *  unadjusted downstream). */
function patchBlipEffects(
  attrs: Record<string, unknown>,
  patch: (blip: Record<string, unknown>) => void,
): void {
  const blip = { ...((attrs.blipEffects ?? {}) as Record<string, unknown>) };
  patch(blip);
  for (const key of ["luminance", "hsl", "alphaModulateFixed"]) {
    const group = blip[key];
    if (group != null && typeof group === "object" && Object.keys(group).length === 0)
      delete blip[key];
  }
  if (Object.keys(blip).length === 0) delete attrs.blipEffects;
  else attrs.blipEffects = blip;
}

/** The extreme zIndex among the document's floating drawings that share the
 *  target's band (behind / in front of text), excluding the target itself.
 *  Bring-to-front sends max + 1, send-to-back min − 1; the identities read as
 *  "no competitor" when the band holds the target alone (−1 maxes below any
 *  z, ∞ mins above the clamped floor), so both commands no-op there. */
function bandExtreme(
  state: EditorState,
  target: NonNullable<ReturnType<typeof floatingDrawingAt>>,
  max: boolean,
): number {
  const behind = floatingOf(target).behindDocument === true;
  let extreme = max ? -1 : Number.POSITIVE_INFINITY;
  state.doc.descendants((node, pos) => {
    if (pos === target.pos) return;
    const attrs = node.attrs as Record<string, unknown>;
    const floating =
      node.type.name === "image"
        ? (attrs.floating as Record<string, unknown> | undefined)
        : node.type.name === "wpsShape"
          ? ((attrs.wpsShape as Record<string, unknown> | null | undefined)?.floating as
              | Record<string, unknown>
              | undefined)
          : undefined;
    if (!floating || (floating.behindDocument === true) !== behind) return;
    if (typeof floating.zIndex !== "number") return;
    extreme = max ? Math.max(extreme, floating.zIndex) : Math.min(extreme, floating.zIndex);
  });
  return extreme;
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

/** The Shadow gallery's direction picks → DrawingML clockwise degrees
 *  (a:outerShdw @dir, measured from the 3-o'clock position). */
const SHADOW_DIRECTIONS: Readonly<Record<string, number>> = {
  "shadow-right": 0,
  "shadow-lower-right": 45,
  "shadow-bottom": 90,
  "shadow-lower-left": 135,
  "shadow-left": 180,
  "shadow-upper-left": 225,
  "shadow-top": 270,
  "shadow-upper-right": 315,
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
  // The Chinese unit words Word's zh boxes display and accept typed
  // ("5 厘米" in the Size boxes).
  ["磅", 20],
  ["派卡", 240],
  ["英寸", 1440],
  ["毫米", 1440 / 25.4],
  ["厘米", 1440 / 2.54],
  ["像素", 15],
];
function parseMeasureTwip(v: unknown): number | null {
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v !== "string") return null;
  const bare = Number(v);
  if (Number.isFinite(bare)) return bare;
  const m = /^(-?[\d.]+)\s*(\S+)$/.exec(v.trim());
  if (!m) return null;
  const unit = MEASURE_TWIP_UNITS.find(([u]) => u.toLowerCase() === m[2]!.toLowerCase());
  return unit ? Number(m[1]) * unit[1] : null;
}

/** parseMeasureEmu — parseMeasureTwip's EMU twin, for the DrawingML length
 *  fields (shape/group transformation widths, which may arrive as a
 *  UniversalMeasure string "5cm" or a bare EMU number). */
const MEASURE_EMU_UNITS: ReadonlyArray<readonly [string, number]> = [
  ["pt", 12700],
  ["pc", 152400],
  ["in", 914400],
  ["mm", 36000],
  ["cm", 360000],
  ["px", 9525],
  ["磅", 12700],
  ["派卡", 152400],
  ["英寸", 914400],
  ["毫米", 36000],
  ["厘米", 360000],
  ["像素", 9525],
];
function parseMeasureEmu(v: unknown): number | null {
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v !== "string") return null;
  const bare = Number(v);
  if (Number.isFinite(bare)) return bare;
  const m = /^(-?[\d.]+)\s*(\S+)$/.exec(v.trim());
  if (!m) return null;
  const unit = MEASURE_EMU_UNITS.find(([u]) => u.toLowerCase() === m[2]!.toLowerCase());
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

/** Word's Insert Chart default: a clustered column frame with the same
 *  sample data Word seeds (Edit Data rewrites it), at Word's 5" × 3" extent
 *  (EMU — the chart payload's transformation is EMU, like every drawing). */
const DEFAULT_CHART = {
  type: "column",
  categories: ["Category 1", "Category 2", "Category 3", "Category 4"],
  series: [
    { name: "Series 1", values: [4.3, 2.5, 3.5, 4.5] },
    { name: "Series 2", values: [2.4, 4.4, 1.8, 2.8] },
    { name: "Series 3", values: [2, 2, 3, 5] },
  ],
  showLegend: true,
  legendPosition: "bottom",
  transformation: { width: 4572000, height: 2743200 },
};
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
      // Insert Chart — Word's default frame (clustered column, sample data,
      // 5" × 3") as an inline atom at the caret. The Edit Data dialog and the
      // chart-format tab rewrite the payload from here.
      chart:
        () =>
        ({ state, dispatch }) => {
          const node = state.schema.nodes.chart.create({ chart: { ...DEFAULT_CHART } });
          if (!node) return false;
          if (dispatch) dispatch(state.tr.replaceSelectionWith(node).scrollIntoView());
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
            if (targetCellPos > 0) {
              tr.setSelection(TextSelection.near(tr.doc.resolve(targetCellPos + 1)));
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
            if (targetCellPos > 0) {
              tr.setSelection(TextSelection.near(tr.doc.resolve(targetCellPos + 1)));
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
          const anchor = tableAncestry(state);
          if (!anchor || anchor.cellAt < 0) return false;
          if (dispatch) {
            const cellPos = state.selection.$from.before(anchor.cellAt);
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
      // The caret's column across all rows: a cell selection over the whole
      // column (span-aware grid math comes free from prosemirror-tables).
      "select-table-column":
        () =>
        ({ state, dispatch }) => {
          const anchor = tableAncestry(state);
          if (!anchor || anchor.cellAt < 0) return false;
          if (dispatch) {
            const cellPos = state.selection.$from.before(anchor.cellAt);
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
      // Word's Repeat Header Rows — one press marks the whole pick (the
      // selected rows gain or lose tblHeader together, the anchor row's
      // current state deciding which way).
      "repeat-header-rows":
        () =>
        ({ state, dispatch }) => {
          const targets = tableTargets(state);
          if (!targets) return false;
          if (dispatch) {
            const firstRow = Math.min(...targets.rows);
            const next = !(targets.tableNode.child(firstRow)!.attrs.tableHeader as boolean);
            const tr = state.tr;
            stampRows(tr, targets, (row) => ({ ...row.attrs, tableHeader: next }));
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
      // Word's Merge Cells over the selection's bounding rectangle. Each
      // spanned row folds its cells into one: the row's first cell takes
      // columnSpan = width, rows below the first take verticalMerge
      // "continue" (their content stays put — the layout folds continue
      // cells into the restart cell, so nothing is lost). Grid math is
      // cellIndex-approximate, so a span-mismatched row is left untouched.
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
          const tableNode = $from.node(fromA.tableAt);
          const grid = (tableNode.attrs.columnWidths as number[] | null)?.length ?? 0;
          const c1 = Math.min($from.index(fromA.rowAt), $to.index(toA.rowAt));
          const c2 = Math.max($from.index(fromA.rowAt), $to.index(toA.rowAt));
          const rowFrom = Math.min($from.index(fromA.tableAt), $to.index(toA.tableAt));
          const rowTo = Math.max($from.index(fromA.tableAt), $to.index(toA.tableAt));
          if (rowFrom === rowTo && c1 === c2) return false;
          if (dispatch) {
            const tablePos = $from.before(fromA.tableAt);
            // Row indices into the table's children — the ancestry depths are
            // not indexes (a depth-2 rowAt would address the last row).
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
          const cols = tableNode.child(0)?.childCount ?? 0;
          if (cols === 0) return false;
          let widths = tableNode.attrs.columnWidths as number[] | null;
          if (!widths || widths.length === 0) {
            widths = Array.from({ length: cols }, () => 2880);
          }
          if (dispatch) {
            const sum = widths.reduce((a, b) => a + b, 0);
            const even = Math.floor(sum / cols);
            const next = Array.from({ length: cols }, (_, c) =>
              c === cols - 1 ? sum - even * (cols - 1) : even,
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
      // Word's Distribute Rows: the selected rows' declared heights split
      // their total evenly (the last row absorbs the rounding remainder); a
      // bare caret (or a single-row pick) applies table-wide, like Distribute
      // Columns. Rows without a declared height stay auto — there is nothing
      // to redistribute from, so at least two rows need one.
      "distribute-rows":
        () =>
        ({ state, dispatch }) => {
          const fromA = ancestryAt(state.selection.$from);
          const toA = ancestryAt(state.selection.$to);
          if (!fromA || !toA || fromA.rowAt < 0 || toA.rowAt < 0) return false;
          const { $from, $to } = state.selection;
          if ($from.before(fromA.tableAt) !== $to.before(toA.tableAt)) return false;
          const tableNode = $from.node(fromA.tableAt);
          const tablePos = $from.before(fromA.tableAt);
          const rowFrom = $from.index(fromA.tableAt);
          const rowTo = $to.index(toA.tableAt);
          const last = rowFrom === rowTo ? tableNode.childCount - 1 : rowTo;
          // Ancestry depths are not indexes — resolve row positions by child
          // offsets (the merge-cells pattern); markup writes keep positions
          // stable, so a forward walk stays valid.
          const rows: { pos: number; height: number }[] = [];
          let rowPos = tablePos + 1;
          for (let r = 0; r <= last; r += 1) {
            const row = tableNode.child(r);
            const h = row.attrs.height as { value?: unknown } | null;
            if (r >= rowFrom && h && typeof h.value === "number" && h.value > 0) {
              rows.push({ pos: rowPos, height: h.value });
            }
            rowPos += row.nodeSize;
          }
          if (rows.length < 2) return false;
          if (dispatch) {
            const sum = rows.reduce((a, b) => a + b.height, 0);
            const even = Math.floor(sum / rows.length);
            const tr = state.tr;
            rows.forEach(({ pos }, i) => {
              const row = tr.doc.nodeAt(pos)!;
              tr.setNodeMarkup(pos, undefined, {
                ...row.attrs,
                height: {
                  value: i === rows.length - 1 ? sum - even * (rows.length - 1) : even,
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
      // Table Properties dialog's OK — writes the caret table's w:jc
      // alignment and w:tblInd indent in one transaction ("left" and 0
      // commit null, OOXML's absent-attribute default).
      "table-properties-apply":
        (patch) =>
        ({ state, dispatch }) => {
          if (
            !patch ||
            (patch.alignment !== "left" &&
              patch.alignment !== "center" &&
              patch.alignment !== "right") ||
            typeof patch.indent !== "number" ||
            patch.indent < 0
          )
            return false;
          const anchor = tableAncestry(state);
          if (!anchor) return false;
          const { $from } = state.selection;
          const tableNode = $from.node(anchor.tableAt);
          if (dispatch) {
            dispatch(
              state.tr
                .setNodeMarkup($from.before(anchor.tableAt), undefined, {
                  ...tableNode.attrs,
                  alignment: patch.alignment === "left" ? null : patch.alignment,
                  indent: patch.indent > 0 ? Math.round(patch.indent) : null,
                })
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
      // Re-home a dragged floating drawing to the paragraph under the drop
      // point (Word re-anchors on drag): value is JSON { to, h, v } — `to` a
      // content position inside the target paragraph, {h, v} the drawing's
      // new offsets in EMU. A drop far above/below the anchor pushes the
      // anchor paragraph itself off its page when the wrap zone re-flows, and
      // the drawing then pins where the anchor landed instead of where it was
      // dropped — anchoring beside the drop keeps the layout stable. One
      // transaction: delete + insert + selection restore.
      "reanchor-drawing":
        (value?) =>
        ({ state, tr }) => {
          if (!value) return false;
          const target = floatingDrawingAt(state);
          if (!target) return false;
          let parsed: { to?: number; h?: number; v?: number };
          try {
            parsed = JSON.parse(value) as { to?: number; h?: number; v?: number };
          } catch {
            return false;
          }
          if (typeof parsed.to !== "number") return false;
          const node = tr.doc.nodeAt(target.pos);
          if (!node) return false;
          const floating = floatingOf(target);
          const h = floating.horizontalPosition as Record<string, unknown> | undefined;
          const v = floating.verticalPosition as Record<string, unknown> | undefined;
          if (!h || !v) return false;
          // Validate the drop target before touching the tr: the command
          // manager dispatches the transaction regardless of the command's
          // return value, so a decline after the delete would still land it
          // and the drawing would vanish.
          if (parsed.to < 0 || parsed.to > state.doc.content.size) return false;
          if (state.doc.resolve(parsed.to).parent.type.name !== "paragraph") return false;
          const next = {
            ...floating,
            horizontalPosition:
              typeof h.offset === "number" ? { ...h, offset: parsed.h ?? h.offset } : h,
            verticalPosition:
              typeof v.offset === "number" ? { ...v, offset: parsed.v ?? v.offset } : v,
          };
          const moved = node.type.create(withFloating(target, next), node.content);
          tr.delete(target.pos, target.pos + node.nodeSize);
          const at = tr.mapping.map(parsed.to, -1);
          tr.insert(at, moved);
          tr.setSelection(NodeSelection.create(tr.doc, at) as never);
          tr.setMeta("scrollIntoView", false);
          return true;
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
          // The chart painter draws no transformation rotation yet — declining
          // is the honest response (the handles still move the chart via
          // move-drawing).
          if (target.kind === "chart") return false;
          const key = target.kind === "shape" ? "wpsShape" : "wpgGroup";
          const payload = target.attrs[key] as Record<string, unknown>;
          const transformation = {
            ...(payload.transformation as Record<string, unknown> | undefined),
          };
          const current = transformation.rotation;
          transformation.rotation = (typeof current === "number" ? current : 0) + delta;
          return stampAttrs(tr, target, {
            ...target.attrs,
            [key]: { ...payload, transformation },
          });
        },
      // The Size-and-Position dialog's OK: absolute geometry in centimeters
      // (the dialog's display unit), converted to each carrier's native unit —
      // an image sizes in px and offsets in EMU, a shape/group payload in EMU
      // (a group's extent change scales its members through the child
      // coordinate space — walkGroup's childScale, Word's group resize).
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
            // The replacement text rides the image's title attr (altText
            // description on export); empty clears it. A shape payload has no
            // alt-text field, so the dialog's field only lands on images.
            if (typeof patch.altText === "string") {
              if (patch.altText) attrs.title = patch.altText;
              else delete attrs.title;
            }
            return stampAttrs(tr, target, {
              ...attrs,
              floating: applyFloatingExtras(floatingOf(target), patch, offsetHCm, offsetVCm),
            });
          }
          const key = FLOATING_CARRIER[target.kind];
          const payload = { ...(target.attrs[key] as Record<string, unknown>) };
          const t = { ...((payload.transformation ?? {}) as Record<string, unknown>) };
          if (widthCm != null) t.width = cmTo(widthCm, EMU_PER_CM);
          if (heightCm != null) t.height = cmTo(heightCm, EMU_PER_CM);
          if (rotationDeg != null) t.rotation = rotationDeg;
          payload.transformation = t;
          // The shape/group/chart offsets ride the same Floating object as an
          // image's.
          payload.floating = applyFloatingExtras(floatingOf(target), patch, offsetHCm, offsetVCm);
          return stampAttrs(tr, target, { ...target.attrs, [key]: payload });
        },
      // The crop overlay's commit: the selected image's new a:srcRect insets
      // as source fractions, stored as the raw ST_Percentage ints the attrs
      // carry (office-open's parse emits raw ints despite the documented
      // integer percent — mirror cropOf's /100000 read side). The extent
      // follows the kept region (Word's crop). All zero clears.
      "drawing-crop-apply":
        (patch?) =>
        ({ state, tr }) =>
          patch && typeof patch === "object" ? applyCropPatch(state, tr, patch) : false,
      // The context menu's Reset Crop — the same all-zero patch that clears
      // the a:srcRect, so the picture shows its full source again and the
      // extent grows back to it (the crop scale is preserved).
      "drawing-crop-reset":
        () =>
        ({ state, tr }) =>
          applyCropPatch(state, tr, { left: 0, top: 0, right: 0, bottom: 0 }),
      // Word's Crop → Aspect Ratio presets ("1:1", "2:3", …) — the largest
      // centered rect of the ratio inside the kept region, applied through
      // the same crop patch (extent follows the kept fractions).
      "drawing-crop-aspect":
        (value) =>
        ({ state, tr }) =>
          applyCropAspect(state, tr, value),
      // The Picture Format tab's Height/Width boxes: the selected image
      // resizes to the typed measure ("5cm"/"2in"/"120px") — inline and
      // floating pictures alike.
      "drawing-width":
        (value) =>
        ({ state, tr }) =>
          applyImageSize(state, tr, "width", value),
      "drawing-height":
        (value) =>
        ({ state, tr }) =>
          applyImageSize(state, tr, "height", value),
      // ── Picture Format > Adjust（更正/颜色/透明度/边框/重置） ──

      // Corrections: "bright:N" / "contrast:N" (percent points — the attr
      // keys are office-open's LuminanceEffectOptions names; Word's
      // sharpen/soften presets ride the same dialog and stay unmodeled). 0
      // clears the field — the preset grid's "0% (normal)" is a reset.
      "picture-correction":
        (value) =>
        ({ state, tr }) => {
          const [kind, raw] = String(value ?? "").split(":");
          if (kind !== "bright" && kind !== "contrast") return false;
          const amount = Number(raw);
          if (!Number.isFinite(amount) || amount < -100 || amount > 100) return false;
          return patchPicture(state, tr, (attrs) =>
            patchBlipEffects(attrs, (blip) => {
              const lum = { ...((blip.luminance ?? {}) as Record<string, unknown>) };
              if (amount === 0) delete lum[kind];
              else lum[kind] = Math.round(amount);
              blip.luminance = lum;
            }),
          );
        },
      // Color: "saturation:N" (Word's saturation label %; the attrs carry the
      // offset from the neutral 100%) or "none" (no recolor — clears the
      // recolor fields).
      "picture-color":
        (value) =>
        ({ state, tr }) => {
          const v = String(value ?? "");
          return patchPicture(state, tr, (attrs) =>
            patchBlipEffects(attrs, (blip) => {
              if (v === "none") {
                delete blip.hsl;
                delete blip.grayscale;
                return;
              }
              if (!v.startsWith("saturation:")) return;
              const percent = Number(v.slice("saturation:".length));
              if (!Number.isFinite(percent) || percent < 0 || percent > 200) return;
              const hsl = { ...((blip.hsl ?? {}) as Record<string, unknown>) };
              const points = Math.round(percent) - 100;
              if (points === 0) delete hsl.saturation;
              else hsl.saturation = points;
              blip.hsl = hsl;
            }),
          );
        },
      // Transparency: "0".."100" percent (the menu speaks transparency; the
      // attrs carry the alpha modulate — the OPACITY — so 50% transparency
      // writes amount 50). 0 restores fully opaque and drops the field.
      "picture-transparency":
        (value) =>
        ({ state, tr }) => {
          const percent = Number(value);
          if (!Number.isFinite(percent) || percent < 0 || percent > 100) return false;
          return patchPicture(state, tr, (attrs) =>
            patchBlipEffects(attrs, (blip) => {
              const amount = Math.round(100 - percent);
              if (amount >= 100) delete blip.alphaModulateFixed;
              else blip.alphaModulateFixed = { amount };
            }),
          );
        },
      // Border: "none", "color:RRGGBB", "width:P" (points) or "dash:token" —
      // incremental merges onto one outline object, like Word's menu sections
      // which each commit independently. A bare color picks the 1 pt default
      // Word starts with.
      "picture-border":
        (value) =>
        ({ state, tr }) => {
          const v = String(value ?? "");
          return patchPicture(state, tr, (attrs) => {
            if (v === "none") {
              delete attrs.outline;
              return;
            }
            const outline = { ...((attrs.outline ?? {}) as Record<string, unknown>) };
            outline.type = "solid";
            if (v.startsWith("color:")) {
              const color = v.slice(6).toUpperCase();
              if (!/^[0-9A-F]{6}$/.test(color)) return;
              outline.color = color;
              if (typeof outline.width !== "number") outline.width = 12700;
            } else if (v.startsWith("width:")) {
              const pt = Number(v.slice(6));
              if (!Number.isFinite(pt) || pt <= 0 || pt > 12) return;
              outline.width = Math.round(pt * 12700); // EMU to the point
            } else if (v.startsWith("dash:")) {
              // "solid" is the no-dash reset (prstDash has no solid token).
              const dash = v.slice(5);
              if (dash === "solid") delete outline.dash;
              else outline.dash = dash;
            } else return;
            attrs.outline = outline;
          });
        },
      // Word's Reset Picture: discard every change made to the picture — the
      // adjustments (blipEffects), the border, the effect list (the shadow),
      // and the crop (the frame grows back through the crop trade — the same
      // geometry Reset Crop takes; clearing the field alone would strand the
      // shrunken frame). Size, rotation, flips, and the anchor stay — those
      // are Reset Picture and Size's to touch.
      "reset-picture":
        () =>
        ({ state, tr }) =>
          patchPicture(state, tr, resetPictureChanges),
      // Word's Reset Picture and Size: the discard above plus the extent back
      // at the source's natural size. The command layer cannot read a decode
      // (that lives in the browser-side painter), so the host's dispatcher
      // resolves the selected picture's decoded dimensions and passes them in
      // — the change-picture pattern. Without them the command degrades to
      // the plain Reset Picture.
      "reset-picture-size":
        (natural?: { width: number; height: number }) =>
        ({ state, tr }) =>
          patchPicture(state, tr, (attrs) => {
            resetPictureChanges(attrs);
            if (
              natural &&
              natural.width > 0 &&
              natural.height > 0 &&
              Number.isFinite(natural.width) &&
              Number.isFinite(natural.height)
            ) {
              attrs.width = Math.round(natural.width);
              attrs.height = Math.round(natural.height);
            }
          }),
      // Shape Fill: the palette picker's hex (or "none") writing the shape's
      // solid fill. A noFill type reads as unfilled in the projection and
      // serializes <a:noFill>.
      "shape-fill":
        (value) =>
        ({ state, tr }) => {
          const target = shapeAt(state);
          if (!target) return false;
          const v = String(value ?? "");
          const shape = { ...(target.attrs.wpsShape as Record<string, unknown>) };
          if (v === "none") shape.fill = { type: "noFill" };
          else if (/^[0-9A-F]{6}$/i.test(v)) shape.fill = { type: "solid", color: v.toUpperCase() };
          else return false;
          return stampAttrs(tr, target, { ...target.attrs, wpsShape: shape });
        },
      // Shape Outline: same value grammar and incremental-merge semantics as
      // picture-border (none / color: / width: / dash:), onto the shape's own
      // outline. A width or dash on an uncolored outline defaults to black —
      // the projection has no theme-outline fallback to inherit.
      "shape-outline":
        (value) =>
        ({ state, tr }) => {
          const target = shapeAt(state);
          if (!target) return false;
          const v = String(value ?? "");
          const shape = { ...(target.attrs.wpsShape as Record<string, unknown>) };
          if (v === "none") {
            shape.outline = { type: "noFill" };
          } else {
            const outline = { ...((shape.outline ?? {}) as Record<string, unknown>) };
            outline.type = "solid";
            if (v.startsWith("color:")) {
              const color = v.slice(6).toUpperCase();
              if (!/^[0-9A-F]{6}$/.test(color)) return false;
              outline.color = color;
              if (typeof outline.width !== "number") outline.width = 12700;
            } else if (v.startsWith("width:")) {
              const pt = Number(v.slice(6));
              if (!Number.isFinite(pt) || pt <= 0 || pt > 12) return false;
              outline.width = Math.round(pt * 12700);
              if (typeof outline.color !== "string") outline.color = "000000";
            } else if (v.startsWith("dash:")) {
              // "solid" is the no-dash reset (prstDash has no solid token).
              const dash = v.slice(5);
              if (dash === "solid") delete outline.dash;
              else outline.dash = dash;
              if (typeof outline.color !== "string") outline.color = "000000";
            } else return false;
            shape.outline = outline;
          }
          return stampAttrs(tr, target, { ...target.attrs, wpsShape: shape });
        },
      // Shape Effects: Word's Shadow section — "none" clears, a direction
      // pick stamps the preset outer shadow (Word's default offset/blur
      // geometry, 60%-opaque black ink). The other effect families (glow,
      // soft edges, bevel, 3-D) have no painter mapping yet and stay greyed
      // at the menu; an emptied effects object drops off the attrs.
      "shape-effects":
        (value) =>
        ({ state, tr }) => {
          const target = shapeAt(state);
          if (!target) return false;
          const v = String(value ?? "");
          const shape: Record<string, unknown> = {
            ...(target.attrs.wpsShape as Record<string, unknown>),
          };
          const effects = { ...((shape.effects ?? {}) as Record<string, unknown>) };
          if (v === "none") {
            if (!("outerShadow" in effects)) return false;
            delete effects.outerShadow;
          } else {
            const dir = SHADOW_DIRECTIONS[v];
            if (dir == null) return false;
            effects.outerShadow = {
              distance: 25400,
              direction: dir,
              blurRadius: 38100,
              color: { value: "000000", transforms: { alpha: 60 } },
            };
          }
          if (Object.keys(effects).length === 0) delete shape.effects;
          else shape.effects = effects;
          return stampAttrs(tr, target, { ...target.attrs, wpsShape: shape });
        },
      // Word's Text Direction menu: "horizontal" clears bodyPr @vert (the
      // OOXML absent-attribute default), "vertical"/"vertical270" stamp the
      // rotated layouts (Rotate all text 90°/270°). The stacked variants
      // (eastAsianVert…) need per-glyph upright layout the renderer has no
      // model for — greyed at the menu.
      "shape-text-direction":
        (value) =>
        ({ state, tr }) => {
          const target = shapeAt(state);
          if (!target) return false;
          if (value !== "horizontal" && value !== "vertical" && value !== "vertical270")
            return false;
          const shape: Record<string, unknown> = {
            ...(target.attrs.wpsShape as Record<string, unknown>),
          };
          const body = { ...((shape.bodyProperties ?? {}) as Record<string, unknown>) };
          if (value === "horizontal") {
            if (!("vertical" in body)) return false;
            delete body.vertical;
          } else {
            body.vertical = value;
          }
          if (Object.keys(body).length === 0) delete shape.bodyProperties;
          else shape.bodyProperties = body;
          return stampAttrs(tr, target, { ...target.attrs, wpsShape: shape });
        },
      // ── Chart Design — type / legend / data (the contextual tab) ────────
      // Chart Type: value is the ChartType token the renderer draws (the
      // unmodeled types grey out at the menu). Pie/doughnut/scatter/bubble/
      // radar have no series grouping — a hand-me-down stacked flag from the
      // previous type would mis-shape them, so it clears.
      "chart-type":
        (value) =>
        ({ state, tr }) => {
          const target = chartAt(state);
          if (!target || !value) return false;
          const type = value as ChartType;
          const chart = { ...target.chart, type };
          if (!GROUPING_CHART_TYPES.includes(chart.type)) delete chart.grouping;
          return stampChart(tr, target, chart);
        },
      // Legend: "none" hides it (showLegend false — the painter's presence
      // check), any other value is the LegendPosition with showLegend forced.
      "chart-legend":
        (value) =>
        ({ state, tr }) => {
          const target = chartAt(state);
          if (!target || !value) return false;
          const chart = { ...target.chart };
          if (value === "none") {
            chart.showLegend = false;
            delete chart.legendPosition;
          } else {
            chart.showLegend = true;
            chart.legendPosition = value as LegendPosition;
          }
          return stampChart(tr, target, chart);
        },
      // Edit Data dialog's commit — JSON {title?, categories?, series?}.
      // Absent optional fields keep the chart's current ones; an empty title
      // string clears it (the painter drops an empty title anyway).
      "chart-data-apply":
        (value) =>
        ({ state, tr }) => {
          const target = chartAt(state);
          if (!target || !value) return false;
          let parsed: {
            title?: string;
            categories?: string[];
            series?: { name?: string; values?: number[] }[];
          };
          try {
            parsed = JSON.parse(value);
          } catch {
            return false;
          }
          const chart = { ...target.chart };
          if (parsed.title != null) {
            if (parsed.title) chart.title = parsed.title;
            else delete chart.title;
          }
          if (parsed.categories) chart.categories = parsed.categories;
          if (parsed.series)
            // The grid edits category-series values only; a scatter/bubble
            // series (xValues/yValues/bubbleSize) keeps its shape untouched —
            // giving it a values field would make the painter's valuesOf
            // read it as a category series.
            chart.series = parsed.series.map((s, i) => {
              const prev = target.chart.series?.[i];
              if (prev && !("values" in prev)) return prev;
              return {
                ...prev,
                name: s.name ?? prev?.name ?? `Series ${i + 1}`,
                values: s.values ?? ("values" in prev ? [...prev.values] : []),
              };
            }) as ChartOptions["series"];
          return stampChart(tr, target, chart);
        },
      // Delete Key on a sub-selected series (Word): the series leaves the
      // chart, the categories stay for the survivors. The last series
      // declines — an empty plot has nothing left to edit.
      "chart-series-delete":
        (value) =>
        ({ state, tr }) => {
          const target = chartAt(state);
          const index = Number.parseInt(value ?? "", 10);
          if (!target || !Number.isInteger(index)) return false;
          const series = [...(target.chart.series ?? [])];
          if (index < 0 || index >= series.length || series.length <= 1) return false;
          series.splice(index, 1);
          // A mutable copy of the series union doesn't assign back to the
          // union-of-arrays — the write is the boundary.
          return stampChart(tr, target, {
            ...target.chart,
            series: series as ChartOptions["series"],
          });
        },
      // The plot's value-drag commit (Excel's drag-a-point editing): JSON
      // {series, point, value} writes one data point, everything else stays.
      "chart-value-apply":
        (value) =>
        ({ state, tr }) => {
          const target = chartAt(state);
          if (!target || !value) return false;
          let parsed: { series?: number; point?: number; value?: number };
          try {
            parsed = JSON.parse(value);
          } catch {
            return false;
          }
          const { series, point, value: v } = parsed;
          if (
            !Number.isInteger(series) ||
            !Number.isInteger(point) ||
            typeof v !== "number" ||
            !Number.isFinite(v)
          )
            return false;
          const prev = target.chart.series?.[series!];
          if (!prev || !("values" in prev)) return false;
          if (point! < 0 || point! >= prev.values.length) return false;
          const next = prev.values.map((old, pi) => (pi === point ? v : old));
          const all = [...(target.chart.series ?? [])];
          all[series!] = { ...prev, values: next };
          return stampChart(tr, target, {
            ...target.chart,
            series: all as ChartOptions["series"],
          });
        },
      // Swap the source (the Change Picture flow's commit; the file-picker
      // side reads the file into a data URL at the UI layer). The frame keeps
      // its size — the new source stretches into it — and the crop resets
      // (a srcRect described the old source's edges, Word resets on swap).
      "change-picture":
        (src) =>
        ({ state, tr }) =>
          typeof src === "string" && src
            ? patchPicture(state, tr, (attrs) => {
                attrs.src = src;
                delete attrs.crop;
              })
            : false,
      // The pixel tools' commit (see the type above): unlike change-picture
      // the crop survives unless the payload says it was baked in.
      "picture-pixels":
        (value) =>
        ({ state, tr }) => {
          let parsed: { src?: unknown; dropCrop?: unknown };
          try {
            parsed = JSON.parse(String(value ?? ""));
          } catch {
            return false;
          }
          if (typeof parsed.src !== "string" || !parsed.src) return false;
          return patchPicture(state, tr, (attrs) => {
            attrs.src = parsed.src as string;
            if (parsed.dropCrop === true) delete attrs.crop;
          });
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
      // Word's Bring to Front / Send to Back: jump to the band's extreme z —
      // one past the document's highest (lowest) same-band relativeHeight,
      // never moving away from an extreme already held (the max/max, min/min
      // guards keep a repeated click a no-op). Clamped at 0 (ST_RelativeHeight
      // is unsigned); a clamp tie falls back to document order in the painter.
      "bring-to-front":
        () =>
        ({ state, tr }) => {
          const target = floatingDrawingAt(state);
          if (!target) return false;
          const floating = floatingOf(target);
          const current = typeof floating.zIndex === "number" ? floating.zIndex : 0;
          const next = Math.max(current, bandExtreme(state, target, true) + 1);
          return next === current
            ? false
            : stampFloating(tr, target, { ...floating, zIndex: next });
        },
      "send-to-back":
        () =>
        ({ state, tr }) => {
          const target = floatingDrawingAt(state);
          if (!target) return false;
          const floating = floatingOf(target);
          const current = typeof floating.zIndex === "number" ? floating.zIndex : 0;
          const next = Math.min(current, Math.max(0, bandExtreme(state, target, false) - 1));
          return next === current
            ? false
            : stampFloating(tr, target, { ...floating, zIndex: next });
        },
      // Word's Wrap Text menu: In Line with Text turns a floating drawing
      // inline (the floating payload is dropped); In Front of Text / Behind
      // Text clear the wrap (wrapNone) and set behindDoc; the four wrap styles
      // stamp the type and drop behindDoc (Word 2013+ honors it for wrapNone
      // anchors only). An inline drawing taking a flow style turns floating in
      // place — anchored to its own paragraph with no offset (Word's
      // keep-position conversion; every kind converts, the floating lands in
      // its own carrier).
      wrap:
        (value) =>
        ({ state, tr }) => {
          const target = floatingDrawingAt(state);
          if (!target) {
            // Inline drawing: "inline" is a no-op (Word greys the row); a
            // flow style converts.
            if (value === "inline") return false;
            const inline = inlineDrawingAt(state);
            if (!inline) return false;
            // positionH has no "paragraph" token (ST_RelFromH) — Word's
            // keep-position conversion anchors the column horizontally, the
            // paragraph vertically, with Square's 0.125" side distances.
            const floating: Record<string, unknown> = {
              horizontalPosition: { relative: "column", offset: 0 },
              verticalPosition: { relative: "paragraph", offset: 0 },
              behindDocument: false,
            };
            if (value === "front" || value === "behind") {
              floating.behindDocument = value === "behind";
            } else if (value === "square" || value === "tight" || value === "through") {
              floating.wrap = { type: value };
              // Word's Square conversion distances: 0.125" left/right, none
              // above/below (Top-and-Bottom carries no side distances).
              floating.margins = { left: 114300, right: 114300 };
            } else if (value === "top-bottom") {
              floating.wrap = { type: "topAndBottom" };
            } else {
              return false;
            }
            return stampFloating(tr, inline, floating);
          }
          const floating = { ...floatingOf(target) };
          if (value === "inline") {
            const attrs = { ...target.attrs };
            if (target.kind === "image") delete attrs.floating;
            else {
              const shape = { ...(attrs.wpsShape as Record<string, unknown>) };
              delete shape.floating;
              attrs.wpsShape = shape;
            }
            return stampAttrs(tr, target, attrs);
          }
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
      // clockwise-positive); the flips toggle the mirror flags. Inline
      // pictures rotate too (Word keeps them inline). The attrs live in two
      // places — an image carries rotation/flipH/flipV on its top level (a
      // tri-state: null omits, true/false emit explicit bytes), a shape
      // mirrors them inside its transformation.
      rotate:
        (value) =>
        ({ state, tr }) => {
          const target = floatingDrawingAt(state) ?? inlineImageAt(state);
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
          const key = target.kind === "group" ? "wpgGroup" : "wpsShape";
          const payload = { ...(target.attrs[key] as Record<string, unknown>) };
          const t = { ...((payload.transformation ?? {}) as Record<string, unknown>) };
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
          payload.transformation = t;
          return stampAttrs(tr, target, { ...target.attrs, [key]: payload });
        },
      // Word's Position gallery: the nine-cell grid stamps margin-relative
      // align tokens on both axes. A fresh position object per stamp — align
      // and offset are mutually exclusive, so a stale offset must not
      // survive next to the new align. An inline drawing takes a cell by
      // converting to a floating one first (Word's gallery converts on
      // click), which is just this stamp — the inline payload carries no
      // Floating to preserve.
      position:
        (value) =>
        ({ state, tr }) => {
          const spec = POSITION_ALIGN[value ?? ""];
          if (!spec) return false;
          const target = floatingDrawingAt(state) ?? inlineDrawingAt(state);
          if (!target) return false;
          return stampFloating(tr, target, {
            ...floatingOf(target),
            horizontalPosition: { relative: "margin", align: spec.h },
            verticalPosition: { relative: "margin", align: spec.v },
          });
        },
      // The Align menu: Word's both-axes margin alignment — the horizontal
      // trio (the primary face defaults to left) and the vertical trio, each
      // stamping its own axis and leaving the other untouched. The two
      // distributes need multi-selection and stay greyed at the menu layer.
      "align-objects":
        (value) =>
        ({ state, tr }) => {
          const h =
            value === "center" || value === "right"
              ? value
              : value === "left" || value == null || value === ""
                ? "left"
                : null;
          // The menu says "middle" (Word's label); OOXML's vertical token is
          // "center" — same translation the position gallery's rows make.
          const v =
            h == null && (value === "top" || value === "middle" || value === "bottom")
              ? value === "middle"
                ? "center"
                : value
              : null;
          if (h == null && v == null) return false;
          const target = floatingDrawingAt(state);
          if (!target) return false;
          return stampFloating(tr, target, {
            ...floatingOf(target),
            ...(h != null ? { horizontalPosition: { relative: "margin", align: h } } : {}),
            ...(v != null ? { verticalPosition: { relative: "margin", align: v } } : {}),
          });
        },
      // Word's Group on a multi-selection: one wpgGroup node replaces the
      // members (at the document-first member's position — the anchor that
      // contributes the floating). The fresh group is 1:1 (chOff 0, chExt =
      // ext), so first grouping is lossless; the anchor's offsets shift by
      // the union's top-left minus the anchor's box, keeping the union where
      // its members sat on the page.
      "drawing-group":
        (payload?) =>
        ({ state, tr }) => {
          const members = multiMembersOf(state, payload);
          if (!members) return false;
          const anchor = members[0]!;
          const anchorFloating = floatingOf(anchor.target);
          const h = anchorFloating.horizontalPosition as Record<string, unknown> | undefined;
          const v = anchorFloating.verticalPosition as Record<string, unknown> | undefined;
          // An align-anchored float has no offset to shift the union onto —
          // decline rather than inventing a base.
          if (typeof h?.offset !== "number" || typeof v?.offset !== "number") return false;
          const union = unionBox(members.map((m) => m.box));
          const emu = (n: number): number => Math.round(n * EMU_PER_PX);
          const childAttrs = members.map(({ target, box }) => {
            const node = state.doc.nodeAt(target.pos)!;
            return memberInGroupAttrs(
              target.kind,
              node.attrs as Record<string, unknown>,
              freshChildEmu(box, union),
            );
          });
          // A chart member has no in-group carrier — decline the grouping.
          if (childAttrs.some((a) => a == null)) return false;
          const groupNode = state.schema.nodes.wpgGroup.create(
            {
              wpgGroup: {
                transformation: { width: emu(union.width), height: emu(union.height) },
                floating: offsetFloating(
                  anchorFloating,
                  emu(union.x - anchor.box.x),
                  emu(union.y - anchor.box.y),
                ),
                childOffsetX: 0,
                childOffsetY: 0,
                childExtentWidth: emu(union.width),
                childExtentHeight: emu(union.height),
              },
            },
            members.map(({ target }, i) => {
              const node = state.doc.nodeAt(target.pos)!;
              return node.type.create(childAttrs[i]!, node.content, node.marks);
            }),
          );
          // Delete the non-anchor members first (descending — positions stay
          // valid), then replace the anchor in place.
          for (let i = members.length - 1; i >= 1; i -= 1) {
            const { target } = members[i]!;
            const node = state.doc.nodeAt(target.pos)!;
            tr.delete(target.pos, target.pos + node.nodeSize);
          }
          const anchorNode = state.doc.nodeAt(anchor.target.pos)!;
          tr.replaceWith(anchor.target.pos, anchor.target.pos + anchorNode.nodeSize, groupNode);
          tr.setSelection(NodeSelection.create(tr.doc, anchor.target.pos));
          return true;
        },
      // Word's Ungroup on a floating group: each member returns to a
      // standalone floating node at the position the group drew it (the
      // group's offsets plus the child-space position through ext/chExt).
      // Pure attrs — no page geometry. Passthrough members (charts, content
      // parts) have no standalone carrier, so such a group declines.
      // Word fires it from inside an entered group too — a member's
      // NodeSelection resolves to its nearest wpgGroup ancestor (one level).
      "drawing-ungroup":
        () =>
        ({ state, tr }) => {
          let target = floatingDrawingAt(state);
          if ((!target || target.kind !== "group") && state.selection instanceof NodeSelection) {
            const $from = state.doc.resolve(state.selection.from);
            for (let d = $from.depth; d > 0; d--) {
              if ($from.node(d).type.name === "wpgGroup") {
                target = drawingAtPos(state.doc, $from.before(d));
                break;
              }
            }
          }
          if (!target || target.kind !== "group") return false;
          const group = target.attrs.wpgGroup as Record<string, unknown>;
          const t = (group.transformation ?? {}) as Record<string, unknown>;
          const extW = parseMeasureEmu(t.width);
          const extH = parseMeasureEmu(t.height);
          if (extW == null || extH == null) return false;
          const chW = group.childExtentWidth;
          const chH = group.childExtentHeight;
          const chExt =
            typeof chW === "number" && chW > 0 && typeof chH === "number" && chH > 0
              ? { x: chW, y: chH }
              : undefined;
          const chOff = {
            x: typeof group.childOffsetX === "number" ? group.childOffsetX : 0,
            y: typeof group.childOffsetY === "number" ? group.childOffsetY : 0,
          };
          const floating = floatingOf(target);
          const h = floating.horizontalPosition as Record<string, unknown> | undefined;
          const v = floating.verticalPosition as Record<string, unknown> | undefined;
          if (typeof h?.offset !== "number" || typeof v?.offset !== "number") return false;
          const groupNode = state.doc.nodeAt(target.pos)!;
          const out: PMNode[] = [];
          const members: PMNode[] = [];
          groupNode.forEach((child) => members.push(child));
          for (const member of members) {
            const child = childBoxOf(member);
            if (!child) return false;
            const node = memberOutNode(
              member,
              floating,
              memberEmuOf(child, chOff, { x: extW, y: extH }, chExt),
            );
            if (!node) return false;
            out.push(node);
          }
          tr.replaceWith(target.pos, target.pos + groupNode.nodeSize, out);
          tr.setSelection(NodeSelection.create(tr.doc, target.pos));
          return true;
        },
      // Word's Distribute Horizontally/Vertically: equal gaps between the
      // selected drawings' page boxes (the outer two hold still, the middles
      // re-space). Deltas are differences, so page-px × EMU_PER_PX lands
      // straight onto the floating offsets regardless of each anchor base.
      // payload is JSON {members: [{pos, box}]} (the host assembles it from
      // the multi-selection overlay; the ribbon dispatch carries only "h"/"v").
      "drawing-distribute":
        (value, payload) =>
        ({ state, tr }) => {
          if (value !== "h" && value !== "v") return false;
          const members = multiMembersOf(state, payload);
          if (!members) return false;
          const axis: "x" | "y" = value === "h" ? "x" : "y";
          const span: "width" | "height" = value === "h" ? "width" : "height";
          const sorted = [...members].sort((a, b) => a.box[axis] - b.box[axis]);
          const first = sorted[0]!;
          const last = sorted[sorted.length - 1]!;
          const gap =
            (last.box[axis] +
              last.box[span] -
              first.box[axis] -
              sorted.reduce((sum, m) => sum + m.box[span], 0)) /
            (sorted.length - 1);
          let cursor = first.box[axis];
          for (const { target, box } of sorted) {
            const delta = Math.round((cursor - box[axis]) * EMU_PER_PX);
            if (delta !== 0) {
              const floating = floatingOf(target);
              const h = floating.horizontalPosition as Record<string, unknown> | undefined;
              const v = floating.verticalPosition as Record<string, unknown> | undefined;
              if (typeof h?.offset !== "number" || typeof v?.offset !== "number") return false;
              stampFloating(
                tr,
                target,
                offsetFloating(floating, value === "h" ? delta : 0, value === "v" ? delta : 0),
              );
            }
            cursor += box[span] + gap;
          }
          // Keep the multi-selection's anchor (the document-first member)
          // selected — stampFloating left the last moved member selected.
          tr.setSelection(NodeSelection.create(tr.doc, members[0]!.target.pos));
          return true;
        },
    };
  },
});
