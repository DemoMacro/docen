import type { ImageAttrs } from "@docen/docx";
import { BULLET_GLYPHS, nextOrderedReference, ORDERED_FORMATS } from "@docen/docx";
import { Extension } from "@docen/docx/core";
import type { Node as PMNode } from "@tiptap/pm/model";
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
      link: (href?: string) => ReturnType;
      style: (styleId?: string) => ReturnType;
      // Editing
      "change-case": (mode?: string) => ReturnType;
      sort: () => ReturnType;
      "multilevel-list": (level?: string) => ReturnType;
      // Picture — names align to Office.js InlinePicture (delete / left / top).
      "delete-picture": () => ReturnType;
      "position-picture": (value?: string) => ReturnType;
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
  "text-direction",
  "convert-to-text",
  "link",
  "style",
  "undo",
  "redo",
  "change-case",
  "sort",
  "multilevel-list",
  "delete-picture",
  "position-picture",
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
  const { table, tableRow, tableCell } = state.schema.nodes;
  const { $from } = state.selection;
  let tableAt = -1;
  let rowAt = -1;
  let cellAt = -1;
  for (let d = $from.depth; d > 0; d -= 1) {
    const node = $from.node(d);
    if (node.type === table && tableAt < 0) tableAt = d;
    else if (node.type === tableRow && rowAt < 0) rowAt = d;
    else if (node.type === tableCell && cellAt < 0) cellAt = d;
  }
  return tableAt < 0 ? null : { tableAt, rowAt, cellAt };
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

type TableBordersLike = Record<string, { style: string; size: number; color: string } | undefined>;
const GRID_BORDER = { style: "single", size: 4, color: "auto" };
const NO_BORDER = { style: "none", size: 0, color: "auto" };

/** Word's Table Styles gallery stand-ins — border presets the canvas paints
 *  today (a full style system waits on the styles pane). */
const TABLE_STYLE_PRESETS: Record<string, TableBordersLike> = {
  "grid-table": {
    top: GRID_BORDER,
    bottom: GRID_BORDER,
    left: GRID_BORDER,
    right: GRID_BORDER,
    insideHorizontal: GRID_BORDER,
    insideVertical: GRID_BORDER,
  },
  "light-list": {
    top: GRID_BORDER,
    bottom: GRID_BORDER,
    left: GRID_BORDER,
    right: GRID_BORDER,
    insideHorizontal: GRID_BORDER,
    insideVertical: NO_BORDER,
  },
  "no-vertical": {
    bottom: GRID_BORDER,
    insideHorizontal: GRID_BORDER,
    left: NO_BORDER,
    right: NO_BORDER,
    top: NO_BORDER,
  },
  "no-border": {
    bottom: NO_BORDER,
    insideHorizontal: NO_BORDER,
    insideVertical: NO_BORDER,
    left: NO_BORDER,
    right: NO_BORDER,
    top: NO_BORDER,
  },
};

/** Border-side stamps for the Layout/Design borders dropdown — value matches
 *  the Home border menu (none/bottom/top/left/right/all/outside). */
function tableBordersStamp(
  value: string,
  current: TableBordersLike | null,
): TableBordersLike | null {
  if (value === "none") return TABLE_STYLE_PRESETS["no-border"]!;
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
      "table-style":
        (value) =>
        ({ state, dispatch }) => {
          const preset = value ? TABLE_STYLE_PRESETS[value] : undefined;
          if (!preset) return false;
          return stampTableBorders(state, dispatch, preset);
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
    };
  },
});
