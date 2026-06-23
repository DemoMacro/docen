import type { Editor } from "@docen/docx/core";

/**
 * Ribbon command event name (kebab-case) → Tiptap command invocation.
 *
 * Event names align with the `RIBBON_ICONS` keys and the `event`
 * attribute set on `<docen-ribbon-button>` / `<docen-ribbon-combobox>`. A few
 * entries (font-name/font-size/style) carry a `value` from comboboxes.
 *
 * This map is document-specific: workbook (RevoGrid) and presentation
 * (LeaferJS) have their own engines and do not reuse it. Ribbon buttons whose
 * command isn't wired yet simply have no entry here — `dispatchRibbonCommand`
 * ignores unknown events, so the visual skeleton stays complete without
 * throwing.
 */
type RibbonCommand = (editor: Editor, value?: string) => void;

const RIBBON_COMMAND_MAP: Readonly<Record<string, RibbonCommand>> = {
  // Font marks
  bold: (editor) => editor.chain().focus().toggleBold().run(),
  italic: (editor) => editor.chain().focus().toggleItalic().run(),
  underline: (editor) => editor.chain().focus().toggleUnderline().run(),
  strike: (editor) => editor.chain().focus().toggleStrike().run(),
  subscript: (editor) => editor.chain().focus().toggleSubscript().run(),
  superscript: (editor) => editor.chain().focus().toggleSuperscript().run(),
  highlight: (editor) => editor.chain().focus().toggleHighlight().run(),
  code: (editor) => editor.chain().focus().toggleCode().run(),
  "clear-format": (editor) => editor.chain().focus().unsetAllMarks().clearNodes().run(),
  // Font family / size — applied as textStyle mark attrs (`font` = name,
  // `size` = points). The combobox dispatches the chosen value; grow/shrink
  // step the current size by 2pt.
  "font-name": (editor, value) =>
    editor.chain().focus().setMark("textStyle", { font: value }).run(),
  "font-size": (editor, value) =>
    editor
      .chain()
      .focus()
      .setMark("textStyle", { size: value ? Number(value) : null })
      .run(),
  "grow-font": (editor) =>
    editor
      .chain()
      .focus()
      .setMark("textStyle", { size: currentSize(editor) + 2 })
      .run(),
  "shrink-font": (editor) =>
    editor
      .chain()
      .focus()
      .setMark("textStyle", { size: Math.max(1, currentSize(editor) - 2) })
      .run(),
  // Paragraph / alignment
  "align-left": (editor) => editor.chain().focus().setTextAlign("left").run(),
  "align-center": (editor) => editor.chain().focus().setTextAlign("center").run(),
  "align-right": (editor) => editor.chain().focus().setTextAlign("right").run(),
  justify: (editor) => editor.chain().focus().setTextAlign("justify").run(),
  // Paragraph indent / line-spacing / shading / border — applied as nested
  // office-open attrs (indent/spacing/shading/border) on the current block.
  // Both paragraph and heading carry these attrs; Word-default values (see
  // the helpers below).
  "indent-increase": (editor) => adjustIndent(editor, INDENT_STEP_TWIPS),
  "indent-decrease": (editor) => adjustIndent(editor, -INDENT_STEP_TWIPS),
  "line-spacing": (editor, value) => applyLineSpacing(editor, value),
  shading: (editor, value) => applyShading(editor, value),
  "font-color": (editor, value) => applyFontColor(editor, value),
  border: (editor, value) => applyBorder(editor, value),
  "bullet-list": (editor) => editor.chain().focus().toggleBulletList().run(),
  "ordered-list": (editor) => editor.chain().focus().toggleOrderedList().run(),
  "task-list": (editor) => editor.chain().focus().toggleTaskList().run(),
  blockquote: (editor) => editor.chain().focus().toggleBlockquote().run(),
  "horizontal-rule": (editor) => editor.chain().focus().setHorizontalRule().run(),
  // Insert atoms — official extension commands. setPageBreak splits the
  // paragraph so the paginator reflows the trailing content to the next page.
  "page-break": (editor) => editor.chain().focus().setPageBreak().run(),
  "column-break": (editor) => editor.chain().focus().setColumnBreak().run(),
  "section-break": (editor) => editor.chain().focus().setSectionBreak().run(),
  // Insert a 3×3 table (Word's default Insert > Table preset).
  "insert-table": (editor) =>
    editor.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run(),
  // Wrap the selection in a link (empty selection → link around the URL text).
  link: (editor, value) => {
    const href = value || (typeof window !== "undefined" && window.prompt("Link URL")) || "";
    if (!href) return;
    editor.chain().focus().extendMarkRange("link").setLink({ href }).run();
  },
  // Style gallery (combobox-driven): value picks the block style.
  style: (editor, value) => applyStyle(editor, value),
  // History
  undo: (editor) => editor.chain().focus().undo().run(),
  redo: (editor) => editor.chain().focus().redo().run(),
};

/** Current font size at the selection (textStyle.size, in points); falls back to
 *  11pt (Word's body default) when the selection has no explicit size. */
function currentSize(editor: Editor): number {
  const size = editor.getAttributes("textStyle").size;
  return typeof size === "number" ? size : 11;
}

/** HeadingLevel literals office-open lifts into `paragraph.heading` (pStyle
 *  val → Tiptap level). Title shares level 1 with Heading1. */
const HEADING_LEVEL_BY_STYLE: Readonly<Record<string, 1 | 2 | 3 | 4 | 5 | 6>> = {
  Heading1: 1,
  Heading2: 2,
  Heading3: 3,
  Heading4: 4,
  Heading5: 5,
  Heading6: 6,
  Title: 1,
};

/** Apply a named block style from the Styles gallery. `value` is a pStyle id
 *  (e.g. "Normal", "Heading1", "Title", or a custom paragraph-style id) that
 *  round-trips via the Paragraph/Heading `styleId` attr. A HeadingLevel id
 *  switches the block to a heading and stamps styleId; everything else becomes
 *  a paragraph carrying styleId so the injected document CSS applies. */
function applyStyle(editor: Editor, value?: string): void {
  const styleId = (value ?? "").trim();
  if (!styleId || styleId === "Normal") {
    editor
      .chain()
      .focus()
      .setParagraph()
      .updateAttributes("paragraph", { styleId: styleId || null })
      .run();
    return;
  }
  const level = HEADING_LEVEL_BY_STYLE[styleId];
  if (level) {
    editor.chain().focus().setHeading({ level }).updateAttributes("heading", { styleId }).run();
    return;
  }
  editor.chain().focus().setParagraph().updateAttributes("paragraph", { styleId }).run();
}

// ── Paragraph indent / line-spacing / shading / border helpers ──
// None of these are arbitrary — each maps to an OOXML unit scale or a Word
// default. The ECMA-376 scales (240, 8) are spec-defined and irreducible;
// inch→twip (1440) is a length conversion; 0.5"/0.75pt are Word's defaults.
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

/** Encode a line-spacing multiple (1.0/1.15/1.5/2.0) as OOXML w:spacing `line`.
 *  Per ECMA-376, `lineRule="auto"` expresses `line` in 240ths of a single line
 *  (240 = 1.0, 360 = 1.5). The renderer inverts this — utils.lineSpacingToCss
 *  divides `line/240` back to a multiple for `calc(pitch × multiple)`. */
function lineMultipleToOoxml(mult: number): number {
  return Math.round(mult * 240);
}

/** The current selection's block node, but only if it carries the office-open
 *  paragraph attrs (paragraph or heading); null otherwise (e.g. inside a
 *  list item or table cell the block differs). */
function formattableBlock(editor: Editor): { type: string; attrs: Record<string, unknown> } | null {
  const { parent } = editor.state.selection.$from;
  return parent.type.name === "paragraph" || parent.type.name === "heading"
    ? { type: parent.type.name, attrs: (parent.attrs ?? {}) as Record<string, unknown> }
    : null;
}

/** Increase/decrease the current block's left indent by `delta` twips. */
function adjustIndent(editor: Editor, delta: number): void {
  const block = formattableBlock(editor);
  if (!block) return;
  const current = (block.attrs.indent ?? {}) as { left?: number; right?: number };
  const left = Math.max(0, (current.left ?? 0) + delta);
  editor
    .chain()
    .focus()
    .updateAttributes(block.type, { indent: { ...current, left } })
    .run();
}

/** Set the current block's line spacing to a multiple of single (1.0/1.15/…).
 *  Preserves existing before/after. "add-before"/"add-after" are not handled. */
function applyLineSpacing(editor: Editor, value?: string): void {
  const mult = parseFloat(value ?? "");
  if (!Number.isFinite(mult)) return;
  const block = formattableBlock(editor);
  if (!block) return;
  const current = (block.attrs.spacing ?? {}) as Record<string, unknown>;
  editor
    .chain()
    .focus()
    .updateAttributes(block.type, {
      spacing: { ...current, line: lineMultipleToOoxml(mult), lineRule: "auto" },
    })
    .run();
}

/** A theme-semantic color pick (from the palette's theme swatches): themeColor
 *  is the OOXML schemeClr name, val the resolved RGB (for rendering + w:val
 *  fallback), themeTint/themeShade the OOXML tint/shade hex. */
interface ThemeColorValue {
  themeColor: string;
  val: string;
  themeTint?: string;
  themeShade?: string;
}

function isThemeColor(value: unknown): value is ThemeColorValue {
  return typeof value === "object" && value !== null && "themeColor" in value && "val" in value;
}

/** Set/clear the current block's shading. "none" clears; a theme pick stores a
 *  themeFill-bound ShadingAttributesProperties (Word keeps it theme-bound); a
 *  bare hex stores fill directly. Theme semantics + verbatim round-trip keep
 *  the color faithful to MS Office. */
function applyShading(editor: Editor, value?: unknown): void {
  const block = formattableBlock(editor);
  if (!block) return;
  if (value === "none") {
    editor.chain().focus().updateAttributes(block.type, { shading: null }).run();
    return;
  }
  if (isThemeColor(value)) {
    const shading: Record<string, unknown> = {
      fill: value.val,
      type: "clear",
      themeFill: value.themeColor,
    };
    if (value.themeTint) shading.themeFillTint = value.themeTint;
    if (value.themeShade) shading.themeFillShade = value.themeShade;
    editor.chain().focus().updateAttributes(block.type, { shading }).run();
    return;
  }
  if (typeof value === "string" && value) {
    editor
      .chain()
      .focus()
      .updateAttributes(block.type, { shading: { fill: value, type: "clear" } })
      .run();
  }
}

/** Set/clear the run font color. "none" clears the color attr; a theme pick
 *  stores a ColorOptions object (themeColor/tint/shade + val) so the DOCX stays
 *  theme-bound; a bare hex stores the color directly. */
function applyFontColor(editor: Editor, value?: unknown): void {
  if (value === "none") {
    editor.chain().focus().setMark("textStyle", { color: null }).run();
    return;
  }
  if (isThemeColor(value) || (typeof value === "string" && value)) {
    editor.chain().focus().setMark("textStyle", { color: value }).run();
  }
}

/** Apply/clear paragraph borders. value picks sides (bottom/top/left/right/all/
 *  outside); "none" clears all. Merges with existing borders so other sides
 *  stay. Default style single 0.75pt, "auto" color (Word default). */
function applyBorder(editor: Editor, value?: string): void {
  const block = formattableBlock(editor);
  if (!block) return;
  if (value === "none") {
    editor.chain().focus().updateAttributes(block.type, { border: null }).run();
    return;
  }
  const sides =
    value === "all" || value === "outside"
      ? BORDER_SIDES
      : value && (BORDER_SIDES as readonly string[]).includes(value)
        ? [value]
        : null;
  if (!sides) return;
  const current = (block.attrs.border ?? {}) as Record<string, unknown>;
  const border = { ...current };
  for (const side of sides) border[side] = { ...DEFAULT_BORDER };
  editor.chain().focus().updateAttributes(block.type, { border }).run();
}

/**
 * Run the Tiptap command mapped to a ribbon `event` name.
 *
 * @returns `true` if a command ran; `false` when no command is mapped. Call
 *   sites ignore unknown events so not every ribbon button needs an entry.
 */
export function dispatchRibbonCommand(editor: Editor, event: string, value?: string): boolean {
  const run = RIBBON_COMMAND_MAP[event];
  if (!run) return false;
  run(editor, value);
  return true;
}
