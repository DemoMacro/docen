import type { ImageAttrs } from "@docen/docx";
import { BULLET_GLYPHS, nextOrderedReference, ORDERED_FORMATS } from "@docen/docx";
import { Extension } from "@docen/docx/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import type { EditorState } from "@tiptap/pm/state";
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
      // Increase/decrease left indent by Word's 0.5" step.
      "indent-increase":
        () =>
        ({ state, chain }) => {
          const block = formattableBlock(state);
          if (!block) return false;
          const current = (block.attrs.indent ?? {}) as { left?: number; right?: number };
          const left = Math.max(0, (current.left ?? 0) + INDENT_STEP_TWIPS);
          return chain()
            .updateAttributes(block.type, { indent: { ...current, left } })
            .run();
        },
      "indent-decrease":
        () =>
        ({ state, chain }) => {
          const block = formattableBlock(state);
          if (!block) return false;
          const current = (block.attrs.indent ?? {}) as { left?: number; right?: number };
          const left = Math.max(0, (current.left ?? 0) - INDENT_STEP_TWIPS);
          return chain()
            .updateAttributes(block.type, { indent: { ...current, left } })
            .run();
        },
      // Line spacing as a multiple of single (1.0/1.15/1.5/2.0); preserves
      // existing before/after. The dropdown's trailing entries are Word's
      // "Add Space Before/After Paragraph": 10pt (200 twips), not a multiple.
      "line-spacing":
        (mult) =>
        ({ state, chain }) => {
          const block = formattableBlock(state);
          if (!block) return false;
          const current = (block.attrs.spacing ?? {}) as Record<string, unknown>;
          if (mult === "add-before" || mult === "add-after") {
            const key = mult === "add-before" ? "before" : "after";
            return chain()
              .updateAttributes(block.type, { spacing: { ...current, [key]: 200 } })
              .run();
          }
          const m = parseFloat(mult ?? "");
          if (!Number.isFinite(m)) return false;
          return chain()
            .updateAttributes(block.type, {
              spacing: { ...current, line: lineMultipleToOoxml(m), lineRule: "auto" },
            })
            .run();
        },
      // Paragraph shading: "none" clears; a theme pick stores a themeFill-bound
      // ShadingProperties; a bare hex stores fill directly.
      shading:
        (value) =>
        ({ state, chain }) => {
          const block = formattableBlock(state);
          if (!block) return false;
          if (value === "none") {
            return chain().updateAttributes(block.type, { shading: null }).run();
          }
          if (isThemeColor(value)) {
            const shading: Record<string, unknown> = {
              fill: value.val,
              type: "clear",
              themeFill: value.themeColor,
            };
            if (value.themeTint) shading.themeFillTint = value.themeTint;
            if (value.themeShade) shading.themeFillShade = value.themeShade;
            return chain().updateAttributes(block.type, { shading }).run();
          }
          if (typeof value === "string" && value) {
            return chain()
              .updateAttributes(block.type, { shading: { fill: value, type: "clear" } })
              .run();
          }
          return false;
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
      // "none" clears all. Merges with existing so other sides stay. Default
      // single 0.75pt, "auto" color (Word default).
      border:
        (side) =>
        ({ state, chain }) => {
          const block = formattableBlock(state);
          if (!block) return false;
          // The split button's main click carries no value — default bottom.
          const s = side ?? "bottom";
          if (s === "none") {
            return chain().updateAttributes(block.type, { border: null }).run();
          }
          const sides =
            s === "all" || s === "outside"
              ? BORDER_SIDES
              : (BORDER_SIDES as readonly string[]).includes(s)
                ? [s]
                : null;
          if (!sides) return false;
          const current = (block.attrs.border ?? {}) as Record<string, unknown>;
          const border = { ...current };
          for (const side of sides) border[side] = { ...DEFAULT_BORDER };
          return chain().updateAttributes(block.type, { border }).run();
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
      // node type exists; every cell is a plain tableCell.
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
          const node = table.createAndFill(null, [mkRow(true), mkRow(false), mkRow(false)]);
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
      // paragraph's list kind and reference. No-op on non-list paragraphs.
      "multilevel-list":
        (level) =>
        ({ state, tr }) => {
          const target = level === "level-1" ? 0 : level === "level-3" ? 2 : 1;
          let touched = false;
          for (const { pos, node } of selectedParagraphs(state)) {
            const attrs = node.attrs as Record<string, unknown>;
            const cur = listStateOf(attrs);
            if (!cur.kind) continue;
            const patch =
              cur.kind === "bullet" && cur.variant === "bullet"
                ? { bullet: { level: target }, numbering: null }
                : { bullet: null, numbering: { reference: cur.reference, level: target } };
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
