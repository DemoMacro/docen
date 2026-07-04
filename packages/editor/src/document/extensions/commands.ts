import { Extension } from "@docen/docx/core";
import type { EditorState } from "@tiptap/pm/state";

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
      highlight: () => ReturnType;
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
      "bullet-list": () => ReturnType;
      "ordered-list": () => ReturnType;
      "task-list": () => ReturnType;
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
  "task-list",
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
]);

// ── Pure helpers (take EditorState, return data; never touch the chain) ──

/** HeadingLevel literals office-open lifts into `paragraph.heading` (pStyle
 *  val → Tiptap level). Title shares level 1 with Heading1. */
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

/** Encode a line-spacing multiple (1.0/1.15/1.5/2.0) as OOXML w:spacing `line`.
 *  Per ECMA-376, `lineRule="auto"` expresses `line` in 240ths of a single line
 *  (240 = 1.0, 360 = 1.5). The renderer inverts this — utils.lineSpacingToCss
 *  divides `line/240` back to a multiple for `calc(pitch × multiple)`. */
function lineMultipleToOoxml(mult: number): number {
  return Math.round(mult * 240);
}

/** The current selection's block node, but only if it carries the office-open
 *  paragraph attrs (paragraph or heading); null otherwise (e.g. inside a list
 *  item or table cell the block differs). */
function formattableBlock(
  state: EditorState,
): { type: string; attrs: Record<string, unknown> } | null {
  const { parent } = state.selection.$from;
  return parent.type.name === "paragraph" || parent.type.name === "heading"
    ? { type: parent.type.name, attrs: (parent.attrs ?? {}) as Record<string, unknown> }
    : null;
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
          commands.toggleBold(),
      italic:
        () =>
        ({ commands }) =>
          commands.toggleItalic(),
      underline:
        () =>
        ({ commands }) =>
          commands.toggleUnderline(),
      strike:
        () =>
        ({ commands }) =>
          commands.toggleStrike(),
      subscript:
        () =>
        ({ commands }) =>
          commands.toggleSubscript(),
      superscript:
        () =>
        ({ commands }) =>
          commands.toggleSuperscript(),
      highlight:
        () =>
        ({ commands }) =>
          commands.toggleHighlight(),
      code:
        () =>
        ({ commands }) =>
          commands.toggleCode(),
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
          commands.setTextAlign("left"),
      "align-center":
        () =>
        ({ commands }) =>
          commands.setTextAlign("center"),
      "align-right":
        () =>
        ({ commands }) =>
          commands.setTextAlign("right"),
      justify:
        () =>
        ({ commands }) =>
          commands.setTextAlign("justify"),

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
      // existing before/after.
      "line-spacing":
        (mult) =>
        ({ state, chain }) => {
          const m = parseFloat(mult ?? "");
          if (!Number.isFinite(m)) return false;
          const block = formattableBlock(state);
          if (!block) return false;
          const current = (block.attrs.spacing ?? {}) as Record<string, unknown>;
          return chain()
            .updateAttributes(block.type, {
              spacing: { ...current, line: lineMultipleToOoxml(m), lineRule: "auto" },
            })
            .run();
        },
      // Paragraph shading: "none" clears; a theme pick stores a themeFill-bound
      // ShadingAttributesProperties; a bare hex stores fill directly.
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
      "bullet-list":
        () =>
        ({ commands }) =>
          commands.toggleBulletList(),
      "ordered-list":
        () =>
        ({ commands }) =>
          commands.toggleOrderedList(),
      "task-list":
        () =>
        ({ commands }) =>
          commands.toggleTaskList(),
      blockquote:
        () =>
        ({ commands }) =>
          commands.toggleBlockquote(),
      "horizontal-rule":
        () =>
        ({ commands }) =>
          commands.setHorizontalRule(),
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
      // Insert a 3×3 table (Word's default Insert > Table preset).
      "insert-table":
        () =>
        ({ commands }) =>
          commands.insertTable({ rows: 3, cols: 3, withHeaderRow: true }),
      // Wrap the selection in a link (empty selection → link around the URL text).
      link:
        (href) =>
        ({ chain }) => {
          const url = href || (typeof window !== "undefined" && window.prompt("Link URL")) || "";
          if (!url) return false;
          return chain().extendMarkRange("link").setLink({ href: url }).run();
        },

      // ── Style gallery (combobox-driven): value picks the block style ──
      // A HeadingLevel id switches the block to a heading and stamps styleId;
      // everything else becomes a paragraph carrying styleId so the injected
      // document CSS applies. setNode bypasses setHeading's options.levels gate
      // — Tiptap's Level type caps at 1-6, so levels 7-9 are valid on the schema
      // attr but setHeading would reject them.
      style:
        (styleId) =>
        ({ chain }) => {
          const id = (styleId ?? "").trim();
          if (!id || id === "Normal") {
            return chain()
              .setParagraph()
              .updateAttributes("paragraph", { styleId: id || null })
              .run();
          }
          const level = HEADING_LEVEL_BY_STYLE[id];
          if (level) {
            return chain()
              .setNode("heading", { level })
              .updateAttributes("heading", { styleId: id })
              .run();
          }
          return chain().setParagraph().updateAttributes("paragraph", { styleId: id }).run();
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
      // Promote/demote the current list item toward a multilevel depth
      // (level-1 = top, level-2/3 = sink once/twice). No-op outside a list.
      "multilevel-list":
        (level) =>
        ({ chain }) => {
          if (level === "level-1") return chain().liftListItem("listItem").run();
          const c = chain();
          const times = level === "level-3" ? 2 : 1;
          for (let i = 0; i < times; i++) c.sinkListItem("listItem");
          return c.run();
        },
    };
  },
});
