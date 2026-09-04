// Inline marks with DOCX run-property hooks (renderDocx/parseDocx) so
// DocxManager resolves and compiles them via reflection. Each renderDocx
// contributes rPr fields for the run; parseDocx returns null when the run does
// not carry the mark, so DocxManager skips emitting it.
//
// Fully custom Mark.create definitions (no upstream mark extensions): a mark
// here is just a name + HTML tag pair plus the OOXML hook — the upstream
// packages wrapped exactly this in input rules and toggle commands we never
// wired (the canvas route types through the bridge and toggles via the core
// `toggleMark` command).

import type { RunOptions } from "@office-open/docx";
import { HighlightColor, HIGHLIGHT_PALETTE_RGB } from "@office-open/docx";
import { Mark } from "@tiptap/core";

import { normalizeColorToHex } from "./utils";

// Each mark surfaces only the "true" state for editor interaction (toolbar
// toggle, <strong>/<em>). The full three-state round-trip (true/false/null) is
// owned by TextStyle's matching attr, so an inherited bold/italic cancelled by
// <w:b w:val="0"/> / <w:i w:val="0"/> round-trips as TextStyle.bold=false (and
// renders font-weight:normal via CSS cascade) without a mark here.

export const Bold = Mark.create({
  name: "bold",
  parseHTML() {
    return [
      { tag: "strong" },
      { tag: "b" },
      // Word/web pastes carry emphasis as styled spans, not tags — match the
      // numeric scale too (600+ is semibold and up).
      {
        style: "font-weight",
        getAttrs: (v) => /^(bold(?:er)?|[5-9]\d{2,})$/s.test(String(v)) && null,
      },
    ];
  },
  renderDocx: () => ({ bold: true }),
  parseDocx: (opts: RunOptions) => (opts.bold ? {} : null),
});

export const Italic = Mark.create({
  name: "italic",
  parseHTML() {
    return [{ tag: "em" }, { tag: "i" }, { style: "font-style=italic" }];
  },
  renderDocx: () => ({ italic: true }),
  parseDocx: (opts: RunOptions) => (opts.italic ? {} : null),
});

export const Underline = Mark.create({
  name: "underline",
  addAttributes() {
    return {
      // w:u val (ST_Underline token) and w:u color (hex without #) — null
      // means Word's default single line in the text color.
      style: { default: null },
      color: { default: null },
    };
  },
  parseHTML() {
    return [
      { tag: "u" },
      // The shorthand is often compound ("underline line-through") — match on
      // the token, not the whole value.
      {
        style: "text-decoration",
        getAttrs: (v) => String(v).includes("underline") && null,
      },
    ];
  },
  renderDocx: (attrs: Record<string, unknown>) => {
    const style = typeof attrs.style === "string" && attrs.style ? attrs.style : "single";
    const color = typeof attrs.color === "string" && attrs.color ? attrs.color : undefined;
    return { underline: color ? { type: style, color } : { type: style } };
  },
  // office-open represents <w:u> as { type, color? }. val="none" means NO
  // underline — Word writes it to cancel an inherited underline (e.g. a run
  // inside a hyperlink style). `{ type: "none" }` is a truthy object, so a
  // plain `opts.underline ?` check turned val="none" into val="single" on
  // round-trip, silently adding underlines the source never had. Only treat
  // a concrete non-"none" type as an underline mark.
  parseDocx: (opts: RunOptions) => {
    const u = opts.underline as { type?: string; color?: string } | undefined;
    return u && u.type && u.type !== "none" ? { style: u.type, color: u.color ?? null } : null;
  },
});

export const Code = Mark.create({
  name: "code",
  parseHTML() {
    return [{ tag: "code" }];
  },
  // rStyle "CodeChar" is the precise round-trip carrier; Consolas is a visual
  // fallback when styles.xml lacks the CodeChar character-style definition.
  renderDocx: () => ({ style: "CodeChar", font: "Consolas" }),
  parseDocx: (opts: RunOptions) => (opts.style === "CodeChar" ? {} : null),
});

// w:highlight val is the fixed ST_HighlightColor enumeration — hex is illegal
// there and Word refuses the file. The palette table lives in office-open next
// to HighlightColor; off-palette colors ride as character shading (the same
// encoding Word itself uses — highlight supersedes shd when both are present).
const HIGHLIGHT_NAME_BY_RGB = new Map(
  Object.entries(HIGHLIGHT_PALETTE_RGB).map(([name, rgb]) => [rgb.toUpperCase(), name]),
);
const HIGHLIGHT_NAMES: ReadonlySet<string> = new Set(Object.values(HighlightColor));

export const Highlight = Mark.create({
  name: "highlight",
  addAttributes() {
    return {
      color: {
        default: null,
        // Word's highlight palette (16 fixed colors); HTML keeps it on
        // data-color so a re-parsed mark survives without CSS interpretation.
        parseHTML: (el: HTMLElement) =>
          el.getAttribute("data-color") || el.style.backgroundColor || null,
      },
    };
  },
  parseHTML() {
    return [{ tag: "mark" }];
  },
  renderDocx: (attrs: Record<string, unknown>) => {
    const color = typeof attrs.color === "string" ? attrs.color : null;
    if (!color) return { highlight: "yellow" };
    // Palette names are lowercase OOXML tokens; accept any case, emit the token.
    const lower = color.toLowerCase();
    if (HIGHLIGHT_NAMES.has(lower)) return { highlight: lower };
    // normalizeColorToHex accepts every form pasted HTML carries (rgb(), #hex,
    // CSS names); a palette hex maps back to its token. Off-palette colors ride
    // as character shading — the same encoding Word itself uses when a
    // highlight color outside the fixed palette is wanted (highlight supersedes
    // shd when both are present, so this reads back as plain background color).
    const hex = normalizeColorToHex(color)?.slice(1);
    if (hex) {
      const named = HIGHLIGHT_NAME_BY_RGB.get(hex);
      return named ? { highlight: named } : { shading: { fill: hex, type: "clear" } };
    }
    // Unparseable color — fall back to the default token rather than emit an
    // illegal w:highlight/w:fill value Word would refuse.
    return { highlight: "yellow" };
  },
  // office-open models <w:highlight w:val="none"/> (Word's "cancel inherited
  // highlight") as the string "none" — a truthy value. Exclude it so the mark
  // is not wrongly applied; TextStyle still passes highlight="none" through
  // verbatim (NATIVE_RUN_ATTRS), so the DOCX round-trips losslessly. Same class
  // as the underline val="none" fix.
  parseDocx: (opts: RunOptions) => {
    const h = opts.highlight;
    return h && h !== "none" ? { color: h as string } : null;
  },
});

export const Subscript = Mark.create({
  name: "subscript",
  // A run is either sub- or superscript — toggling one must clear the other.
  excludes: "superscript",
  parseHTML() {
    return [{ tag: "sub" }];
  },
  renderDocx: () => ({ verticalAlign: "subscript" }),
  parseDocx: (opts: RunOptions) => (opts.verticalAlign === "subscript" ? {} : null),
});

export const Superscript = Mark.create({
  name: "superscript",
  excludes: "subscript",
  parseHTML() {
    return [{ tag: "sup" }];
  },
  renderDocx: () => ({ verticalAlign: "superscript" }),
  parseDocx: (opts: RunOptions) => (opts.verticalAlign === "superscript" ? {} : null),
});

/**
 * Strike mark — editor interaction only (toolbar toggle, `<s>`).
 *
 * OOXML has two mutually exclusive strikethrough booleans: `strike` (single)
 * and `doubleStrike` (double). Both are three-state and ride TextStyle's native
 * attrs for round-trip + CSS cascade, so this mark only surfaces the
 * "single strike = true" case for editing. doubleStrike=true (no single-strike
 * mark) round-trips purely through TextStyle.
 */
export const Strike = Mark.create({
  name: "strike",
  parseHTML() {
    return [
      { tag: "s" },
      { tag: "del" },
      { tag: "strike" },
      {
        style: "text-decoration",
        getAttrs: (v) => String(v).includes("line-through") && null,
      },
    ];
  },
  renderDocx: () => ({ strike: true }),
  parseDocx: (opts: RunOptions): Record<string, unknown> | null => (opts.strike ? {} : null),
});
