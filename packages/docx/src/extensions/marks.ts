// Base inline marks extended with DOCX run-property hooks (renderDocx/parseDocx)
// so DocxManager resolves and compiles them via reflection. Each renderDocx
// contributes rPr fields for the run; parseDocx returns null when the run does
// not carry the mark, so DocxManager skips emitting it.

import type { RunOptions } from "@office-open/docx";
import { Bold as BaseBold } from "@tiptap/extension-bold";
import { Code as BaseCode } from "@tiptap/extension-code";
import { Highlight as BaseHighlight } from "@tiptap/extension-highlight";
import { Italic as BaseItalic } from "@tiptap/extension-italic";
import { Subscript as BaseSubscript } from "@tiptap/extension-subscript";
import { Superscript as BaseSuperscript } from "@tiptap/extension-superscript";
import { Underline as BaseUnderline } from "@tiptap/extension-underline";

export const Bold = BaseBold.extend({
  renderDocx: () => ({ bold: true }),
  parseDocx: (opts: RunOptions) => (opts.bold ? {} : null),
});

export const Italic = BaseItalic.extend({
  renderDocx: () => ({ italic: true }),
  parseDocx: (opts: RunOptions) => (opts.italic ? {} : null),
});

export const Underline = BaseUnderline.extend({
  renderDocx: () => ({ underline: { type: "single" } }),
  // office-open represents <w:u> as { type, color? }. val="none" means NO
  // underline — Word writes it to cancel an inherited underline (e.g. a run
  // inside a hyperlink style). `{ type: "none" }` is a truthy object, so a
  // plain `opts.underline ?` check turned val="none" into val="single" on
  // round-trip, silently adding underlines the source never had. Only treat
  // a concrete non-"none" type as an underline mark.
  parseDocx: (opts: RunOptions) => {
    const u = opts.underline as { type?: string } | undefined;
    return u && u.type && u.type !== "none" ? {} : null;
  },
});

export const Code = BaseCode.extend({
  // rStyle "CodeChar" is the precise round-trip carrier; Consolas is a visual
  // fallback when styles.xml lacks the CodeChar character-style definition.
  renderDocx: () => ({ style: "CodeChar", font: "Consolas" }),
  parseDocx: (opts: RunOptions) => (opts.style === "CodeChar" ? {} : null),
});

export const Highlight = BaseHighlight.extend({
  renderDocx: (attrs: Record<string, unknown>) => ({ highlight: attrs.color ?? "yellow" }),
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

export const Subscript = BaseSubscript.extend({
  renderDocx: () => ({ subScript: true }),
  parseDocx: (opts: RunOptions) => (opts.subScript ? {} : null),
});

export const Superscript = BaseSuperscript.extend({
  renderDocx: () => ({ superScript: true }),
  parseDocx: (opts: RunOptions) => (opts.superScript ? {} : null),
});
