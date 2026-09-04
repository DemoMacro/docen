// Inline run projection: text runs (rPr resolved over the paragraph
// default), hard breaks, tabs, pictures, footnote/endnote references, fields
// (PAGE/NUMPAGES dynamic, complex-field results re-hydrated), and the
// container children (hyperlink / tracked insertion / deletion).

import {
  emuToPx,
  ptToPx,
  twipToPx,
  type LayoutCombine,
  type LayoutInline,
  type LayoutTextStyle,
} from "@docen/layout";

import { mergeStyleChain } from "../../style-cascade";
import type { ProjectContext } from "./context";
import { cropOf } from "./drawing";
import { isRecord, measureEmu, num, str, unescapeXml, type Rec } from "./guards";
import { metafileMembers, pictureSrc } from "./media";
import { romanNumeral } from "./numbering";
import { fontAttr, toFamily, runStyleOf } from "./styles";

/** Word display presets merged UNDER a container's runs (explicit run props
 *  win per field): tracked insertions underline and tracked deletions strike
 *  in the first author's revision red — Word's "By author" palette starts at
 *  red, and a single default author sees red for every revision. Hyperlinks
 *  carry no preset: Word styles them only via the run's "Hyperlink" character
 *  style (w:rStyle), which the cascade resolves per run. */
const INSERTION_DISPLAY = { underline: { type: "single" }, color: "FF0000" } as const;
const DELETION_DISPLAY = { strike: true, color: "FF0000" } as const;

/** A footnote/endnote reference's note id — the bare number form (`{
 *  footnoteReference: 1 }` / `{ endnoteReference: 1 }`) or the option object
 *  form (`{ id }`); anything else is not one. */
function noteRefId(child: Rec, key: "footnoteReference" | "endnoteReference"): number | undefined {
  const ref = child[key];
  if (typeof ref === "number") return ref;
  if (isRecord(ref)) return num(ref.id);
  return undefined;
}

/** The displayed ordinal for a note id — assign the next number on first
 *  reference, reuse it afterward (the Nth distinct note referenced shows N). */
function noteOrdinal(ordinals: Map<number, number>, id: number): number {
  let ordinal = ordinals.get(id);
  if (ordinal == null) {
    ordinal = ordinals.size + 1;
    ordinals.set(id, ordinal);
  }
  return ordinal;
}

/** A short structural label for a formula placeholder box — the outermost
 *  structure's glyph approximation (√□, □/□, ∑□), not a rendering of it. */
function mathLabelOf(math: Rec): string {
  const first = Array.isArray(math.children)
    ? math.children.find((c) => isRecord(c) && !("text" in c))
    : undefined;
  if (!isRecord(first)) return "fx";
  if ("fraction" in first) return "□/□";
  if ("superScript" in first) return "□^□";
  if ("subScript" in first) return "□_□";
  if ("radical" in first) return "√□";
  if ("sum" in first) return "∑□";
  if ("integral" in first) return "∫□";
  return "fx";
}

const COMBINE_BRACKETS: Record<string, "round" | "square" | "angle" | "curly"> = {
  round: "round",
  square: "square",
  angle: "angle",
  curly: "curly",
};

/** Two-lines-in-one metadata (w:eastAsianLayout @w:combine) for a run's text:
 *  the two lines the run packs into and its optional bracket pair. Word's
 *  dialog lets spaces mark the split — folded away here for an even split of
 *  the remaining characters. */
function combineOf(rPr: Rec, text: string): LayoutCombine | undefined {
  const layout = isRecord(rPr.eastAsianLayout) ? rPr.eastAsianLayout : {};
  const on = layout.combine === true || layout.combine === "1" || layout.combine === 1;
  if (!on) return undefined;
  const packed = text.replace(/\s+/g, "");
  if (!packed) return undefined;
  const half = Math.ceil(packed.length / 2);
  const bracket = str(layout.combineBrackets);
  return {
    first: packed.slice(0, half),
    second: packed.slice(half),
    ...(bracket && COMBINE_BRACKETS[bracket] ? { bracket: COMBINE_BRACKETS[bracket] } : {}),
  };
}

/** Inline content: text runs (rPr resolved over the paragraph default), hard
 *  breaks, pictures (paragraph-child or run-child slot), and the container
 *  children (hyperlink / insertion / deletion — their runs project with the
 *  Word display preset above) as atoms. The members arrive as unknown — the
 *  ParagraphChild union is wide and its runtime shapes are looser still
 *  (compile pushes `{text, …rPr}` run forms), so each leg is validated rather
 *  than trusted.
 *  Known-but-unprojected inline atoms (tab, chart, math, fields) carry no box
 *  yet — they render as absence, a registered gap to close type by type. */
export function projectRuns(
  runs: readonly unknown[],
  chainRPr: Rec,
  docRPr: Rec,
  defRun: LayoutTextStyle,
  ctx: ProjectContext,
): LayoutInline[] {
  const { openComments } = ctx;
  const out: LayoutInline[] = [];
  const textStyleOf = (rPr: Rec): LayoutTextStyle => {
    // A run's character style (w:rStyle, e.g. a body link's "Hyperlink") slots
    // between the paragraph-style chain and direct formatting: its props beat
    // chainRPr/docRPr (they land in `own`), and an explicit rPr field beats it
    // (the spread keeps rPr's keys on top). Word paints w:hyperlink content
    // solely through this style — the container itself carries no look.
    const styleId = str(rPr.style);
    const charRun = styleId ? mergeStyleChain(ctx.characterStyles, styleId).run : undefined;
    const own = runStyleOf(charRun ? { ...charRun, ...rPr } : rPr);
    const isNoteRun = ctx.currentNoteOrdinal != null;
    const chainSizeRaw = num(chainRPr.size);
    const chainSize = isNoteRun && chainSizeRaw === 20 ? 10 : chainSizeRaw;
    const effectiveSizePt = own.sizePt ?? chainSize ?? num(docRPr.size) ?? (isNoteRun ? 10 : 12);
    return {
      family: toFamily(own.font, fontAttr(chainRPr.font) ?? fontAttr(docRPr.font)) ?? defRun.family,
      sizePx: ptToPx(effectiveSizePt),
      bold: own.bold ?? defRun.bold,
      italic: own.italic ?? defRun.italic,
      color: own.color ?? defRun.color,
      highlight: own.highlight ?? str(chainRPr.highlight) ?? str(docRPr.highlight),
      shadingFill:
        own.shadingFill ??
        str((isRecord(chainRPr.shading) ? chainRPr.shading : undefined)?.fill) ??
        str((isRecord(docRPr.shading) ? docRPr.shading : undefined)?.fill) ??
        defRun.shadingFill,
      underline: own.underline ?? defRun.underline,
      underlineStyle: own.underlineStyle ?? defRun.underlineStyle,
      underlineColor: own.underlineColor ?? defRun.underlineColor,
      strikethrough: own.strikethrough ?? defRun.strikethrough,
      letterSpacingPx:
        own.characterSpacingTw != null ? twipToPx(own.characterSpacingTw) : defRun.letterSpacingPx,
      verticalAlign: own.verticalAlign ?? defRun.verticalAlign,
    };
  };
  const pushText = (text: string, rPr: Rec): void => {
    if (!text) return;
    const commentIds =
      openComments && openComments.size > 0 ? [...openComments].sort((a, b) => a - b) : undefined;
    // Two-lines-in-one (双行合一 / 合并字符): the run packs into two
    // half-size lines; the dialog's spaces mark the split in Word, here
    // folded away for an even split.
    const combine = combineOf(rPr, text);
    out.push({
      kind: "text",
      text,
      style: textStyleOf(rPr),
      commentIds,
      ...(combine ? { combine } : {}),
    });
  };
  /** A field (w:fldSimple / complexField): PAGE/NUMPAGES become dynamic atoms
   *  (the painter resolves the number per page — `text` is a measuring
   *  placeholder); anything else renders its cached result. A structured
   *  result (resultRunsXml — present when the result runs hold anything but
   *  plain text, e.g. a TOC hyperlink's tab + nested PAGEREF) is re-hydrated
   *  item by item; only then does the flat `result` string stand in. */
  const pushField = (field: Rec, rPr: Rec): void => {
    const instr =
      typeof field.instruction === "string" ? field.instruction.trim().toUpperCase() : "";
    const cached =
      typeof field.result === "string" ? field.result : (field.cachedValue as string | undefined);
    const style = textStyleOf(rPr);
    if (instr.startsWith("PAGE") && !instr.startsWith("PAGES") && !instr.startsWith("PAGEREF")) {
      out.push({ kind: "text", text: "0", style, field: "page" });
    } else if (instr.startsWith("NUMPAGES")) {
      out.push({ kind: "text", text: "0", style, field: "numPages" });
    } else if (typeof field.resultRunsXml === "string") {
      pushFieldResultRuns(field.resultRunsXml, rPr);
    } else if (cached) {
      pushText(cached, rPr);
    }
  };
  /** Walk a complex field's verbatim result-run XML: text runs become text,
   *  w:tab atoms become tab jumps, and a nested field (a TOC entry's PAGEREF)
   *  contributes its separated result — instruction runs are skipped. */
  const pushFieldResultRuns = (xml: string, rPr: Rec): void => {
    // Stack of nested fields, each "instr" until its separate, then "result".
    const stack: ("instr" | "result")[] = [];
    const tokens =
      xml.match(/<w:(fldChar|instrText|tab|t)\b[^>]*(?:\/>|>([\s\S]*?)<\/w:\1>)/g) ?? [];
    for (const tk of tokens) {
      if (tk.startsWith("<w:fldChar")) {
        const type = /w:fldCharType="(\w+)"/.exec(tk)?.[1];
        if (type === "begin") stack.push("instr");
        else if (type === "separate" && stack.length > 0) stack[stack.length - 1] = "result";
        else if (type === "end") stack.pop();
      } else if (tk.startsWith("<w:instrText")) {
        continue;
      } else if (stack[stack.length - 1] === "instr") {
        continue;
      } else if (tk.startsWith("<w:tab")) {
        out.push({ kind: "tab" });
      } else {
        const text = unescapeXml(/>([\s\S]*?)<\/w:t>$/.exec(tk)?.[1] ?? "");
        pushText(text, rPr);
      }
    }
  };
  /** Flatten a w:rt / w:rubyBase content's runs to plain text — the
   *  annotation has no formatting of its own beyond its font size. */
  const rubyContentText = (content: unknown): string => {
    if (!isRecord(content) || !Array.isArray(content.children)) return "";
    let text = "";
    for (const c of content.children) {
      if (typeof c === "string") text += c;
      else if (isRecord(c) && typeof c.text === "string") text += c.text;
    }
    return text;
  };
  const pushRuby = (ruby: Rec, rPr: Rec): void => {
    const baseText = rubyContentText(ruby.base);
    if (!baseText) return;
    const props = isRecord(ruby.properties) ? ruby.properties : {};
    const style = textStyleOf(rPr);
    // Word's default annotation is half the base size (w:hps absent).
    const sizePt = num(props.fontSize);
    out.push({
      kind: "text",
      text: baseText,
      style,
      ruby: {
        text: rubyContentText(ruby.text),
        alignment: str(props.alignment),
        fontSizePx: sizePt != null ? ptToPx(sizePt) : style.sizePx / 2,
      },
    });
  };
  const pushPicture = (pic: Rec): void => {
    // A floating picture is an anchored drawing (projectDrawings), not an
    // inline atom — projecting it here too would double-render it.
    if (isRecord(pic.floating)) return;
    const tr = isRecord(pic.transformation) ? pic.transformation : {};
    const w = measureEmu(tr.width);
    const h = measureEmu(tr.height);
    if (w != null && h != null) {
      const widthPx = emuToPx(w);
      const heightPx = emuToPx(h);
      // The metafile replay (WMF vector layers) is the main battlefield for
      // inline pictures — the flat DIB src only fills in when replay fails.
      // A flat src carries its a:srcRect crop; a replay folds the same crop
      // into its frame mapping — dropping it stretches the WHOLE source into
      // the extent box.
      const members = metafileMembers(pic, widthPx, heightPx, cropOf(pic));
      out.push({
        kind: "picture",
        widthPx,
        heightPx,
        src: members ? undefined : pictureSrc(pic),
        crop: members ? undefined : cropOf(pic),
        members,
      });
    }
  };
  /** One nesting level's walk. `preset` carries the enclosing containers'
   *  display fields (outermost first, an inner leg overrides) merged under
   *  each run's own rPr — explicit run props always win per field. */
  const pushRuns = (items: readonly unknown[], preset: Rec): void => {
    for (const child of items) {
      if (typeof child === "string") {
        pushText(child, preset);
        continue;
      }
      if (!isRecord(child)) continue;
      // Comment range markers are zero-width: a start opens tinting for every
      // text atom after it, an end closes it. The set lives across paragraphs
      // (the caller's walk), matching Word's range semantics.
      if (isRecord(child.commentRangeStart) && num(child.commentRangeStart.id) != null)
        openComments?.add(num(child.commentRangeStart.id)!);
      if (isRecord(child.commentRangeEnd) && num(child.commentRangeEnd.id) != null)
        openComments?.delete(num(child.commentRangeEnd.id)!);
      const rPr: Rec = { ...preset, ...child };
      // A footnote/endnote reference is a superscript ordinal (Word's
      // FootnoteReference/EndnoteReference style look) — numbered by
      // first-reference order, the same id twice showing the same number;
      // endnotes paint lowercase Roman (Word's endnote default numFmt). The
      // reference run's own rPr still applies.
      const fnRefId = noteRefId(child, "footnoteReference");
      if (fnRefId != null) {
        const ordinal = noteOrdinal(ctx.footnoteOrdinals, fnRefId);
        out.push({
          kind: "text",
          text: String(ordinal),
          style: { ...textStyleOf(rPr), verticalAlign: "superscript" },
          noteRef: { kind: "footnote", id: fnRefId, ordinal },
        });
      }
      const enRefId = noteRefId(child, "endnoteReference");
      if (enRefId != null) {
        const ordinal = noteOrdinal(ctx.endnoteOrdinals, enRefId);
        out.push({
          kind: "text",
          text: romanNumeral(ordinal, false),
          style: { ...textStyleOf(rPr), verticalAlign: "superscript" },
          noteRef: { kind: "endnote", id: enRefId, ordinal },
        });
      }
      if (child.footnoteRef === true) {
        out.push({
          kind: "text",
          text: String(ctx.currentNoteOrdinal ?? 1),
          style: { ...textStyleOf(rPr), verticalAlign: "superscript" },
        });
      }
      if (child.endnoteRef === true) {
        out.push({
          kind: "text",
          text: romanNumeral(ctx.currentNoteOrdinal ?? 1, false),
          style: { ...textStyleOf(rPr), verticalAlign: "superscript" },
        });
      }
      if (typeof child.text === "string") pushText(child.text, rPr);
      // A phonetic guide (w:ruby — Word's 拼音指南): the base text stays the
      // atom's own text (the guide rides as paint metadata); consumed here so
      // the children walk below does not re-emit the base runs verbatim.
      if (isRecord(child.ruby)) pushRuby(child.ruby, rPr);
      if (child.break != null) out.push({ kind: "break" });
      if (child.tab != null) out.push({ kind: "tab" });
      if (isRecord(child.math)) {
        // An OMML formula the engine does not lay out yet: a fixed
        // placeholder slot with a structural label, styled as an inert
        // annotation (Word shows empty argument slots until the math
        // layout engine lands).
        const style = { ...textStyleOf(rPr), italic: true, color: "#808080" };
        const label = mathLabelOf(child.math);
        out.push({
          kind: "math",
          label,
          widthPx: label.length * style.sizePx * 0.7 + 8,
          heightPx: style.sizePx * 1.6,
        });
      }
      if (isRecord(child.picture)) pushPicture(child.picture);
      if (isRecord(child.complexField)) pushField(child.complexField, rPr);
      if (isRecord(child.simpleField)) pushField(child.simpleField, rPr);
      if (isRecord(child.hyperlink) && Array.isArray(child.hyperlink.children)) {
        pushRuns(child.hyperlink.children, preset);
      }
      if (isRecord(child.insertion) && Array.isArray(child.insertion.children)) {
        pushRuns(child.insertion.children, { ...preset, ...INSERTION_DISPLAY });
      }
      if (isRecord(child.deletion) && Array.isArray(child.deletion.children)) {
        pushRuns(child.deletion.children, { ...preset, ...DELETION_DISPLAY });
      }
      if (Array.isArray(child.children)) pushRuns(child.children, preset);
    }
  };
  pushRuns(runs, {});
  return out;
}
