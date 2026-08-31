// Inline run projection: text runs (rPr resolved over the paragraph
// default), hard breaks, tabs, pictures, footnote/endnote references, fields
// (PAGE/NUMPAGES dynamic, complex-field results re-hydrated), and the
// container children (hyperlink / tracked insertion / deletion).

import { emuToPx, ptToPx, twipToPx, type LayoutInline, type LayoutTextStyle } from "@docen/layout";

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
    return {
      family: toFamily(own.font, fontAttr(chainRPr.font) ?? fontAttr(docRPr.font)) ?? defRun.family,
      sizePx: ptToPx(own.sizePt ?? num(chainRPr.size) ?? num(docRPr.size) ?? 12),
      bold: own.bold ?? defRun.bold,
      italic: own.italic ?? defRun.italic,
      color: own.color ?? defRun.color,
      highlight: own.highlight ?? str(chainRPr.highlight) ?? str(docRPr.highlight),
      underline: own.underline ?? defRun.underline,
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
    out.push({ kind: "text", text, style: textStyleOf(rPr), commentIds });
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
    if (instr.startsWith("PAGE") && !instr.startsWith("PAGES")) {
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
        out.push({
          kind: "text",
          text: String(noteOrdinal(ctx.footnoteOrdinals, fnRefId)),
          style: { ...textStyleOf(rPr), verticalAlign: "superscript" },
        });
      }
      const enRefId = noteRefId(child, "endnoteReference");
      if (enRefId != null) {
        out.push({
          kind: "text",
          text: romanNumeral(noteOrdinal(ctx.endnoteOrdinals, enRefId), false),
          style: { ...textStyleOf(rPr), verticalAlign: "superscript" },
        });
      }
      if (typeof child.text === "string") pushText(child.text, rPr);
      if (child.break != null) out.push({ kind: "break" });
      if (child.tab != null) out.push({ kind: "tab" });
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
