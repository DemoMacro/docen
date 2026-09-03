// Paragraph projection: one LayoutParagraph out of the full style cascade
// (direct → style chain → docDefaults) — spacing, indent, tab stops,
// borders, shading, the numbering marker, and the inline run walk.

import {
  ptToPx,
  twipToPx,
  type LayoutInline,
  type LayoutLineHeight,
  type LayoutParagraph,
  type LayoutParagraphBorderEdge,
  type LayoutSpacing,
  type LayoutTabStop,
  type LayoutTextStyle,
} from "@docen/layout";

import type { ProjectContext } from "./context";
import { projectDrawings } from "./drawing";
import { isRecord, measureTwip, num, str, type BodyParagraph, type Rec } from "./guards";
import { BUILTIN_BULLET_LEVEL, formatListNumber } from "./numbering";
import { projectRuns } from "./runs";
import {
  alignOf,
  docDefaultsOf,
  fontAttr,
  pick,
  runStyleOf,
  styleChainOf,
  toFamily,
  type FontAttr,
} from "./styles";

// ── paragraph projection ──

function toLineHeight(line: number | undefined, rule: unknown): LayoutLineHeight | undefined {
  if (line == null) return undefined;
  if (rule === "exact") return { rule: "exact", px: twipToPx(line) };
  if (rule === "atLeast") return { rule: "atLeast", px: twipToPx(line) };
  return { rule: "multiple", factor: line / 240 };
}

export function projectParagraph(p: BodyParagraph, ctx: ProjectContext): LayoutParagraph {
  const pPr: Rec = isRecord(p) ? p : {};
  const styleId = str(pPr.style) ?? str(pPr.heading);
  const chain = styleChainOf(ctx.styles, styleId);
  const chainPPr: Rec = chain.paragraph;
  const chainRPr: Rec = chain.run;
  const docDefaults = docDefaultsOf(ctx.styles);
  const docPPr: Rec = isRecord(docDefaults.paragraph) ? docDefaults.paragraph : {};
  const docRPr: Rec = isRecord(docDefaults.run) ? docDefaults.run : {};

  // ¶-mark strut: direct rPr, else style chain run over docDefaults.
  const markRun: Rec = isRecord(pPr.run) ? pPr.run : {};
  const markSize = num(markRun.size);
  const markSizePt = markSize ?? num(chainRPr.size) ?? num(docRPr.size) ?? 12;
  const defFont: FontAttr = fontAttr(chainRPr.font) ?? fontAttr(docRPr.font) ?? null;
  // The paragraph's default run style — every field cascades from the style
  // chain over docDefaults (Word's effective-rPr resolution), so a style's
  // color/underline/strikethrough reach runs that carry no rPr of their own,
  // exactly as bold/italic already did.
  const chainDefRun = runStyleOf({ ...docRPr, ...chainRPr });
  const defaultTextStyle: LayoutTextStyle = {
    family: toFamily(null, defFont) ?? {},
    sizePx: ptToPx(markSizePt),
    bold: chainDefRun.bold,
    italic: chainDefRun.italic,
    color: chainDefRun.color,
    underline: chainDefRun.underline,
    strikethrough: chainDefRun.strikethrough,
    letterSpacingPx:
      chainDefRun.characterSpacingTw != null ? twipToPx(chainDefRun.characterSpacingTw) : undefined,
  };

  // Spacing/indent cascade: direct attr wins per-field, else chain, else docDefaults.
  const direct: Rec = isRecord(pPr.spacing) ? pPr.spacing : {};
  const styleSp: Rec = isRecord(chainPPr.spacing) ? chainPPr.spacing : {};
  const docSp: Rec = isRecord(docPPr.spacing) ? docPPr.spacing : {};
  const spacing: LayoutSpacing = {
    beforePx: twipToPx(measureTwip(pick([direct, styleSp, docSp], "before")) ?? 0),
    afterPx: twipToPx(measureTwip(pick([direct, styleSp, docSp], "after")) ?? 0),
    lineHeight: toLineHeight(
      measureTwip(pick([direct, styleSp, docSp], "line")),
      pick([direct, styleSp, docSp], "lineRule"),
    ),
  };

  const dInd: Rec = isRecord(pPr.indent) ? pPr.indent : {};
  const sInd: Rec = isRecord(chainPPr.indent) ? chainPPr.indent : {};
  const docInd: Rec = isRecord(docPPr.indent) ? docPPr.indent : {};
  const ind = (key: string): unknown => pick([dInd, sInd, docInd], key);

  // Numbering: the paragraph's own numPr wins, else the style chain's. The
  // level's indent fills gaps the paragraph left unset (Word: direct w:ind
  // overrides w:lvl's); the level's marker (bullet glyph or its live counter)
  // prepends + a tab hop to the body-text start. A `bullet {level}` paragraph
  // (no numbering definition) resolves against the built-in bullet table.
  const numRef: Rec | null = isRecord(pPr.numbering)
    ? pPr.numbering
    : isRecord(chainPPr.numbering)
      ? chainPPr.numbering
      : null;
  const numReference = numRef ? str(numRef.reference) : undefined;
  const numLevelIndex = num(numRef?.level) ?? 0;
  const levels = numReference ? ctx.numberings.get(numReference) : undefined;
  const bulletLevel = isRecord(pPr.bullet) ? (num(pPr.bullet.level) ?? 0) : undefined;
  const level =
    levels?.[numLevelIndex] ??
    (bulletLevel != null && !numRef ? BUILTIN_BULLET_LEVEL(bulletLevel) : undefined);

  // Indent cascade: direct w:ind > the numbering level's w:ind > style chain
  // > docDefaults. The level beating the style is Word's rule — applying a
  // list re-indents styled paragraphs (ListParagraph's 720tw must not pin
  // every level to level 0's indent). Char-unit attributes beat their twip
  // twins, w:start/w:end are the modern names of w:left/w:right (Word reads
  // each pair as one slot), and w:hanging is firstLine's negative twin
  // winning the pair within a tier — the engine takes it as a negative first
  // line (line 0 starts LEFT of the indent, where a list marker sits).
  const charsPx = (v: unknown): number | undefined => {
    const n = num(v);
    return n != null && n > 0 ? (n / 100) * defaultTextStyle.sizePx : undefined;
  };
  const twPx = (v: unknown): number | undefined => {
    const t = measureTwip(v);
    return t != null ? twipToPx(t) : undefined;
  };
  const leftPx =
    charsPx(dInd.leftChars ?? dInd.startChars) ??
    twPx(dInd.left ?? dInd.start) ??
    (level?.leftTw != null ? twipToPx(level.leftTw) : undefined) ??
    charsPx(pick([sInd, docInd], "leftChars") ?? pick([sInd, docInd], "startChars")) ??
    twPx(pick([sInd, docInd], "left") ?? pick([sInd, docInd], "start"));
  const firstLinePx = (() => {
    const directHanging = charsPx(dInd.hangingChars) ?? twPx(dInd.hanging);
    if (directHanging != null) return -directHanging;
    const directTw = twPx(dInd.firstLine);
    if (directTw != null) return Math.max(0, directTw);
    const directChars = charsPx(dInd.firstLineChars);
    if (directChars != null) return directChars;
    if (level?.hangingTw != null && level.hangingTw > 0) return -twipToPx(level.hangingTw);
    const styleHanging =
      charsPx(pick([sInd, docInd], "hangingChars")) ?? twPx(pick([sInd, docInd], "hanging"));
    if (styleHanging != null) return -styleHanging;
    const styleTw = twPx(pick([sInd, docInd], "firstLine"));
    if (styleTw != null) return Math.max(0, styleTw);
    return charsPx(pick([sInd, docInd], "firstLineChars"));
  })();
  const indent = {
    leftPx: leftPx || undefined,
    rightPx: twipToPx(measureTwip(ind("right") ?? ind("end")) ?? 0) || undefined,
    firstLinePx,
  };

  // Tab stops: twips from the content-box left edge → px from the TEXT-box
  // edge (the engine measures x from the left indent). "decimal" renders as
  // left for now; the exotic bar/clear/end kinds carry no box.
  const tabStops: LayoutTabStop[] | undefined = Array.isArray(pPr.tabStops)
    ? pPr.tabStops.flatMap((ts) => {
        if (!isRecord(ts)) return [];
        const positionPx = measureTwip(ts.position);
        if (positionPx == null) return [];
        const type =
          ts.type === "right" ? "right" : ts.type === "center" ? "center" : ("left" as const);
        const leader =
          ts.leader === "dot" ||
          ts.leader === "heavy" ||
          ts.leader === "hyphen" ||
          ts.leader === "middleDot" ||
          ts.leader === "underscore"
            ? ts.leader
            : undefined;
        return [{ positionPx: twipToPx(positionPx) - (indent.leftPx ?? 0), type, leader }];
      })
    : undefined;

  // Paragraph borders (w:pBdr): direct, else the style chain's.
  const bRec: Rec = isRecord(pPr.border)
    ? pPr.border
    : isRecord(chainPPr.border)
      ? chainPPr.border
      : {};
  const borderEdge = (v: unknown): LayoutParagraphBorderEdge | undefined => {
    if (!isRecord(v)) return undefined;
    const size = num(v.size);
    const space = num(v.space);
    return {
      style: typeof v.style === "string" ? v.style : undefined,
      px: size != null ? (size / 8) * ptToPx(1) : undefined,
      spacePx: space != null ? ptToPx(space) : undefined,
    };
  };
  const borders = {
    top: borderEdge(bRec.top),
    right: borderEdge(bRec.right),
    bottom: borderEdge(bRec.bottom),
    left: borderEdge(bRec.left),
  };
  // Paragraph shading (w:shd): direct, else the style chain's — same
  // direct-else-chain as borders; the fill gate mirrors the cell projection.
  const shdRec: Rec = isRecord(pPr.shading)
    ? pPr.shading
    : isRecord(chainPPr.shading)
      ? chainPPr.shading
      : {};
  const shadingFill =
    typeof shdRec.fill === "string" && shdRec.fill !== "auto" && shdRec.type !== "nil"
      ? shdRec.fill
      : undefined;

  // The list marker: a bullet emits its glyph; a numbered level advances its
  // counter (resetting deeper levels) and substitutes %k in w:lvlText with the
  // formatted counter of level k-1 — "%1.%2" at level 1 → "2.3".
  const markerInline: LayoutInline[] = (() => {
    if (!level) return [];
    if (level.format === "bullet") {
      return level.text
        ? [
            { kind: "text", text: level.text, style: defaultTextStyle, synthetic: true },
            { kind: "tab", toPx: 0 },
          ]
        : [];
    }
    if (level.format === "none") return [];
    if (!numReference) return [];
    const counters = ctx.listCounters.get(numReference) ?? [];
    ctx.listCounters.set(numReference, counters);
    counters[numLevelIndex] = (counters[numLevelIndex] ?? 0) + 1;
    counters.length = numLevelIndex + 1;
    const marker = (level.text ?? "%1.").replace(/%([1-9])/g, (_, k: string) => {
      const idx = Number(k) - 1;
      const lvl = levels?.[idx];
      return formatListNumber(lvl?.format ?? level.format, counters[idx] ?? 1);
    });
    return [
      { kind: "text", text: marker, style: defaultTextStyle, synthetic: true },
      { kind: "tab", toPx: 0 },
    ];
  })();

  // `p?.` not `p.`: compiled/parsed documents can carry `paragraph: null`
  // (an empty paragraph leg) even though the public type says otherwise.
  const runs: readonly unknown[] =
    typeof p === "string" ? [p] : (p?.children ?? (p?.text != null ? [p.text] : []));
  const drawings = projectDrawings(runs, ctx);
  const inline = projectRuns(runs, chainRPr, docRPr, defaultTextStyle, ctx);
  return {
    kind: "paragraph",
    inline: markerInline.length ? markerInline.concat(inline) : inline,
    drawings: drawings.length > 0 ? drawings : undefined,
    spacing,
    indent,
    tabStops: tabStops && tabStops.length > 0 ? tabStops : undefined,
    defaultTabStopPx: ctx.defaultTabStopPx,
    borders: borders.top || borders.right || borders.bottom || borders.left ? borders : undefined,
    shadingFill,
    markSizePx: markSize != null ? ptToPx(markSize) : undefined,
    defaultTextStyle,
    snapToGrid: typeof pPr.snapToGrid === "boolean" ? pPr.snapToGrid : null,
    align: alignOf(pick([pPr, chainPPr, docPPr], "alignment")),
    keepLines: pPr.keepLines === true || chainPPr.keepLines === true,
    keepNext: pPr.keepNext === true || chainPPr.keepNext === true,
    widowControl: pick([pPr, chainPPr], "widowControl") !== false,
    pageBreakBefore: pPr.pageBreakBefore === true || chainPPr.pageBreakBefore === true,
    suppressLineNumbers: pPr.suppressLineNumbers === true || chainPPr.suppressLineNumbers === true,
  };
}
