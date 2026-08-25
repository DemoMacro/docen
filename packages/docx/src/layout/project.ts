// The docx adapter — projects office-open's DocumentOptions into
// @docen/layout's LayoutDoc. The PERSISTENCE model is the projection source
// (not the editor's Tiptap JSON subset): every body shape office-open can
// round-trip reaches the layout engine, and shapes the adapter cannot lay out
// yet (toc, sdt, textbox, altChunk, customXml, rawXml) become placeholder
// boxes instead of silently vanishing. Callers chain
// Tiptap JSON --compileDocument--> DocumentOptions --this--> LayoutDoc.
//
// Zero-DOM discipline is inherited from @docen/layout — this module is
// Node-safe (headless export) by construction.

import {
  ptToPx,
  twipToPx,
  emuToPx,
  type LayoutBlock,
  type LayoutBorderEdge,
  type LayoutCellInsets,
  type LayoutInline,
  type LayoutLineHeight,
  type LayoutParagraph,
  type LayoutParagraphBorderEdge,
  type LayoutSpacing,
  type LayoutTabStop,
  type LayoutTable,
  type LayoutTableWidth,
  type LayoutTextStyle,
} from "@docen/layout";
import type {
  DocumentOptions,
  ParagraphOptions,
  SectionChild,
  StylesOptions,
  TableCellOptions,
  TableOptions,
} from "@office-open/docx";

import { resolvePageSize } from "../extensions/utils";
import { defaultParagraphStyleId, indexParagraphStyles, mergeStyleChain } from "../style-cascade";

// The paragraph leg of SectionChild is `string | ParagraphOptions` (shorthand
// or full options); null appears at runtime (empty paragraph legs from
// parse/compile), so the projection accepts it defensively.
type BodyParagraph = string | ParagraphOptions | null;
type LayoutCell = LayoutTable["rows"][number]["cells"][number];

// ── loose-shape guards ──

type Rec = Record<string, unknown>;

/** Options unions are structurally loose at their edges (optional everything,
 *  per-side sub-objects); this guard narrows unknown/union picks to a record so
 *  the rest of the module reads fields without per-site casts. */
function isRecord(v: unknown): v is Rec {
  return !!v && typeof v === "object";
}

const num = (v: unknown): number | undefined => (typeof v === "number" ? v : undefined);

const str = (v: unknown): string | undefined => (typeof v === "string" && v ? v : undefined);

/** Estimated height of one placeholder box: three default body lines. */
const PLACEHOLDER_PX = 3 * 16;

// ── universal-measure parsing (number = native unit, string = UM) ──

const UM_IN_TWIPS = { pt: 20, pc: 240, in: 1440, mm: 1440 / 25.4, cm: 1440 / 2.54, px: 15 };
const UM_RE = /^(-?[\d.]+)(pt|pc|in|mm|cm|px)$/;

/** A measure field to twips: number passes through (native), UM resolves. */
function measureTwip(v: unknown): number | undefined {
  const n = num(v);
  if (n != null) return n;
  if (typeof v !== "string") return undefined;
  const m = UM_RE.exec(v);
  return m ? Number(m[1]) * UM_IN_TWIPS[m[2] as keyof typeof UM_IN_TWIPS] : undefined;
}

/** A measure field whose native unit is EMU (drawing extents): number passes,
 *  UM resolves (px at 96 dpi). */
function measureEmu(v: unknown): number | undefined {
  const n = num(v);
  if (n != null) return n;
  if (typeof v !== "string") return undefined;
  const tw = measureTwip(v);
  return tw != null ? (tw / 1440) * 914400 : undefined;
}

// ── picture media (renderer passthrough) ──

const MIME_BY_TYPE: Record<string, string> = {
  jpg: "image/jpeg",
  png: "image/png",
  gif: "image/gif",
  bmp: "image/bmp",
  tif: "image/tiff",
  ico: "image/x-icon",
  svg: "image/svg+xml",
};

/** PictureOptions (type + data) → a data URL the painter can load. Absent
 *  bytes (linked-only) or undecodable input yields undefined — the renderer
 *  draws an empty frame. */
function pictureSrc(pic: Rec): string | undefined {
  if (typeof pic.type !== "string") return undefined;
  const mime = MIME_BY_TYPE[pic.type];
  if (!mime) return undefined;
  const { data } = pic;
  if (typeof data === "string") {
    return data.startsWith("data:") ? data : `data:${mime};base64,${data}`;
  }
  if (data instanceof Uint8Array) {
    // Chunked binary → base64 (btoa is universal: browsers and Node ≥ 16).
    let bin = "";
    for (let i = 0; i < data.length; i += 0x8000) {
      bin += String.fromCharCode(...data.subarray(i, i + 0x8000));
    }
    return `data:${mime};base64,${btoa(bin)}`;
  }
  return undefined;
}

// ── style cascade (direct pPr → style chain → docDefaults) ──

/** `default.document` — the docDefaults object (run directly on it, paragraph
 *  props nested one level down under `paragraph`). */
function docDefaultsOf(styles: StylesOptions | undefined): Rec {
  const doc = styles?.default?.document;
  return isRecord(doc) ? doc : {};
}

/** w:jc (AlignmentType, ST_Jc) → the engine's alignment semantics. The
 *  kashida/thai/numericTab variants (Arabic elongation, Thai word-break
 *  justification, list tab alignment) have no faithful canvas algorithm —
 *  they fall back to the left default until one lands. */
const ALIGN_TO_LAYOUT = {
  left: "left",
  start: "left",
  right: "right",
  end: "right",
  center: "center",
  both: "both",
  distribute: "distribute",
} as const;

/** The merged {run, paragraph} for a style id — the same mergeStyleChain the
 *  DOM route's stylesToCss resolves, so both renderings share one cascade.
 *  A style-less paragraph resolves to the default paragraph style (usually
 *  Normal); docDefaults sits UNDER the chain, not in it. */
function styleChainOf(styles: StylesOptions | undefined, styleId: string | null | undefined) {
  if (!styles) return { run: {}, paragraph: {} };
  return mergeStyleChain(indexParagraphStyles(styles), styleId || defaultParagraphStyleId(styles));
}

const pick = (layers: Rec[], key: string): unknown => {
  for (const layer of layers) if (layer[key] != null) return layer[key];
  return undefined;
};

/** The cascaded w:jc value → the engine's alignment (undefined → left). */
function alignOf(jc: unknown): LayoutParagraph["align"] {
  if (typeof jc !== "string") return undefined;
  return jc in ALIGN_TO_LAYOUT ? ALIGN_TO_LAYOUT[jc as keyof typeof ALIGN_TO_LAYOUT] : undefined;
}

// ── run/style resolution ──

/** OOXML font: string or rFonts {ascii, hAnsi, eastAsia} → engine slots. */
type FontAttr = string | Rec | null | undefined;

/** An unknown font pick → the FontAttr domain (string or rFonts record). */
function fontAttr(v: unknown): FontAttr {
  return isRecord(v) || typeof v === "string" ? v : undefined;
}

function toFamily(font: FontAttr, def: FontAttr): LayoutTextStyle["family"] {
  const f = font ?? def;
  if (typeof f === "string") return f;
  const latin = str(f?.ascii) ?? str(f?.hAnsi);
  const eastAsia = str(f?.eastAsia);
  return latin || eastAsia ? { latin, eastAsia } : {};
}

interface RunStyle {
  sizePt?: number;
  font?: FontAttr;
  characterSpacingTw?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  underline?: boolean;
  strikethrough?: boolean;
}

/** rPr (a run's own, or the ¶-mark/paragraph default) → resolved fields. */
function runStyleOf(rPr: Rec): RunStyle {
  const underline = isRecord(rPr.underline)
    ? rPr.underline.type !== "none"
      ? true
      : undefined
    : undefined;
  return {
    sizePt: num(rPr.size),
    font: fontAttr(rPr.font),
    characterSpacingTw: measureTwip(rPr.characterSpacing),
    bold: rPr.bold === true ? true : undefined,
    italic: rPr.italic === true ? true : undefined,
    color: typeof rPr.color === "string" && rPr.color !== "auto" ? rPr.color : undefined,
    underline,
    strikethrough: rPr.strike === true || rPr.doubleStrike === true ? true : undefined,
  };
}

// ── numbering (list) resolution ──

/** One numbering level's layout-relevant fields (w:lvl). */
interface NumberingLevel {
  format: string;
  text: string;
  leftTw?: number;
  hangingTw?: number;
}

/** reference → levels indexed by w:lvl/@w:ilvl. Bullet levels render today;
 *  numbered formats (decimal…) need a document-order counter — a registered
 *  gap (the projection is a pure per-paragraph walk today). */
type NumberingIndex = Map<string, NumberingLevel[]>;

function indexNumberings(numbering: unknown): NumberingIndex {
  const index: NumberingIndex = new Map();
  if (!isRecord(numbering) || !Array.isArray(numbering.abstractNumberings)) return index;
  for (const abs of numbering.abstractNumberings) {
    if (!isRecord(abs)) continue;
    const reference = str(abs.reference);
    const levels: NumberingLevel[] = [];
    if (reference && Array.isArray(abs.levels)) {
      for (const lvl of abs.levels) {
        if (!isRecord(lvl)) continue;
        const ind: Rec =
          isRecord(lvl.paragraph) && isRecord(lvl.paragraph.indent) ? lvl.paragraph.indent : {};
        levels[num(lvl.level) ?? 0] = {
          format: typeof lvl.format === "string" ? lvl.format : "bullet",
          text: typeof lvl.text === "string" ? lvl.text : "",
          leftTw: measureTwip(ind.left),
          hangingTw: measureTwip(ind.hanging),
        };
      }
      index.set(reference, levels);
    }
  }
  return index;
}

/** Per-document projection context, resolved once and threaded down. */
interface ProjectContext {
  styles: StylesOptions | undefined;
  numberings: NumberingIndex;
  /** Live list counters per numbering reference (level → count), advanced in
   *  document order as numbered paragraphs project. */
  listCounters: Map<string, number[]>;
}

// ── list-number formats (w:numFmt) ──

const CJK_DIGITS = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
const CJK_UNITS = ["", "十", "百", "千"];

/** chineseCounting composition (零 fill between non-zero groups; the 10-19
 *  range drops the leading 一). */
function chineseNumeral(n: number): string {
  if (n < 1 || n > 9999) return String(n);
  const digits: number[] = [];
  for (let rest = n; rest > 0; rest = Math.floor(rest / 10)) digits.unshift(rest % 10);
  let out = "";
  let zeroPending = false;
  digits.forEach((d, i) => {
    const unit = CJK_UNITS[digits.length - 1 - i];
    if (d === 0) {
      if (out) zeroPending = true;
      return;
    }
    if (zeroPending) {
      out += CJK_DIGITS[0];
      zeroPending = false;
    }
    // 10-19 is 十X, not 一十X.
    if (!(d === 1 && unit === "十" && digits.length === 2)) out += CJK_DIGITS[d];
    out += unit;
  });
  return out;
}

const ROMAN_PAIRS: [number, string][] = [
  [1000, "M"],
  [900, "CM"],
  [500, "D"],
  [400, "CD"],
  [100, "C"],
  [90, "XC"],
  [50, "L"],
  [40, "XL"],
  [10, "X"],
  [9, "IX"],
  [5, "V"],
  [4, "IV"],
  [1, "I"],
];

function romanNumeral(n: number, upper: boolean): string {
  let rest = n;
  let out = "";
  for (const [value, glyph] of ROMAN_PAIRS) {
    while (rest >= value) {
      out += glyph;
      rest -= value;
    }
  }
  return upper ? out : out.toLowerCase();
}

/** 1→a…26→z, 27→aa (spreadsheet-style, Word's letter numbering). */
function letterNumeral(n: number, upper: boolean): string {
  let out = "";
  let rest = n;
  while (rest > 0) {
    rest--;
    out = String.fromCharCode(97 + (rest % 26)) + out;
    rest = Math.floor(rest / 26);
  }
  return upper ? out.toUpperCase() : out;
}

/** One level's counter under its w:numFmt. Unsupported formats render decimal. */
function formatListNumber(format: string, n: number): string {
  switch (format) {
    case "lowerLetter":
      return letterNumeral(n, false);
    case "upperLetter":
      return letterNumeral(n, true);
    case "lowerRoman":
      return romanNumeral(n, false);
    case "upperRoman":
      return romanNumeral(n, true);
    case "chineseCounting":
    case "chineseLegalSimplified":
    case "japaneseCounting":
      return chineseNumeral(n);
    default:
      return String(n);
  }
}

// ── paragraph projection ──

function toLineHeight(line: number | undefined, rule: unknown): LayoutLineHeight | undefined {
  if (line == null) return undefined;
  if (rule === "exact") return { rule: "exact", px: twipToPx(line) };
  if (rule === "atLeast") return { rule: "atLeast", px: twipToPx(line) };
  return { rule: "multiple", factor: line / 240 };
}

function projectParagraph(p: BodyParagraph, ctx: ProjectContext): LayoutParagraph {
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
  const defaultTextStyle: LayoutTextStyle = {
    family: toFamily(null, defFont),
    sizePx: ptToPx(markSizePt),
    bold: chainRPr.bold === true || docRPr.bold === true || undefined,
    italic: chainRPr.italic === true || docRPr.italic === true || undefined,
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
  // prepends + a tab hop to the body-text start.
  const numRef: Rec | null = isRecord(pPr.numbering)
    ? pPr.numbering
    : isRecord(chainPPr.numbering)
      ? chainPPr.numbering
      : null;
  const numReference = numRef ? str(numRef.reference) : undefined;
  const numLevelIndex = num(numRef?.level) ?? 0;
  const levels = numReference ? ctx.numberings.get(numReference) : undefined;
  const level = levels?.[numLevelIndex];

  let leftTw = measureTwip(ind("left"));
  if (leftTw == null && level?.leftTw != null) leftTw = level.leftTw;
  const firstLineTw = measureTwip(ind("firstLine"));
  const firstLineChars = num(ind("firstLineChars"));
  const indent = {
    leftPx: leftTw != null ? twipToPx(leftTw) || undefined : undefined,
    rightPx: twipToPx(measureTwip(ind("right")) ?? 0) || undefined,
    firstLinePx:
      firstLineTw != null
        ? Math.max(0, twipToPx(firstLineTw))
        : firstLineChars != null && firstLineChars > 0
          ? (firstLineChars / 100) * defaultTextStyle.sizePx
          : level?.hangingTw != null && level.hangingTw > 0
            ? -twipToPx(level.hangingTw)
            : undefined,
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
        return [{ positionPx: twipToPx(positionPx) - (indent.leftPx ?? 0), type }];
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

  // The list marker: a bullet emits its glyph; a numbered level advances its
  // counter (resetting deeper levels) and substitutes %k in w:lvlText with the
  // formatted counter of level k-1 — "%1.%2" at level 1 → "2.3".
  const markerInline: LayoutInline[] = (() => {
    if (!level) return [];
    if (level.format === "bullet") {
      return level.text
        ? [
            { kind: "text", text: level.text, style: defaultTextStyle },
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
      { kind: "text", text: marker, style: defaultTextStyle },
      { kind: "tab", toPx: 0 },
    ];
  })();

  // `p?.` not `p.`: compiled/parsed documents can carry `paragraph: null`
  // (an empty paragraph leg) even though the public type says otherwise.
  const runs: readonly unknown[] =
    typeof p === "string" ? [p] : (p?.children ?? (p?.text != null ? [p.text] : []));
  return {
    kind: "paragraph",
    inline: markerInline.length
      ? markerInline.concat(projectRuns(runs, chainRPr, docRPr, defaultTextStyle))
      : projectRuns(runs, chainRPr, docRPr, defaultTextStyle),
    spacing,
    indent,
    tabStops: tabStops && tabStops.length > 0 ? tabStops : undefined,
    borders: borders.top || borders.right || borders.bottom || borders.left ? borders : undefined,
    markSizePx: markSize != null ? ptToPx(markSize) : undefined,
    defaultTextStyle,
    snapToGrid: typeof pPr.snapToGrid === "boolean" ? pPr.snapToGrid : null,
    align: alignOf(pick([pPr, chainPPr, docPPr], "alignment")),
    keepLines: pPr.keepLines === true || chainPPr.keepLines === true,
    keepNext: pPr.keepNext === true || chainPPr.keepNext === true,
    widowControl: pick([pPr, chainPPr], "widowControl") !== false,
    pageBreakBefore: pPr.pageBreakBefore === true || chainPPr.pageBreakBefore === true,
  };
}

/** Inline content: text runs (rPr resolved over the paragraph default), hard
 *  breaks, and pictures (paragraph-child or run-child slot) as atoms. The
 *  members arrive as unknown — the ParagraphChild union is wide and its
 *  runtime shapes are looser still (compile pushes `{text, …rPr}` run forms),
 *  so each leg is validated rather than trusted.
 *  Known-but-unprojected inline atoms (tab, chart, math, fields, hyperlinks)
 *  carry no box yet — they render as absence, a registered gap to close type
 *  by type. */
function projectRuns(
  runs: readonly unknown[],
  chainRPr: Rec,
  docRPr: Rec,
  defRun: LayoutTextStyle,
): LayoutInline[] {
  const out: LayoutInline[] = [];
  const textStyleOf = (rPr: Rec): LayoutTextStyle => {
    const own = runStyleOf(rPr);
    const font: FontAttr = own.font ?? fontAttr(chainRPr.font) ?? fontAttr(docRPr.font) ?? null;
    return {
      family: font != null ? toFamily(own.font, font) : defRun.family,
      sizePx: ptToPx(own.sizePt ?? num(chainRPr.size) ?? num(docRPr.size) ?? 12),
      bold: own.bold ?? defRun.bold,
      italic: own.italic ?? defRun.italic,
      color: own.color ?? defRun.color,
      underline: own.underline ?? defRun.underline,
      strikethrough: own.strikethrough ?? defRun.strikethrough,
      letterSpacingPx:
        own.characterSpacingTw != null ? twipToPx(own.characterSpacingTw) : undefined,
    };
  };
  const pushText = (text: string, rPr: Rec): void => {
    if (!text) return;
    out.push({ kind: "text", text, style: textStyleOf(rPr) });
  };
  /** A field (w:fldSimple / complexField): PAGE/NUMPAGES become dynamic atoms
   *  (the painter resolves the number per page — `text` is a measuring
   *  placeholder); anything else renders its cached result as static text. */
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
    } else if (cached) {
      pushText(cached, rPr);
    }
  };
  const pushPicture = (pic: Rec): void => {
    const tr = isRecord(pic.transformation) ? pic.transformation : {};
    const w = measureEmu(tr.width);
    const h = measureEmu(tr.height);
    if (w != null && h != null) {
      out.push({
        kind: "picture",
        widthPx: emuToPx(w),
        heightPx: emuToPx(h),
        src: pictureSrc(pic),
      });
    }
  };
  for (const child of runs) {
    if (typeof child === "string") {
      pushText(child, {});
      continue;
    }
    if (!isRecord(child)) continue;
    if (typeof child.text === "string") pushText(child.text, child);
    if (child.break != null) out.push({ kind: "break" });
    if (child.tab != null) out.push({ kind: "tab" });
    if (isRecord(child.picture)) pushPicture(child.picture);
    if (isRecord(child.complexField)) pushField(child.complexField, child);
    if (isRecord(child.simpleField)) pushField(child.simpleField, child);
    if (Array.isArray(child.children)) {
      for (const inner of child.children) {
        if (typeof inner === "string") {
          pushText(inner, child);
        } else if (isRecord(inner)) {
          if (typeof inner.text === "string") pushText(inner.text, inner);
          else if (inner.break != null) out.push({ kind: "break" });
          else if (inner.tab != null) out.push({ kind: "tab" });
          else if (isRecord(inner.picture)) pushPicture(inner.picture);
          else if (isRecord(inner.complexField)) pushField(inner.complexField, inner);
          else if (isRecord(inner.simpleField)) pushField(inner.simpleField, inner);
        }
      }
    }
  }
  return out;
}

// ── table projection ──

function toTableWidth(w: unknown): LayoutTableWidth | undefined {
  if (!isRecord(w)) return undefined;
  if (w.type === "auto" || w.type === "nil") return undefined;
  if (w.type === "percent") {
    const size = w.size;
    const pct =
      typeof size === "string" && size.endsWith("%") ? Number(size.slice(0, -1)) : num(size);
    if (pct != null && Number.isFinite(pct)) return { type: "percent", percent: pct };
    return undefined;
  }
  const tw = measureTwip(w.size);
  return tw != null && tw > 0 ? { type: "px", px: twipToPx(tw) } : undefined;
}

function toCellInsets(m: unknown): LayoutCellInsets | undefined {
  if (!isRecord(m)) return undefined;
  const side = (v: unknown): number | undefined => {
    const size = isRecord(v) ? measureTwip(v.size) : undefined;
    return size != null ? twipToPx(size) : undefined;
  };
  const insets = {
    top: side(m.top),
    right: side(m.right),
    bottom: side(m.bottom),
    left: side(m.left),
  };
  return insets.top != null || insets.right != null || insets.bottom != null || insets.left != null
    ? insets
    : undefined;
}

type CellBorders = NonNullable<LayoutCell["borders"]>;

/** One w:tcBorders/w:tblBorders edge → px + color (nil/none survive as declared
 *  zero-weight edges; the conflict resolver skips them). */
function toBorderEdge(v: unknown): LayoutBorderEdge | undefined {
  if (!isRecord(v)) return undefined;
  const size = num(v.size);
  const color = typeof v.color === "string" && v.color !== "auto" ? v.color : undefined;
  return {
    style: typeof v.style === "string" ? v.style : undefined,
    px: size != null ? (size / 8) * ptToPx(1) : undefined,
    color,
  };
}

function toBorders(b: unknown): CellBorders | undefined {
  if (!isRecord(b)) return undefined;
  const out = {
    top: toBorderEdge(b.top),
    right: toBorderEdge(b.right),
    bottom: toBorderEdge(b.bottom),
    left: toBorderEdge(b.left),
  };
  return out.top || out.right || out.bottom || out.left ? out : undefined;
}

/** w:tblBorders → the engine's table-level defaults, merging the direct
 *  tblPr borders over the table style's per side. */
function toTableBorders(direct: unknown, styleTable: unknown): LayoutTable["borders"] | undefined {
  const d = isRecord(direct) ? direct : undefined;
  const s = isRecord(styleTable)
    ? isRecord(styleTable.borders)
      ? styleTable.borders
      : undefined
    : undefined;
  if (!d && !s) return undefined;
  const edge = (side: string): LayoutBorderEdge | undefined =>
    toBorderEdge(d?.[side]) ?? toBorderEdge(s?.[side]);
  const out = {
    top: edge("top"),
    bottom: edge("bottom"),
    left: edge("left"),
    right: edge("right"),
    insideHorizontal: edge("insideHorizontal"),
    insideVertical: edge("insideVertical"),
  };
  return out.top ||
    out.bottom ||
    out.left ||
    out.right ||
    out.insideHorizontal ||
    out.insideVertical
    ? out
    : undefined;
}

function projectCell(c: TableCellOptions, ctx: ProjectContext): LayoutCell {
  const shd = isRecord(c.shading) ? c.shading : undefined;
  const fill =
    shd && typeof shd.fill === "string" && shd.fill !== "auto" && shd.type !== "nil"
      ? shd.fill
      : undefined;
  return {
    colspan: c.columnSpan,
    rowspan: c.rowSpan,
    insets: toCellInsets(c.margins),
    borders: toBorders(c.borders),
    fill,
    blocks: c.children
      .map((child) => projectChild(child, ctx))
      .filter((b): b is LayoutBlock => b !== null),
  };
}

function projectTable(t: TableOptions, ctx: ProjectContext): LayoutTable {
  const rows: LayoutTable["rows"] = [];
  for (const row of t.rows ?? []) {
    // sdt/customXml row wrappers have no cells of their own — a later gap.
    if (!("cells" in row)) continue;
    const trHeight: Rec = isRecord(row.height) ? row.height : {};
    const heightValue = measureTwip(trHeight.value);
    const height =
      heightValue != null && heightValue > 0
        ? {
            rule: trHeight.rule === "exact" ? ("exact" as const) : ("atLeast" as const),
            px: twipToPx(heightValue),
          }
        : undefined;
    rows.push({
      cells: (row.cells ?? [])
        .filter((cell): cell is TableCellOptions => "children" in cell)
        .map((cell) => projectCell(cell, ctx)),
      height,
    });
  }

  const columnWidthsPx = t.columnWidths?.map((w) => twipToPx(measureTwip(w) ?? 0));
  const styleTable = t.style
    ? ctx.styles?.tableStyles?.find((st) => st.id === t.style)?.table
    : undefined;
  return {
    kind: "table",
    width: toTableWidth(t.width),
    columnWidthsPx: columnWidthsPx && columnWidthsPx.length > 0 ? columnWidthsPx : undefined,
    cellInsets: toCellInsets(t.margins),
    borders: toTableBorders(t.borders, styleTable),
    rows,
  };
}

// ── section-child dispatch ──

/** One body child → layout block. Unprojectable shapes become placeholder
 *  boxes; zero-height markers (bookmarks) vanish — both OOXML-faithful. */
function projectChild(child: SectionChild, ctx: ProjectContext): LayoutBlock | null {
  if ("paragraph" in child) return projectParagraph(child.paragraph, ctx);
  if ("table" in child) return projectTable(child.table, ctx);
  if ("bookmarkStart" in child || "bookmarkEnd" in child) return null;
  // toc, sdt, textbox, altChunk, customXml, rawXml → a labeled box.
  const label = Object.keys(child)[0];
  return { kind: "placeholder", heightPx: PLACEHOLDER_PX, label };
}

// ── section flow geometry ──

/** The flow box a section defines, in px: paper size minus margins
 *  (orientation already resolved by resolvePageSize) and the docGrid pitch. */
export interface ProjectedFlowBox {
  pageWidthPx: number;
  pageHeightPx: number;
  contentWidthPx: number;
  contentHeightPx: number;
  /** Content-box origin within the page (margin left/top) — where the flow's
   *  (0,0) sits on paper; the painter anchors page content here. */
  contentLeftPx: number;
  contentTopPx: number;
  linePitchPx?: number;
}

export function projectFlowBox(properties: unknown): ProjectedFlowBox {
  const sp: Rec = isRecord(properties) ? properties : {};
  const { width, height } = resolvePageSize(sp.pageSize);
  const m: Rec = isRecord(sp.pageMargin) ? sp.pageMargin : {};
  const side = (v: unknown, d: number): number => twipToPx(measureTwip(v) ?? d);
  const top = side(m.top, 1440);
  const bottom = side(m.bottom, 1440);
  const left = side(m.left, 1800);
  const right = side(m.right, 1800);
  const grid: Rec = isRecord(sp.grid) ? sp.grid : {};
  const pitchTw = measureTwip(grid.linePitch);
  const linePitchPx =
    grid.type && grid.type !== "default" && pitchTw && pitchTw > 0 ? twipToPx(pitchTw) : undefined;
  return {
    pageWidthPx: twipToPx(width),
    pageHeightPx: twipToPx(height),
    contentWidthPx: twipToPx(width) - left - right,
    contentHeightPx: twipToPx(height) - top - bottom,
    contentLeftPx: left,
    contentTopPx: top,
    linePitchPx,
  };
}

/** Page furniture (headers/footers) projected for painting: the block lists
 *  per slot (already projected like body blocks — the stage lays them out once
 *  at the content width) plus the placement flags read at paint time.
 *  `headerDistancePx`/`footerDistancePx` are w:pgMar's @w:header/@w:footer
 *  (page edge to the header/footer box; 720 twips = Word's default). */
export interface ProjectedPageFurniture {
  header?: LayoutBlock[];
  firstHeader?: LayoutBlock[];
  evenHeader?: LayoutBlock[];
  footer?: LayoutBlock[];
  firstFooter?: LayoutBlock[];
  evenFooter?: LayoutBlock[];
  /** w:titlePg — page 1 uses the `first` slots instead of `default`. */
  titlePage: boolean;
  /** settings' w:evenAndOddHeaders — even pages use the `even` slots. */
  evenAndOddHeaders: boolean;
  headerDistancePx: number;
  footerDistancePx: number;
}

/** Project the first section's headers/footers. An absent slot stays
 *  undefined (the painter falls back per OOXML: page 1 without titlePage and
 *  even pages without evenAndOddHeaders both use `default`). */
function projectPageFurniture(doc: DocumentOptions): ProjectedPageFurniture {
  const section = doc.sections?.[0];
  const ctx: ProjectContext = {
    styles: doc.styles,
    numberings: indexNumberings(doc.numbering),
    listCounters: new Map(),
  };
  const projectSlots = (side: unknown): LayoutBlock[] | undefined => {
    if (!Array.isArray(side)) return undefined;
    const blocks: LayoutBlock[] = [];
    for (const child of side) {
      const block = projectChild(child, ctx);
      if (block) blocks.push(block);
    }
    return blocks.length > 0 ? blocks : undefined;
  };
  const props: Rec = isRecord(section?.properties) ? section.properties : {};
  const margin: Rec = isRecord(props.pageMargin) ? props.pageMargin : {};
  return {
    header: projectSlots(section?.headers?.default),
    firstHeader: projectSlots(section?.headers?.first),
    evenHeader: projectSlots(section?.headers?.even),
    footer: projectSlots(section?.footers?.default),
    firstFooter: projectSlots(section?.footers?.first),
    evenFooter: projectSlots(section?.footers?.even),
    titlePage: props.titlePage === true,
    evenAndOddHeaders: doc.settings?.evenAndOddHeaders === true,
    headerDistancePx: twipToPx(measureTwip(margin.header) ?? 720),
    footerDistancePx: twipToPx(measureTwip(margin.footer) ?? 720),
  };
}

/** Project a full DocumentOptions into the engine's input: the FIRST section's
 *  body and flow box (multi-section flow — later sectPrs arrive as body-level
 *  section breaks — is a later milestone; sections beyond the first are
 *  concatenated into the first's flow). */
export function projectDocumentOptions(doc: DocumentOptions): {
  blocks: LayoutBlock[];
  flow: ProjectedFlowBox;
  furniture: ProjectedPageFurniture;
} {
  const ctx: ProjectContext = {
    styles: doc.styles,
    numberings: indexNumberings(doc.numbering),
    listCounters: new Map(),
  };
  const blocks: LayoutBlock[] = [];
  for (const section of doc.sections ?? []) {
    for (const child of section.children ?? []) {
      const block = projectChild(child, ctx);
      if (block) blocks.push(block);
    }
  }
  return {
    blocks,
    flow: projectFlowBox(doc.sections?.[0]?.properties),
    furniture: projectPageFurniture(doc),
  };
}
