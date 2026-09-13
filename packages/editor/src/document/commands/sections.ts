import type { BorderOptions, PageBordersOptions, SectionPropertiesOptions } from "@docen/docx";
import { convertMillimetersToTwip, sectionPageSizeDefaults } from "@docen/docx";
import type { Editor } from "@docen/docx/core";

import type { ColumnsValues } from "../../ui/components/workspace/columns-dialog";
import type { LineNumbersValues } from "../../ui/components/workspace/line-numbers-dialog";
import type { PageSetupValues } from "../../ui/components/workspace/page-setup-dialog";
import type { BorderSideState, BordersDialogPatch } from "../extensions/commands";
import { MARGINS, PAPER_SIZES, marginTwipsFromCss, mergeSectionProperties } from "../page-setup";

/** The section commands' view of the host — resolved per call so the
 *  controller can be built before a document opens. */
export interface SectionsHost {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The story bridge — dialog commits target the active story. */
  bridge(): { activeEditor(): Editor; focus(): void } | undefined;
  /** The host element — shadow-DOM root for the dialog components. */
  element(): HTMLElement;
  /** The first section's flow box (column-width budgeting), undefined before
   *  the first layout. */
  flow(): { contentWidthPx: number } | undefined;
}

/**
 * The "this section" domain, split out of the host element: locating the
 * caret's sectPr, reading and mutating it (Word's This Section semantics —
 * a section-carrying paragraph at/after the caret owns it, otherwise the
 * body-level sectPr), the page-setup presets, the line-number/column-count
 * toggles, and the page-setup/columns/borders dialog commits.
 */
export class SectionCommands {
  constructor(private readonly host: SectionsHost) {}

  #target(): Editor | null | undefined {
    return this.host.bridge()?.activeEditor() ?? this.host.editor();
  }

  /** The doc position carrying the current section's sectPr — the first
   *  section-carrying paragraph at/after the caret (OOXML: its sectPr ends
   *  that section), or null when the caret sits in the final section (the
   *  sectPr is body-level on doc.attrs). */
  sectionSectPrPos(): number | null {
    const editor = this.host.editor();
    if (!editor) return null;
    const from = editor.state.selection.from;
    let targetPos: number | null = null;
    editor.state.doc.descendants((node, nodePos) => {
      if (targetPos != null) return true;
      // Paragraphs ending at/before the caret close earlier sections; a
      // paragraph CONTAINING the caret owns the current section (OOXML: its
      // sectPr ends that section, caret position included).
      if (nodePos + node.nodeSize <= from) return true;
      if (
        node.type.name === "paragraph" &&
        (node.attrs as { sectionProperties?: unknown }).sectionProperties != null
      ) {
        targetPos = nodePos;
        return false;
      }
      return true;
    });
    return targetPos;
  }

  /** The current section's sectPr content — the read side of
   *  {@link SectionCommands.updateSectionGeometry}'s write side (same "this
   *  section" rule). */
  currentSectionProperties(): SectionPropertiesOptions | undefined {
    const editor = this.host.editor();
    if (!editor) return undefined;
    const pos = this.sectionSectPrPos();
    if (pos != null) {
      const node = editor.state.doc.nodeAt(pos);
      if (!node) return undefined;
      return (node.attrs as { sectionProperties?: SectionPropertiesOptions }).sectionProperties;
    }
    return (editor.state.doc.attrs as { sectionProperties?: SectionPropertiesOptions })
      .sectionProperties;
  }

  /** Rewrite the current section's sectPr through `mutate` (Word's "this
   *  section" semantics — a section-carrying paragraph at/after the caret
   *  owns it, otherwise the body-level sectPr) and dispatch. The transaction
   *  re-renders every page of the canvas. */
  mutateCurrentSection(
    mutate: (cur: SectionPropertiesOptions | undefined) => SectionPropertiesOptions,
  ): void {
    const editor = this.host.editor();
    if (!editor) return;
    const { doc, tr } = editor.state;
    const targetPos = this.sectionSectPrPos();
    if (targetPos != null) {
      const node = doc.nodeAt(targetPos);
      if (node) {
        const cur = (node.attrs as { sectionProperties?: SectionPropertiesOptions })
          .sectionProperties;
        tr.setNodeMarkup(targetPos, undefined, { ...node.attrs, sectionProperties: mutate(cur) });
      }
    } else {
      const cur = (doc.attrs as { sectionProperties?: SectionPropertiesOptions }).sectionProperties;
      tr.setDocAttribute("sectionProperties", mutate(cur));
    }
    editor.view.dispatch(tr);
  }

  /** Deep-merge a sectionProperties patch into the CURRENT section's sectPr and
   *  dispatch it — Word's "this section" semantics. The dispatched transaction
   *  re-renders every page of the canvas. */
  updateSectionGeometry(patch: SectionPropertiesOptions): void {
    const editor = this.host.editor();
    if (!editor) return;
    this.mutateCurrentSection((cur) => mergeSectionProperties(cur, patch));
  }

  /** Apply a paper-size preset (a4/letter/…) — writes the size into the
   *  document-model sectionProperties (Word stores page setup in the sectPr)
   *  so layout/export share one geometry source. */
  setPageSize(value?: string): void {
    const size = value ? PAPER_SIZES[value] : undefined;
    if (size) {
      this.updateSectionGeometry({
        pageSize: {
          width: convertMillimetersToTwip(size[0]),
          height: convertMillimetersToTwip(size[1]),
        },
      });
    }
  }

  /** Apply orientation (portrait/landscape) — writes orientation onto
   *  page.size, deep-merged with the current (or engine-default) size so the
   *  projection can swap edges for landscape. */
  setOrientation(value?: string): void {
    if (!value) return;
    const cur = this.currentSectionProperties()?.pageSize;
    const size =
      cur && typeof cur.width === "number" && typeof cur.height === "number"
        ? cur
        : { width: sectionPageSizeDefaults.WIDTH, height: sectionPageSizeDefaults.HEIGHT };
    this.updateSectionGeometry({
      pageSize: { ...size, orientation: value as "portrait" | "landscape" },
    });
  }

  /** Apply a margin preset (normal/narrow/…) — writes the margins into the
   *  document-model sectionProperties so a page-setup change actually
   *  re-lays-out. */
  setMargins(value?: string): void {
    if (value && MARGINS[value]) {
      this.updateSectionGeometry({ pageMargin: marginTwipsFromCss(MARGINS[value]) });
    }
  }

  /** Design → Page Borders presets — stamp w:pgBorders on the current
   *  section (Word's Borders and Shading gallery): none clears it; box is a
   *  plain rule; shadow thickens the bottom/right edges; double and dashed
   *  swap the rule's style. Sides measure from the text margin (Word's
   *  default offsetFrom), 0.5 pt black. */
  setPageBorders(preset?: string): void {
    if (!preset) return;
    const side = (style: BorderOptions["style"], size = 4): BorderOptions => ({
      style,
      size,
      space: 0,
    });
    const rule: BorderOptions["style"] =
      preset === "double" ? "double" : preset === "dashed" ? "dashSmallGap" : "single";
    const borders: PageBordersOptions | undefined =
      preset === "none"
        ? undefined
        : preset === "shadow"
          ? {
              offsetFrom: "text",
              top: side("single"),
              left: side("single"),
              bottom: side("single", 18),
              right: side("single", 18),
            }
          : {
              offsetFrom: "text",
              top: side(rule),
              right: side(rule),
              bottom: side(rule),
              left: side(rule),
            };
    // pageBorders rides the top-level spread in mergeSectionProperties (an
    // undefined patch value removes the pgBorders — Word's "none").
    this.updateSectionGeometry({ pageBorders: borders });
  }

  /** Toggle a slot-visibility flag — Word's Different First Page / Odd &
   *  Even Pages. titlePage (w:titlePg) IS a sectPr child and goes to the
   *  current section; evenAndOddHeaders is a settings.xml flag (CT_SectPr
   *  has no such child), so a section-level write would be dropped on
   *  export and the furniture projection (doc.settings) would never see it
   *  — it toggles document-wide through documentExtras instead. */
  toggleSectionFlag(flag: "titlePage" | "evenAndOddHeaders"): void {
    if (flag === "evenAndOddHeaders") {
      const editor = this.host.editor();
      if (!editor) return;
      const { doc, tr } = editor.state;
      const extras =
        (doc.attrs as { documentExtras?: Record<string, unknown> }).documentExtras ?? {};
      const settings = (extras.settings ?? {}) as Record<string, unknown>;
      tr.setDocAttribute("documentExtras", {
        ...extras,
        settings: { ...settings, evenAndOddHeaders: !settings.evenAndOddHeaders },
      });
      editor.view.dispatch(tr);
      return;
    }
    this.mutateCurrentSection((cur) => ({
      ...cur,
      titlePage: !cur?.titlePage,
    }));
  }

  /** Column count for the current section (Word's Page Layout → Columns
   *  presets). The rest of the columns object survives (the gap, the
   *  separator), so toggling back to one column and re-applying keeps the
   *  original geometry. */
  setColumnCount(count: number): void {
    this.mutateCurrentSection((cur) => ({
      ...cur,
      columns: { ...cur?.columns, count },
    }));
  }

  /** Line numbering mode for the current section (w:lnNumType) — Word's
   *  Layout → Line Numbers menu. "none" clears the numbering; the restart
   *  modes map to w:lnNumType's @w:restart values. The rest of the lnNumType
   *  (start/countBy/distance from the options dialog) survives the toggle,
   *  with Word's countBy 1 seeding a fresh numbering. */
  setLineNumbers(mode: "none" | "continuous" | "newPage" | "newSection"): void {
    this.mutateCurrentSection((cur) => ({
      ...cur,
      lineNumberType:
        mode === "none" ? undefined : { countBy: 1, ...cur?.lineNumberType, restart: mode },
    }));
  }

  /** Open the Line Numbering Options dialog prefilled from the current
   *  section's w:lnNumType (the Line Numbers menu's options entry). */
  openLineNumbersOptions(): void {
    const cur = this.currentSectionProperties()?.lineNumberType;
    const atLeast = (v: number | string | undefined): number | undefined =>
      typeof v === "number" && v >= 1 ? v : undefined;
    (
      this.host.element().shadowRoot?.querySelector("docen-line-numbers-dialog") as {
        show(values?: Partial<LineNumbersValues>): void;
      } | null
    )?.show({
      start: atLeast(cur?.start) ?? 1,
      countBy: atLeast(cur?.countBy) ?? 1,
      // Twips → centimeters; absent distance = Word's auto margin placement.
      distance:
        typeof cur?.distance === "number" && cur.distance > 0
          ? Math.round(((cur.distance * 2.54) / 1440) * 100) / 100
          : undefined,
      restart:
        cur?.restart === "newPage" || cur?.restart === "newSection" ? cur.restart : "continuous",
    });
  }

  /** The Line Numbering Options dialog's OK — write start/countBy/restart and
   *  the centimeters-converted distance back onto the current section's
   *  w:lnNumType; the transaction re-renders (the painter draws the numbers).
   *  A cleared distance removes the attribute (Word's 自动). */
  readonly onLineNumbersOk = (event: CustomEvent<LineNumbersValues | undefined>): void => {
    const values = event.detail;
    if (!values) return;
    this.mutateCurrentSection((cur) => {
      const next: SectionPropertiesOptions["lineNumberType"] = {
        ...cur?.lineNumberType,
        start: Math.max(1, Math.round(values.start) || 1),
        countBy: Math.max(1, Math.round(values.countBy) || 1),
        restart: values.restart,
      };
      if (values.distance != null && values.distance > 0) {
        next.distance = convertMillimetersToTwip(values.distance * 10);
      } else {
        delete next.distance;
      }
      return { ...cur, lineNumberType: next };
    });
  };

  /** The Normal pitch — docDefaults' run size in points (Word's zh-CN default
   *  五号 = 10.5 pt when unset). The document grid's character pitch measures
   *  its w:charSpace delta (1/4096 pt) against this base. */
  #normalPitchPt(): number {
    const editor = this.host.editor();
    const styles = (editor ? editor.state.doc.attrs.styles : null) as {
      default?: { document?: { run?: { size?: number } } };
    } | null;
    const size = styles?.default?.document?.run?.size;
    return typeof size === "number" && size > 0 ? size : 10.5;
  }

  /** Open the Page Setup dialog prefilled from the current section's geometry
   *  in centimeters (the Margins menu's Custom Margins and the Size menu's
   *  More Paper Sizes entries). */
  openPageSetup(): void {
    const cur = this.currentSectionProperties();
    // Twips → centimeters for the inputs (2 decimals is Word's display
    // precision); absent geometry — or a UniversalMeasure string form, which
    // the dialog doesn't parse — falls back to Word defaults.
    const cm = (twips?: number | string): number | undefined =>
      typeof twips === "number" ? Math.round(((twips * 2.54) / 1440) * 100) / 100 : undefined;
    // pageMargin/pageSize carry `false` (explicit removal) alongside the
    // properties object — narrow to the object form before reading fields.
    const margin = cur?.pageMargin && typeof cur.pageMargin === "object" ? cur.pageMargin : {};
    const size = cur?.pageSize && typeof cur.pageSize === "object" ? cur.pageSize : {};
    const grid = cur?.grid && typeof cur.grid === "object" ? cur.grid : undefined;
    // PageMargin fields may be UniversalMeasure strings the dialog doesn't
    // parse — narrow to the number form for the usable-height budget.
    const numTw = (v: number | string | undefined, d: number): number =>
      typeof v === "number" ? v : d;
    const heightTw = typeof size.height === "number" ? size.height : 16838;
    const widthTw = typeof size.width === "number" ? size.width : 11906;
    const usableTw = heightTw - numTw(margin.top, 1440) - numTw(margin.bottom, 1440);
    // Lines per page from the pitch (twips → lines, floored — Word's zh A4
    // default is 44, not the rounded 45); absent pitch falls back to Word's
    // CJK default (312 twips).
    const pitchTw = typeof grid?.linePitch === "number" ? grid.linePitch : 312;
    // Characters per line from w:charSpace — the grid's character pitch is
    // the Normal font pitch plus a delta in 1/4096 pt
    // (charsPerLine = floor(textWidth / (pitch + charSpace/4096))); an absent
    // charSpace means the Normal pitch itself (Word's zh A4 default = 39).
    const charSpace = typeof grid?.charSpace === "number" ? grid.charSpace : 0;
    const textWidthPt = (widthTw - numTw(margin.left, 1440) - numTw(margin.right, 1440)) / 20;
    const charsPerLine = Math.max(
      1,
      Math.floor(textWidthPt / (this.#normalPitchPt() + charSpace / 4096)),
    );
    (
      this.host.element().shadowRoot?.querySelector("docen-page-setup-dialog") as {
        show(values?: {
          margins?: Partial<PageSetupValues["margins"]>;
          size?: Partial<PageSetupValues["size"]>;
          verticalAlign?: PageSetupValues["verticalAlign"];
          gutter?: number;
          headerDistance?: number;
          footerDistance?: number;
          titlePage?: boolean;
          sectionStart?: NonNullable<PageSetupValues["sectionStart"]>;
          grid?: NonNullable<PageSetupValues["grid"]>;
        }): void;
      } | null
    )?.show({
      margins: {
        top: cm(margin.top),
        bottom: cm(margin.bottom),
        left: cm(margin.left),
        right: cm(margin.right),
      },
      size: { width: cm(size.width), height: cm(size.height) },
      verticalAlign: cur?.verticalAlign ?? "top",
      gutter: cm(margin.gutter) ?? 0,
      headerDistance: cm(margin.header) ?? 1.5,
      footerDistance: cm(margin.footer) ?? 1.75,
      titlePage: cur?.titlePage === true,
      // "nextColumn" exists in OOXML but has no dropdown slot — Word's Layout
      // tab offers the same four; it falls back to the nextPage default.
      sectionStart:
        cur?.type === "continuous" || cur?.type === "oddPage" || cur?.type === "evenPage"
          ? cur.type
          : "nextPage",
      grid: {
        type: grid?.type ?? "lines",
        charsPerLine,
        linesPerPage: pitchTw > 0 ? Math.floor(usableTw / pitchTw) : undefined,
      },
    });
  }

  /** Open the Columns dialog prefilled from the current section's w:cols
   *  (the Columns menu's More Columns entry). */
  openColumnsDialog(): void {
    const cur = this.currentSectionProperties()?.columns;
    // Twips → centimeters for the inputs; absent fields take Word's defaults
    // inside the dialog.
    const columns =
      cur && typeof cur === "object"
        ? cur
        : ({} as Partial<SectionPropertiesOptions["columns"]> & Record<string, unknown>);
    const cm = (twips?: number | string): number | undefined =>
      typeof twips === "number" ? Math.round(((twips * 2.54) / 1440) * 100) / 100 : undefined;
    const raw = columns as {
      count?: number;
      space?: number | string;
      separate?: boolean;
      equalWidth?: boolean;
    };
    (
      this.host.element().shadowRoot?.querySelector("docen-columns-dialog") as {
        show(values?: Partial<ColumnsValues>): void;
      } | null
    )?.show({
      count: typeof raw.count === "number" ? raw.count : undefined,
      space: cm(raw.space),
      separate: raw.separate === true,
      equalWidth: raw.equalWidth !== false,
    });
  }

  // Open the Borders and Shading dialog on `tab`, prefilling the border tab
  // from the caret paragraph's w:pBdr and the page tab from the current
  // section's w:pgBorders.
  openBordersDialog(tab: "border" | "page" | "shading"): void {
    const target = this.#target();
    const dialog = this.host
      .element()
      .shadowRoot?.querySelector("docen-borders-shading-dialog") as {
      show(tab: "border" | "page" | "shading", border?: unknown, page?: unknown): void;
    } | null;
    if (!target || !dialog) return;
    // The caret paragraph's attrs (formattable block only — a code block or
    // a table cell still carries paragraph attrs here).
    const { $from } = target.state.selection;
    const block = $from.parent.type.isTextblock
      ? ($from.parent.attrs as Record<string, unknown>)
      : null;
    const border = (block?.border ?? null) as Record<string, unknown> | null;
    const page = (this.currentSectionProperties()?.pageBorders ?? null) as Record<
      string,
      unknown
    > | null;
    dialog.show(tab, border, page);
  }

  // The Columns dialog's OK — convert back to twips and write the current
  // section's w:cols. Unequal widths get evenly-split explicit children (the
  // w:col list the projection needs once equalWidth is false); per-column
  // manual widths stay out until the dialog grows inputs for them.
  readonly onColumnsOk = (event: CustomEvent<ColumnsValues | undefined>): void => {
    const values = event.detail;
    if (!values) return;
    const count = Math.max(1, Math.min(9, Math.trunc(values.count) || 1));
    const space = convertMillimetersToTwip(values.space * 10);
    const children =
      values.equalWidth || count <= 1
        ? undefined
        : Array.from({ length: count }, () => ({
            width: Math.max(
              1,
              Math.floor(
                ((this.host.flow()?.contentWidthPx ?? 0) * 15 - space * (count - 1)) / count,
              ),
            ),
          }));
    this.mutateCurrentSection((cur) => ({
      ...cur,
      columns: {
        ...cur?.columns,
        count,
        space,
        // Explicit both ways — a conditional spread would let a stale
        // separate:true from the previous w:cols survive an unchecked box.
        separate: values.separate,
        equalWidth: values.equalWidth,
        ...(children ? { children } : {}),
      },
    }));
  };

  // The Borders and Shading dialog's OK — route by tab: the border tab
  // stamps the selected paragraphs' w:pBdr, the page tab the current
  // section's w:pgBorders, and the shading tab the paragraph fill.
  readonly onBordersShadingOk = (event: CustomEvent<BordersDialogPatch | undefined>): void => {
    const patch = event.detail;
    if (!patch) return;
    if (patch.tab === "shading") {
      const target = this.#target();
      target?.commands.shading?.(patch.fill ? patch.fill : "none");
      this.host.bridge()?.focus();
      return;
    }
    if (patch.tab === "border") {
      const target = this.#target();
      target?.commands["borders-apply"]?.(patch);
      this.host.bridge()?.focus();
      return;
    }
    // Page tab — every edge null removes the pgBorders (Word's "none").
    const sides = patch.sides ?? {};
    const edge = (s: BorderSideState | null | undefined): BorderOptions | undefined =>
      s
        ? {
            style: s.style as BorderOptions["style"],
            size: Math.max(2, Math.round(s.size)),
            color: s.color ?? "auto",
            space: 0,
          }
        : undefined;
    const borders: PageBordersOptions | undefined =
      sides.top || sides.bottom || sides.left || sides.right
        ? {
            offsetFrom: "text",
            top: edge(sides.top),
            left: edge(sides.left),
            bottom: edge(sides.bottom),
            right: edge(sides.right),
          }
        : undefined;
    this.updateSectionGeometry({ pageBorders: borders });
  };

  // The Page Setup dialog's OK — convert its centimeters back to twips (the
  // presets go through the same convertMillimetersToTwip) and write the
  // current section's geometry; the transaction re-renders the canvas.
  readonly onPageSetupOk = (event: CustomEvent<PageSetupValues | undefined>): void => {
    const values = event.detail;
    if (!values) return;
    const twip = (cm: number): number => convertMillimetersToTwip(cm * 10);
    const { margins, size, verticalAlign, gutter, grid } = values;
    // Lines per page → line pitch over the usable page height (the same
    // budget the dialog's prefill divides); a cleared field keeps the
    // section's current pitch. A count the prefill would have derived from
    // the current pitch keeps that pitch untouched — dividing it back would
    // drift the stored value a few twips on every OK.
    const cur = this.currentSectionProperties();
    const curGrid = cur?.grid && typeof cur.grid === "object" ? cur.grid : undefined;
    const usableTw = twip(size.height) - twip(margins.top) - twip(margins.bottom);
    let pitchTw = curGrid?.linePitch ?? 312;
    if (
      typeof grid?.linesPerPage === "number" &&
      grid.linesPerPage > 0 &&
      grid.linesPerPage !== Math.floor(usableTw / pitchTw)
    ) {
      pitchTw = Math.round(usableTw / grid.linesPerPage);
    }
    // Characters per line → w:charSpace, the delta from the Normal font pitch
    // in 1/4096 pt over the usable width. Floored (Word's pairing): only then
    // does floor(textWidth / (pitch + charSpace/4096)) read the same count
    // back — rounding here drops the next reopen to chars−1. Only the
    // char-grid behaviors carry it — a lines-only grid serializes without
    // w:charSpace (Word's 只指定行网格), so the explicit undefined also strips
    // a stale value.
    const charsOk = grid?.type === "linesAndChars" || grid?.type === "snapToChars";
    const textWidthPt = (twip(size.width) - twip(margins.left) - twip(margins.right)) / 20;
    const charSpace =
      charsOk && grid?.charsPerLine != null && grid.charsPerLine > 0
        ? Math.floor((textWidthPt / grid.charsPerLine - this.#normalPitchPt()) * 4096)
        : undefined;
    this.updateSectionGeometry({
      pageMargin: {
        top: twip(margins.top),
        bottom: twip(margins.bottom),
        left: twip(margins.left),
        right: twip(margins.right),
        gutter: twip(gutter ?? 0),
        header: twip(values.headerDistance ?? 1.5),
        footer: twip(values.footerDistance ?? 1.75),
      },
      pageSize: { width: twip(size.width), height: twip(size.height) },
      verticalAlign: verticalAlign === "top" ? undefined : verticalAlign,
      // Omitted w:type reads as "nextPage" (Word drops the attribute); "no
      // grid" clears w:docGrid entirely (Word's 无网格).
      type: values.sectionStart === "nextPage" ? undefined : values.sectionStart,
      titlePage: values.titlePage === true,
      grid: grid?.type === "default" ? false : { type: grid?.type, linePitch: pitchTw, charSpace },
    });
  };
}
