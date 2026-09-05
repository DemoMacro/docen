import { quickStyles, type StylesOptions } from "@docen/docx";

import type {
  RibbonButton,
  RibbonColorPicker,
  RibbonCombobox,
  RibbonControl,
  RibbonControlOrLayout,
  RibbonControlSize,
  RibbonGallery,
  RibbonGroup,
  RibbonLayout,
  RibbonMenu,
  RibbonMenuItem,
  RibbonSeparator,
  RibbonSplit,
  RibbonTab,
} from "../ui";
/**
 * Default MS Office Word ribbon for `<docen-document>` — all 9 standard tabs
 * (Home/Insert/Draw/Design/Layout/References/Mailings/Review/View) with the
 * canonical groups and primary commands.
 *
 * `ribbonTabs()` builds the RibbonTab schema with i18n keys (not translated
 * strings); `renderRibbonFromSchema()` consumes that tree and resolves every
 * label via `t("ribbon.*")` when building the ribbon DOM. The
 * host stamps the result into its `<docen-ribbon>` and re-runs it on language
 * change. Callers wanting a tailored ribbon merge their own tabs/groups into
 * the schema before render.
 *
 * Layout helpers (`.rb-col` / `.rb-row` / `.rb-vsep`) are injected by the host
 * style — Office groups stack a large button beside rows/columns of small
 * `icon-only` buttons.
 *
 * Each command carries an `event` name. `DocumentCommands` (extensions/commands)
 * wires the ones the Tiptap engine supports today (marks, lists, alignment,
 * styles, breaks, history); the rest render as a complete visual skeleton and
 * no-op on click until wired.
 */
import { resolveLang, t, registerIcon } from "../ui";
import { TABLE_STYLE_PRESETS, type TableStylePreset } from "./extensions/commands";
import { FONT_NAMES, FONT_SIZES_CN, FONT_SIZES_PT, UNDERLINE_STYLES } from "./font-lists";

// --- i18n shortcuts (ribbon.* keys, resolved at call time) -------------------

// Ribbon i18n keys — the schema stores keys (not translated strings); the
// render pass (renderRibbonFromSchema) resolves them via t(). t() returns the
// key itself when no translation is registered, so an addin may also pass a
// plain display string as a label (escape hatch).
const tab = (id: string): string => `ribbon.tab.${id}`;
const grp = (id: string): string => `ribbon.group.${id}`;
const cmd = (event: string): string => `ribbon.cmd.${event}`;
const opt = (value: string): string => `ribbon.opt.${value}`;

// --- Option sets (menu/combobox items) ---------------------------------------

const fontItems = (): string => JSON.stringify(FONT_NAMES.map((text) => ({ text })));
// Font-size options: a zh locale lists the Chinese names ("小四 (12)") above
// the point sizes; other locales show point sizes only. The emitted `value` is
// always the pt string, so the two lists stay compatible across locales.
const sizeItems = (): string => {
  const zh = resolveLang().toLowerCase().startsWith("zh");
  const cn = zh
    ? FONT_SIZES_CN.map(([name, pt]) => ({ text: `${name} (${pt})`, value: String(pt) }))
    : [];
  const pt = FONT_SIZES_PT.map((p) => ({ text: String(p), value: String(p) }));
  return JSON.stringify([...cn, ...pt]);
};

/** Minimal built-in set shown when a document carries no styles.xml (e.g. a
 *  blank editor) so the Styles gallery is never empty. */
const FALLBACK_STYLE_ITEMS = (): string =>
  JSON.stringify([
    { text: opt("normal"), value: "Normal" },
    { text: opt("heading-1"), value: "Heading1" },
    { text: opt("heading-2"), value: "Heading2" },
    { text: opt("heading-3"), value: "Heading3" },
    { text: opt("title"), value: "Title" },
  ]);

/** Build the Styles gallery items from the loaded document's styles.xml model:
 *  named paragraph styles (Normal + any custom) first, then the built-in named
 *  styles nested under `default` (title/heading1-9). Display text is the style's
 *  own `name` from the model (falling back to its id); the value is the pStyle
 *  id, which round-trips via the paragraph `style` attr. */
const styleItems = (styles?: StylesOptions | null): string => {
  // quickStyles() returns the document's quickFormat paragraph styles (Word's
  // Quick Styles gallery behavior), ordered by uiPriority. The value is the
  // pStyle id, which round-trips via the Paragraph/Heading styleId attr.
  const entries = quickStyles(styles);
  if (entries.length === 0) return FALLBACK_STYLE_ITEMS();
  return JSON.stringify(entries.map((e) => ({ text: e.name, value: e.id })));
};

const pasteItems = (): string =>
  JSON.stringify([
    { text: opt("paste"), value: "paste" },
    { text: opt("paste-special"), value: "paste-special" },
    { text: opt("keep-text-only"), value: "keep-text-only" },
  ]);

// Edit / View mode pick — the tab-row "Editing" trailing action. Default is
// Edit checked; the host (#syncEditModeMenu in document/index.ts) rewrites the
// label + checked state to match the live editable state, so this is only the
// initial stamp.
const editItems = (): string =>
  JSON.stringify([
    { text: opt("editing"), event: "edit-mode", value: "edit", checked: true },
    { text: opt("viewing"), event: "edit-mode", value: "view" },
  ]);

// Word's Underline split menu — the ST_Underline patterns plus the clear
// entry ("none" and the pattern list live in font-lists.ts, shared with the
// Font dialog's Underline style dropdown).
const underlineItems = (): string =>
  JSON.stringify([
    { text: opt("none"), event: "underline-style", value: "none" },
    ...UNDERLINE_STYLES.map(([value, key]) => ({
      text: opt(key),
      event: "underline-style",
      value,
    })),
  ]);

const caseItems = (): string =>
  JSON.stringify([
    { text: opt("sentence-case"), value: "sentence" },
    { text: opt("lowercase"), value: "lower" },
    { text: opt("uppercase"), value: "upper" },
    { text: opt("capitalize"), value: "capitalize" },
    { text: opt("toggle-case"), value: "toggle" },
  ]);

// Word's Accept / Reject splits (Review → Tracking): the face accepts or
// rejects the selected revision and moves to the next; the drop-down repeats
// that and adds the accept/reject-all sweep. ("Accept All Changes Shown"
// needs the markup view filter — stays out until that exists.)
const acceptItems = (): string =>
  JSON.stringify([
    { text: opt("accept-and-next"), event: "accept-change" },
    { text: opt("accept-all-changes"), event: "accept-all-changes" },
  ]);

const rejectItems = (): string =>
  JSON.stringify([
    { text: opt("reject-and-next"), event: "reject-change" },
    { text: opt("reject-all-changes"), event: "reject-all-changes" },
  ]);

// Word's Chinese Layout (中文版式) drop-down in the Paragraph group — both
// entries open the shared two-lines-in-one dialog (the dialog's bracket
// checkbox covers 合并字符's no-bracket form).
const chineseLayoutItems = (): string =>
  JSON.stringify([
    { text: opt("combine-characters"), event: "two-lines-in-one", value: "combine-characters" },
    { text: opt("two-lines-in-one"), event: "two-lines-in-one", value: "two-lines-in-one" },
  ]);

const bulletItems = (): string =>
  JSON.stringify([
    { text: opt("bullet"), value: "bullet" },
    { text: opt("circle"), value: "circle" },
    { text: opt("square"), value: "square" },
    // Word's Change List Level — demote one level (the menu stand-in for Tab).
    { text: opt("change-list-level"), value: "in", event: "multilevel-list" },
  ]);

const numberItems = (): string =>
  JSON.stringify([
    { text: opt("decimal"), value: "decimal" },
    { text: opt("lower-alpha"), value: "lower-alpha" },
    { text: opt("lower-roman"), value: "lower-roman" },
    { text: opt("change-list-level"), value: "in", event: "multilevel-list" },
  ]);

// Word's multilevel List Library — each entry names its marker shape (the
// gallery thumbnails' text) and applies that preset; the last entry opens
// Define New Multilevel List.
const multilevelItems = (): string =>
  JSON.stringify([
    { text: "1., 1.1., 1.1.1.", event: "multilevel-list", value: "preset:cascade" },
    { text: "1), 1.1), 1.1.1)", event: "multilevel-list", value: "preset:cascade-paren" },
    { text: "1., a., i.", event: "multilevel-list", value: "preset:hybrid" },
    { text: "1), a), i)", event: "multilevel-list", value: "preset:hybrid-paren" },
    { text: "一、（一）1.", event: "multilevel-list", value: "preset:cjk" },
    { text: "第X章 第X节", event: "multilevel-list", value: "preset:cjk-chapter" },
    { text: opt("define-new-multilevel-list"), event: "define-new-list" },
  ]);

const spacingItems = (): string =>
  JSON.stringify([
    { text: "1.0", value: "1.0" },
    { text: "1.15", value: "1.15" },
    { text: "1.5", value: "1.5" },
    { text: "2.0", value: "2.0" },
    { text: opt("add-before"), value: "add-before" },
    { text: opt("add-after"), value: "add-after" },
  ]);

const borderItems = (): string =>
  JSON.stringify([
    { text: opt("no-border"), value: "none" },
    { text: opt("bottom"), value: "bottom" },
    { text: opt("top"), value: "top" },
    { text: opt("left"), value: "left" },
    { text: opt("right"), value: "right" },
    { text: opt("all"), value: "all" },
    { text: opt("outside"), value: "outside" },
    { text: opt("borders-shading"), value: "borders-shading" },
  ]);

const findItems = (): string =>
  JSON.stringify([
    { text: opt("find"), value: "find" },
    { text: opt("go-to"), value: "go-to" },
  ]);

const selectItems = (): string =>
  JSON.stringify([
    { text: opt("select-all"), value: "all" },
    // Need a canvas selection model for objects / similar-format picks.
    { text: opt("select-objects"), value: "objects", disabled: true },
    { text: opt("select-similar"), value: "similar", disabled: true },
  ]);

const coverItems = (): string =>
  JSON.stringify([
    { text: opt("cover"), value: "cover", event: "cover-page" },
    { text: opt("blank"), value: "blank", event: "blank-page" },
    { text: cmd("page-break"), value: "page-break", event: "page-break" },
    { text: cmd("section-break"), value: "section-break", event: "section-break" },
  ]);

// Word's Object menu: OLE embedding is not built (greyed), but "Text from
// File" reads a plain-text file in at the caret.
const objectItems = (): string =>
  JSON.stringify([
    { text: opt("object-dialog"), value: "object", disabled: true },
    { text: opt("file-text"), value: "file", event: "insert-file-text" },
  ]);

const tableItems = (): string =>
  JSON.stringify([
    // Insert Table opens the classic dialog shape of the table grid entry.
    { text: opt("insert-table"), value: "insert", event: "table-dialog" },
    // Draw Table / Convert Text / Excel / Quick Tables are not built yet.
    { text: opt("draw-table"), value: "draw", disabled: true },
    { text: opt("convert-text"), value: "convert", disabled: true },
    { text: opt("excel"), value: "excel", disabled: true },
    { text: opt("quick-tables"), value: "quick", disabled: true },
  ]);

const marginsItems = (): string =>
  JSON.stringify([
    { text: opt("normal-margin"), value: "normal" },
    { text: opt("narrow"), value: "narrow" },
    { text: opt("moderate"), value: "moderate" },
    { text: opt("wide"), value: "wide" },
    // Custom Margins opens the Page Setup dialog (host-handled).
    { text: opt("custom-margins"), value: "custom" },
  ]);

// Word's Cell Margins presets; "custom" opens the cell options dialog (not
// built yet) and stays greyed.
const cellMarginItems = (): string =>
  JSON.stringify([
    { text: opt("cell-margin-normal"), value: "default" },
    { text: opt("cell-margin-none"), value: "none" },
    { text: opt("narrow"), value: "narrow" },
    { text: opt("wide"), value: "wide" },
    { text: opt("custom-margins"), value: "custom", disabled: true },
  ]);

const orientationItems = (): string =>
  JSON.stringify([
    { text: opt("portrait"), value: "portrait" },
    { text: opt("landscape"), value: "landscape" },
  ]);

const sizePaperItems = (): string =>
  JSON.stringify([
    { text: opt("letter"), value: "letter" },
    { text: opt("legal"), value: "legal" },
    { text: opt("tabloid"), value: "tabloid" },
    { text: opt("a3"), value: "a3" },
    { text: opt("a4"), value: "a4" },
    { text: opt("a5"), value: "a5" },
    { text: opt("b5"), value: "b5" },
    { text: opt("statement"), value: "statement" },
    { text: opt("executive"), value: "executive" },
    // More Paper Sizes opens the same Page Setup dialog (host-handled).
    { text: opt("more-sizes"), value: "more" },
  ]);

const columnsItems = (): string =>
  JSON.stringify([
    { text: opt("one-col"), value: "1" },
    { text: opt("two-col"), value: "2" },
    { text: opt("three-col"), value: "3" },
    // More Columns opens the Columns dialog (host-handled).
    { text: opt("more-columns"), value: "more" },
  ]);

const breaksItems = (): string =>
  JSON.stringify([
    { text: cmd("page-break"), value: "page-break", event: "page-break" },
    { text: opt("column-break"), value: "column-break", event: "column-break" },
    // Text Wrapping opens Word's layout-options dialog (not built yet).
    { text: opt("text-wrapping"), value: "text-wrapping", event: "text-wrapping", disabled: true },
    // Word's four section-break types; Even/Odd Page need engine support for
    // starting sections on even/odd pages (with blank interleaves) — greyed
    // until then.
    { text: opt("next-page-section"), value: "section-break-next", event: "section-break-next" },
    {
      text: opt("continuous-section"),
      value: "section-break-continuous",
      event: "section-break-continuous",
    },
    { text: opt("even-page-section"), value: "even", disabled: true },
    { text: opt("odd-page-section"), value: "odd", disabled: true },
  ]);

// Word's Line Numbers menu: the numbering mode writes w:lnNumType on the
// current section; the trailing options entry opens Word's Line Numbering
// dialog (start-at / count-by / distance) — greyed until that dialog exists.
const lineNumbersItems = (): string =>
  JSON.stringify([
    { text: opt("no-line-numbers"), value: "none", event: "line-numbers" },
    { text: opt("continuous-line-numbers"), value: "continuous", event: "line-numbers" },
    { text: opt("restart-each-page"), value: "newPage", event: "line-numbers" },
    { text: opt("restart-each-section"), value: "newSection", event: "line-numbers" },
    { text: opt("line-numbering-options"), value: "options", disabled: true },
  ]);

// Word's Hyphenation menu. The layout engine has no hyphenation dictionary or
// soft-break insertion, so every mode is greyed until it lands — the control
// is here so the Layout tab matches Word's silhouette.
const hyphenationItems = (): string =>
  JSON.stringify([
    { text: opt("hyphenation-none"), value: "none", disabled: true },
    { text: opt("hyphenation-manual"), value: "manual", disabled: true },
    { text: opt("hyphenation-auto"), value: "auto", disabled: true },
    { text: opt("hyphenation-options"), value: "options", disabled: true },
  ]);

const indentItems = (): string =>
  JSON.stringify([
    { text: opt("increase-indent"), value: "increase", event: "indent-increase" },
    { text: opt("decrease-indent"), value: "decrease", event: "indent-decrease" },
  ]);

const groupItems = (): string =>
  JSON.stringify([
    { text: cmd("group"), value: "group" },
    { text: opt("ungroup"), value: "ungroup" },
  ]);

const rotateItems = (): string =>
  JSON.stringify([
    { text: opt("rotate-right"), value: "right" },
    { text: opt("rotate-left"), value: "left" },
    { text: opt("flip-vertical"), value: "flip-v" },
    { text: opt("flip-horizontal"), value: "flip-h" },
  ]);

// The Arrange group's floating-drawing menus: Wrap Text (Word's six) and the
// Position gallery's nine-cell grid (option texts reuse the 9-grid keys).
const wrapItems = (): string =>
  JSON.stringify([
    { text: opt("wrap-front"), value: "front" },
    { text: opt("wrap-behind"), value: "behind" },
    { text: opt("wrap-square"), value: "square" },
    { text: opt("wrap-tight"), value: "tight" },
    { text: opt("wrap-through"), value: "through" },
    { text: opt("wrap-top-bottom"), value: "top-bottom" },
  ]);

const positionItems = (): string =>
  JSON.stringify(
    ["tl", "tc", "tr", "ml", "mc", "mr", "bl", "bc", "br"].map((cell) => ({
      text: opt(`cell-align-${cell}`),
      value: cell,
    })),
  );

const alignObjectsItems = (): string =>
  JSON.stringify([
    { text: cmd("align-left"), value: "left" },
    { text: cmd("align-center"), value: "center" },
    { text: cmd("align-right"), value: "right" },
  ]);

// References > Add Text: the TOC levels (Word's menu minus the missing-level
// caption line).
const addTextItems = (): string =>
  JSON.stringify([
    { text: opt("add-text-level-1"), value: "level-1" },
    { text: opt("add-text-level-2"), value: "level-2" },
    { text: opt("add-text-level-3"), value: "level-3" },
    { text: opt("add-text-none"), value: "none" },
  ]);

// References > Table of Contents drop-down (Word lists two auto galleries;
// one auto build covers the same command here): the auto build, the custom
// dialog, and the removal.
const tocItems = (): string =>
  JSON.stringify([
    { text: opt("toc-auto"), value: "toc", event: "toc" },
    { text: opt("toc-custom"), value: "toc-custom", event: "toc-dialog" },
    { text: opt("toc-remove"), value: "remove-toc", event: "remove-toc" },
  ]);

// References > Update Table: Word's page-numbers-only vs whole-table pass.
const updateTocItems = (): string =>
  JSON.stringify([
    { text: opt("toc-update-page"), value: "update-toc-page", event: "update-toc-page" },
    { text: opt("toc-update-all"), value: "update-toc", event: "update-toc" },
  ]);

const footnoteItems = (): string =>
  JSON.stringify([
    { text: cmd("insert-footnote"), value: "footnote" },
    { text: opt("endnote"), value: "endnote" },
    { text: opt("next-footnote"), value: "next" },
    { text: opt("previous-footnote"), value: "prev" },
  ]);

const startMergeItems = (): string =>
  JSON.stringify([
    { text: opt("letters"), value: "letters" },
    { text: opt("email"), value: "email" },
    { text: opt("envelopes"), value: "envelopes" },
    { text: cmd("labels"), value: "labels" },
    { text: opt("directory"), value: "directory" },
  ]);

const finishMergeItems = (): string =>
  JSON.stringify([
    { text: opt("edit-docs"), value: "edit" },
    { text: opt("print-docs"), value: "print" },
    { text: opt("send-email"), value: "email" },
  ]);

const zoomItems = (): string =>
  JSON.stringify([
    { text: opt("200"), value: "200" },
    { text: opt("100"), value: "100" },
    { text: opt("75"), value: "75" },
    { text: opt("50"), value: "50" },
    { text: opt("page-width"), value: "page-width" },
    { text: opt("text-width"), value: "text-width" },
    { text: opt("fit-page"), value: "fit-page" },
    { text: opt("zoom-dialog"), value: "zoom-dialog" },
  ]);

/** The Page Borders split's presets — Word's Borders and Shading gallery in
 *  menu form: none, a plain box, Word's shadow box (thick bottom/right), a
 *  double rule, and a dashed rule. */
const pageBorderItems = (): string =>
  JSON.stringify([
    { text: opt("page-border-none"), value: "none" },
    { text: opt("page-border-box"), value: "box" },
    { text: opt("page-border-shadow"), value: "shadow" },
    { text: opt("page-border-double"), value: "double" },
    { text: opt("page-border-dashed"), value: "dashed" },
    { text: opt("borders-shading"), value: "borders-shading" },
  ]);

/** The Design tab's style-set gallery (Word's Document Formatting group): the
 *  document's opening model ("default") plus three font-family presets. The
 *  presets stamp through the style-set command; "default" restores from the
 *  host's open snapshot. */
const styleSetItems = (): string =>
  JSON.stringify([
    { text: opt("style-set-default"), value: "default" },
    { text: opt("style-set-modern"), value: "modern" },
    { text: opt("style-set-classic"), value: "classic" },
    { text: opt("style-set-elegant"), value: "elegant" },
  ]);

/** The Watermark split's drop-down: Word's preset gallery plus Remove. */
const watermarkItems = (): string =>
  JSON.stringify([
    { text: opt("watermark-confidential-1"), value: "confidential-1" },
    { text: opt("watermark-confidential-2"), value: "confidential-2" },
    { text: opt("watermark-confidential-3"), value: "confidential-3" },
    { text: opt("watermark-urgent"), value: "urgent" },
    { text: opt("watermark-asap"), value: "asap" },
    { text: opt("watermark-draft"), value: "draft" },
    { text: opt("watermark-sample"), value: "sample" },
    { text: opt("watermark-remove"), value: "remove" },
    { text: opt("watermark-custom"), value: "custom" },
  ]);

/** The Header/Footer split's drop-down: edit (the split's main action),
 *  remove, and the slot-visibility flags. Static seed only — the host
 *  re-stamps the items with live `checked` flags on every transaction
 *  (the flags live in sectionProperties, which the static schema can't
 *  read). */
const storyMenuItems = (kind: "header" | "footer"): string =>
  JSON.stringify([
    { text: opt(kind === "header" ? "edit-header" : "edit-footer"), value: "edit" },
    {
      text: opt(kind === "header" ? "remove-header" : "remove-footer"),
      value: kind === "header" ? "remove-header" : "remove-footer",
    },
    { text: opt("different-first"), value: "title-page" },
    { text: opt("odd-even"), value: "odd-even" },
  ]);

/** The Page Number split's drop-down: placement (the main button is the
 *  Word default, bottom of page) plus removal. */
const pageNumberItems = (): string =>
  JSON.stringify([
    { text: opt("page-num-top"), value: "top" },
    { text: opt("page-num-bottom"), value: "bottom" },
    { text: opt("remove-page-numbers"), value: "remove-numbers" },
  ]);

/** The Paragraph Spacing menu (Word's Design → Paragraph Spacing): document-
 *  default spacing presets stamped onto the styles' docDefaults — Word's
 *  "default" restores its factory 8pt-after / 1.08-line spacing. */
const paragraphSpacingItems = (): string =>
  JSON.stringify([
    { text: opt("paragraph-spacing-default"), value: "default" },
    { text: opt("paragraph-spacing-none"), value: "none" },
    { text: opt("paragraph-spacing-compact"), value: "compact" },
    { text: opt("narrow"), value: "narrow" },
    { text: opt("wide"), value: "wide" },
  ]);

/** The Equation menu: the common OMML structures as empty-argument templates
 *  (Word's equation tool's frequent structures) — the canvas paints each as a
 *  dashed placeholder slot until the math layout engine lands. */
const equationItems = (): string =>
  JSON.stringify([
    { text: opt("equation-fraction"), value: "fraction" },
    { text: opt("equation-script"), value: "superScript" },
    { text: opt("equation-radical"), value: "radical" },
    { text: opt("equation-sum"), value: "sum" },
    { text: opt("equation-integral"), value: "integral" },
  ]);

/** The Shapes gallery — the presets the canvas paints today (box presets,
 *  ellipse, straight line); values are the ST_ShapeType tokens verbatim. */
const shapeItems = (): string =>
  JSON.stringify([
    { text: opt("shape-rect"), value: "rect", event: "shapes" },
    { text: opt("shape-round-rect"), value: "roundRect", event: "shapes" },
    { text: opt("shape-ellipse"), value: "ellipse", event: "shapes" },
    { text: opt("shape-line"), value: "line", event: "shapes" },
  ]);

// --- Tabs --------------------------------------------------------------------

/** Default active tab id. */
export const DEFAULT_RIBBON_TAB = "home";

/** The full ordered set of ribbon tab ids (Home → View). */
export const RIBBON_TAB_IDS = [
  "home",
  "insert",
  "draw",
  "design",
  "layout",
  "references",
  "mailings",
  "review",
  "view",
] as const;
export type RibbonTabId = (typeof RIBBON_TAB_IDS)[number];

/** Options for {@link buildRibbonInnerHTML}. */
export interface RibbonOptions {
  /** Whitelist of tab ids to render; omitted/empty = all tabs (back-compat). */
  tabs?: readonly RibbonTabId[];
}

/**
 * Build the ribbon DOM (fluent-tablist + one panel per tab + trailing actions)
 * imperatively from a {@link RibbonTab} schema tree. Same shape as the old
 * HTML-string builder, but typed and data-driven — an addin merges its own
 * tabs/groups into {@link ribbonTabs} before this runs, so the ribbon is
 * externally customizable without host internals. Re-call on a locale change
 * (labels re-resolve in the schema).
 *
 * `tabs` is already the visible subset ({@link ribbonTabs}); the active tab
 * falls back to the first so fluent-tablist never points activeid at a missing id.
 */
export function renderRibbonFromSchema(
  tabs: readonly RibbonTab[],
  actions: readonly RibbonControl[] = [],
  scope: Element = document.documentElement,
): DocumentFragment {
  const frag = document.createDocumentFragment();

  const tablist = document.createElement("fluent-tablist");
  tablist.setAttribute("slot", "tabs");
  tablist.setAttribute("appearance", "transparent");
  const activeId = tabs[0]?.id ?? "";
  if (activeId) tablist.setAttribute("activeid", activeId);
  frag.append(tablist);

  for (const tab of tabs) {
    if (tab.contextual) continue; // the host appends these on selection
    const built = buildContextualTab(tab, scope);
    tablist.append(built.tab);
    frag.append(built.panel);
  }

  for (const c of actions) {
    const el = buildControl(c, scope);
    el.setAttribute("slot", "actions");
    frag.append(el);
  }
  return frag;
}

/** Build one tab's DOM pair — the `docen-ribbon-tab` heading and its
 *  `docen-ribbon-panel` of groups. The static render appends these into the
 *  tablist/panel container; the host reuses the same builder to append/remove
 *  contextual tabs (Table Design/Layout) as the selection enters/leaves their
 *  owning context. */
export function buildContextualTab(
  tab: RibbonTab,
  scope: Element = document.documentElement,
): { tab: HTMLElement; panel: HTMLElement } {
  const tabEl = document.createElement("docen-ribbon-tab");
  tabEl.setAttribute("slot", "tab");
  tabEl.id = tab.id;
  tabEl.textContent = t(tab.label, scope);
  const panel = document.createElement("docen-ribbon-panel");
  panel.setAttribute("value", tab.id);
  for (const g of tab.groups) panel.append(buildGroup(g, scope));
  return { tab: tabEl, panel };
}

function buildGroup(g: RibbonGroup, scope: Element): HTMLElement {
  const el = document.createElement("docen-ribbon-group");
  el.setAttribute("label", t(g.label, scope));
  if (g.launcher) el.setAttribute("launcher", g.launcher);
  for (const c of g.controls) el.append(buildControlOrLayout(c, scope));
  return el;
}

function buildControlOrLayout(c: RibbonControlOrLayout, scope: Element): HTMLElement {
  return c.type === "layout" ? buildLayout(c, scope) : buildControl(c, scope);
}

function buildLayout(l: RibbonLayout, scope: Element): HTMLElement {
  const el = document.createElement("div");
  el.className = l.layout === "column" ? "rb-col" : l.layout === "row" ? "rb-row" : "rb-grid";
  for (const c of l.controls) el.append(buildControlOrLayout(c, scope));
  return el;
}

/** Stamp the shared base attrs (icon/label/event/value/iconOnly/size/disabled)
 *  every control component reads. */
function applyBase(
  el: HTMLElement,
  c: {
    icon?: string;
    label?: string;
    event?: string;
    value?: string;
    iconOnly?: boolean;
    size?: RibbonControlSize;
    disabled?: boolean;
  },
  scope: Element,
): void {
  if (c.icon) el.setAttribute("icon", c.icon);
  if (c.label) el.setAttribute("label", t(c.label, scope));
  if (c.event) el.setAttribute("event", c.event);
  if (c.value) el.setAttribute("value", c.value);
  if (c.iconOnly) el.setAttribute("icon-only", "");
  if (c.size) el.setAttribute("size", c.size);
  if (c.disabled) el.setAttribute("disabled", "");
}

/** Resolve each item's `text` i18n key to the active locale — the schema stores
 *  keys, the render pass is the single translate point (mirrors label). */
const translateItems = (items: readonly RibbonMenuItem[], scope: Element): RibbonMenuItem[] =>
  items.map((it) => ({ ...it, text: it.text ? t(it.text, scope) : it.text }));

function buildControl(c: RibbonControl, scope: Element): HTMLElement {
  switch (c.type) {
    case "separator": {
      const el = document.createElement("span");
      el.className = "rb-vsep";
      return el;
    }
    case "button": {
      const el = document.createElement("docen-ribbon-button");
      applyBase(el, c, scope);
      return el;
    }
    case "checkbox": {
      const el = document.createElement("docen-ribbon-checkbox");
      applyBase(el, c, scope);
      if (c.checked) el.setAttribute("checked", "");
      return el;
    }
    case "menu": {
      const el = document.createElement("docen-ribbon-menu");
      applyBase(el, c, scope);
      el.setAttribute("items", JSON.stringify(translateItems(c.items ?? [], scope)));
      return el;
    }
    case "split": {
      const el = document.createElement("docen-ribbon-split-button");
      applyBase(el, c, scope);
      el.setAttribute("items", JSON.stringify(translateItems(c.items ?? [], scope)));
      return el;
    }
    case "combobox": {
      const el = document.createElement("docen-ribbon-combobox");
      applyBase(el, c, scope);
      if (c.value != null) el.setAttribute("value", c.value);
      el.setAttribute("items", JSON.stringify(translateItems(c.items ?? [], scope)));
      if (c.source) el.setAttribute("source", c.source);
      if (c.comboboxSize === "short") el.setAttribute("size", "short");
      return el;
    }
    case "color-picker": {
      const el = document.createElement("docen-color-picker");
      applyBase(el, c, scope);
      if (c.defaultColor) el.setAttribute("default-color", c.defaultColor);
      if (c.palette) el.setAttribute("palette", c.palette);
      return el;
    }
    case "gallery": {
      const el = document.createElement("docen-ribbon-gallery");
      applyBase(el, c, scope);
      el.setAttribute("items", JSON.stringify(translateItems(c.items ?? [], scope)));
      if (c.visibleCount != null) el.setAttribute("visible-count", String(c.visibleCount));
      return el;
    }
  }
}

// --- Data-driven ribbon (RibbonTab tree) -------------------------------------
// The 9 tabs expressed as data (i18n keys, not translated strings);
// renderRibbonFromSchema consumes this tree to build the ribbon DOM and
// resolves the keys via t() — so re-rendering on a locale change re-localizes.

/** Parse a legacy items JSON string (the form the *Panel helpers emit) into
 *  RibbonMenuItem data for the data-driven ribbon. */
const parsedItems = (json: string): RibbonMenuItem[] => JSON.parse(json) as RibbonMenuItem[];

const col = (controls: readonly RibbonControlOrLayout[]): RibbonLayout => ({
  type: "layout",
  layout: "column",
  controls,
});
const row = (controls: readonly RibbonControlOrLayout[]): RibbonLayout => ({
  type: "layout",
  layout: "row",
  controls,
});
const grid = (controls: readonly RibbonControlOrLayout[]): RibbonLayout => ({
  type: "layout",
  layout: "grid",
  controls,
});
const sep = (): RibbonSeparator => ({ type: "separator" });

const btn = (
  icon: string,
  event: string,
  o: { size?: "large"; iconOnly?: boolean } = {},
): RibbonButton => ({
  type: "button",
  icon,
  event,
  label: cmd(event),
  ...(o.size ? { size: o.size } : {}),
  ...(o.iconOnly ? { iconOnly: true } : {}),
});

const split = (
  icon: string,
  event: string,
  items: RibbonMenuItem[],
  o: { size?: "large"; iconOnly?: boolean; label?: string } = {},
): RibbonSplit => ({
  type: "split",
  icon,
  event,
  label: o.label ?? cmd(event),
  items,
  ...(o.size ? { size: o.size } : {}),
  ...(o.iconOnly ? { iconOnly: true } : {}),
});

const menu = (
  icon: string,
  event: string,
  items: RibbonMenuItem[],
  o: { label?: string; size?: "large"; iconOnly?: boolean } = {},
): RibbonMenu => ({
  type: "menu",
  icon,
  event,
  label: o.label ?? cmd(event),
  items,
  ...(o.size ? { size: o.size } : {}),
  ...(o.iconOnly ? { iconOnly: true } : {}),
});

const combo = (
  event: string,
  value: string,
  items: RibbonMenuItem[],
  o: { source?: "local-fonts"; comboboxSize?: "short" } = {},
): RibbonCombobox => ({
  type: "combobox",
  event,
  value,
  items,
  ...(o.source ? { source: o.source } : {}),
  ...(o.comboboxSize ? { comboboxSize: o.comboboxSize } : {}),
});

const picker = (
  icon: string,
  event: string,
  defaultColor: string,
  o: { palette?: "theme" | "highlight" } = {},
): RibbonColorPicker => ({
  type: "color-picker",
  icon,
  event,
  label: cmd(event),
  defaultColor,
  iconOnly: true,
  ...(o.palette ? { palette: o.palette } : {}),
});

const group = (
  id: string,
  controls: readonly RibbonControlOrLayout[],
  launcher?: string,
): RibbonGroup => ({
  id,
  label: grp(id),
  controls,
  ...(launcher ? { launcher } : {}),
});

const tabNode = (id: RibbonTabId, groups: RibbonGroup[]): RibbonTab => ({
  id,
  label: tab(id),
  groups,
});

/** Default ribbon tabs for the active locale (and the loaded document's styles,
 *  for the Styles gallery). Pass `{ tabs }` to render a subset. */
export function ribbonTabs(styles?: StylesOptions | null, opts: RibbonOptions = {}): RibbonTab[] {
  const visible: readonly RibbonTabId[] =
    opts.tabs && opts.tabs.length > 0 ? opts.tabs : RIBBON_TAB_IDS;
  const show = (id: RibbonTabId): boolean => visible.includes(id);
  const tabs: RibbonTab[] = [];
  if (show("home")) tabs.push(homeTab(styles));
  if (show("insert")) tabs.push(insertTab());
  if (show("draw")) tabs.push(drawTab());
  if (show("design")) tabs.push(designTab());
  if (show("layout")) tabs.push(layoutTab());
  if (show("references")) tabs.push(referencesTab());
  if (show("mailings")) tabs.push(mailingsTab());
  if (show("review")) tabs.push(reviewTab());
  if (show("view")) tabs.push(viewTab());
  return tabs;
}

/** Trailing ribbon actions (right of the tabs): comment / edit-mode / share. */
export function ribbonActions(): RibbonControl[] {
  return [
    btn("comment", "comment"),
    menu("edit", "edit-mode", parsedItems(editItems()), { label: cmd("editing") }),
    btn("share", "share"),
  ];
}

const homeTab = (styles?: StylesOptions | null): RibbonTab =>
  tabNode("home", [
    group(
      "clipboard",
      [
        split("paste", "paste", parsedItems(pasteItems()), { size: "large" }),
        col([
          btn("cut", "cut", { iconOnly: true }),
          btn("copy", "copy", { iconOnly: true }),
          btn("format-painter", "format-painter", { iconOnly: true }),
        ]),
      ],
      "clipboard-dialog",
    ),
    group(
      "font",
      [
        col([
          row([
            combo("font-name", "Microsoft YaHei", parsedItems(fontItems()), {
              source: "local-fonts",
            }),
            combo("font-size", "14", parsedItems(sizeItems()), { comboboxSize: "short" }),
            btn("font-size", "grow-font", { iconOnly: true }),
            btn("font-size", "shrink-font", { iconOnly: true }),
            split("case", "change-case", parsedItems(caseItems()), { iconOnly: true }),
            btn("clear-format", "clear-format", { iconOnly: true }),
          ]),
          row([
            btn("bold", "bold", { iconOnly: true }),
            btn("italic", "italic", { iconOnly: true }),
            split("underline", "underline", parsedItems(underlineItems()), { iconOnly: true }),
            btn("strike", "strike", { iconOnly: true }),
            btn("superscript", "superscript", { iconOnly: true }),
            btn("subscript", "subscript", { iconOnly: true }),
            sep(),
            btn("phonetic-guide", "phonetic-guide", { iconOnly: true }),
            sep(),
            picker("highlight", "highlight", "FFFF00", { palette: "highlight" }),
            picker("font-color", "font-color", "000000"),
          ]),
        ]),
      ],
      "font-dialog",
    ),
    group(
      "paragraph",
      [
        col([
          row([
            split("list", "bullet-list", parsedItems(bulletItems()), { iconOnly: true }),
            split("numbering", "ordered-list", parsedItems(numberItems()), { iconOnly: true }),
            split("multilevel", "multilevel-list", parsedItems(multilevelItems()), {
              iconOnly: true,
            }),
            btn("indent-decrease", "indent-decrease", { iconOnly: true }),
            btn("indent-increase", "indent-increase", { iconOnly: true }),
            split("two-in-one", "two-lines-in-one", parsedItems(chineseLayoutItems()), {
              iconOnly: true,
            }),
            btn("sort", "sort", { iconOnly: true }),
            btn("show-marks", "show-marks", { iconOnly: true }),
          ]),
          row([
            btn("align-left", "align-left", { iconOnly: true }),
            btn("align-center", "align-center", { iconOnly: true }),
            btn("align-right", "align-right", { iconOnly: true }),
            btn("justify", "justify", { iconOnly: true }),
            btn("align-distribute", "justify-distribute", { iconOnly: true }),
            sep(),
            split("line-spacing", "line-spacing", parsedItems(spacingItems()), { iconOnly: true }),
            picker("shading", "shading", "FFFF00"),
            split("border", "border", parsedItems(borderItems()), { iconOnly: true }),
          ]),
        ]),
      ],
      "paragraph-dialog",
    ),
    group(
      "styles",
      [col([combo("style", "Normal", parsedItems(styleItems(styles)))])],
      "styles-pane",
    ),
    group(
      "editing",
      [
        split("search", "search", parsedItems(findItems()), { size: "large" }),
        col([
          btn("replace", "replace", { iconOnly: true }),
          split("board", "select", parsedItems(selectItems()), { iconOnly: true }),
        ]),
      ],
      "find-dialog",
    ),
  ]);

const insertTab = (): RibbonTab =>
  tabNode("insert", [
    group("pages", [
      split("page-break", "page-break", parsedItems(coverItems()), { size: "large" }),
    ]),
    group("tables", [
      split("table-add", "insert-table", parsedItems(tableItems()), { size: "large" }),
    ]),
    group("illustrations", [
      btn("picture", "insert-picture", { size: "large" }),
      btn("online-picture", "online-picture", { size: "large" }),
      split("shapes", "shapes", parsedItems(shapeItems()), { size: "large" }),
      btn("icon-library", "icons", { size: "large" }),
      btn("3d-model", "3d-model", { size: "large" }),
      btn("smartart", "smartart", { size: "large" }),
      btn("chart", "chart", { size: "large" }),
      btn("insert-picture", "screenshot", { size: "large" }),
    ]),
    group("links", [
      btn("hyperlink", "link", { size: "large" }),
      btn("bookmark", "bookmark", { size: "large" }),
    ]),
    group("comments", [btn("comment-add", "comment", { size: "large" })]),
    group("header-footer", [
      split("header", "header", parsedItems(storyMenuItems("header")), { size: "large" }),
      split("footer", "footer", parsedItems(storyMenuItems("footer")), { size: "large" }),
      split("page-number", "page-number", parsedItems(pageNumberItems()), { size: "large" }),
    ]),
    group("text", [
      btn("text-box", "text-box", { size: "large" }),
      btn("wordart", "wordart", { size: "large" }),
      btn("date-time", "date-time", { size: "large" }),
      menu("object", "object", parsedItems(objectItems()), { size: "large" }),
    ]),
    group("symbols", [
      menu("equation", "equation", parsedItems(equationItems()), { size: "large" }),
      btn("symbol", "symbol", { size: "large" }),
    ]),
  ]);

const drawTab = (): RibbonTab =>
  tabNode("draw", [
    group("pens", [
      btn("pen", "draw-pen", { size: "large" }),
      col([
        grid([
          btn("pencil", "draw-pencil"),
          btn("highlight", "draw-highlighter"),
          btn("eraser", "draw-eraser"),
        ]),
      ]),
    ]),
    group("draw-tools", [
      btn("lasso", "lasso-select", { size: "large" }),
      col([grid([btn("board", "select-objects"), btn("action-pen", "action-pen")])]),
    ]),
    group("ink-convert", [
      btn("ink-shape", "ink-to-shape", { size: "large" }),
      col([grid([btn("equation", "ink-to-math"), btn("sync", "replay-ink")])]),
    ]),
  ]);

const designTab = (): RibbonTab =>
  tabNode("design", [
    group(
      "document-formatting",
      [
        // The paint-brush glyph doubles here: a style set is Word's "apply a
        // formatting theme to the document" action.
        split("format-painter", "style-set", parsedItems(styleSetItems()), { size: "large" }),
        btn("theme", "theme", { size: "large" }),
        btn("font-color", "colors", { size: "large" }),
        btn("text-font", "fonts", { size: "large" }),
        btn("text-effects", "effects", { size: "large" }),
        col([
          grid([
            menu("line-spacing", "paragraph-spacing", parsedItems(paragraphSpacingItems())),
            btn("page-border", "set-default"),
          ]),
        ]),
      ],
      "themes-dialog",
    ),
    group("page-background", [
      split("watermark", "watermark", parsedItems(watermarkItems()), { size: "large" }),
      {
        type: "color-picker",
        icon: "page-color",
        event: "page-color",
        label: cmd("page-color"),
        defaultColor: "FFFFFF",
        size: "large",
      },
      split("page-border", "page-border", parsedItems(pageBorderItems()), { size: "large" }),
    ]),
  ]);

const layoutTab = (): RibbonTab =>
  tabNode("layout", [
    group(
      "page-setup",
      [
        split("page-color", "margins", parsedItems(marginsItems()), { size: "large" }),
        split("orientation", "orientation", parsedItems(orientationItems()), { size: "large" }),
        split("page-color", "page-size", parsedItems(sizePaperItems()), { size: "large" }),
        split("multilevel", "columns", parsedItems(columnsItems()), { size: "large" }),
        split("page-break", "page-break", parsedItems(breaksItems()), {
          size: "large",
          label: cmd("breaks"),
        }),
        split("number-symbol", "line-numbers", parsedItems(lineNumbersItems()), { size: "large" }),
        split("hyphenation", "hyphenation", parsedItems(hyphenationItems()), {
          size: "large",
          label: cmd("hyphenation"),
        }),
      ],
      "page-setup-dialog",
    ),
    group(
      "paragraph",
      [
        split("indent-increase", "indent-increase", parsedItems(indentItems()), { size: "large" }),
        split("line-spacing", "line-spacing", parsedItems(spacingItems()), { size: "large" }),
      ],
      "paragraph-dialog",
    ),
    group("arrange", [
      col([
        row([
          split("orientation", "position", parsedItems(positionItems())),
          split("wrap", "wrap", parsedItems(wrapItems())),
        ]),
        row([btn("orientation", "bring-forward"), btn("orientation", "send-backward")]),
      ]),
      split("align-left", "align-objects", parsedItems(alignObjectsItems()), { size: "large" }),
      split("group-objects", "group", parsedItems(groupItems()), { size: "large" }),
      split("rotate", "rotate", parsedItems(rotateItems()), { size: "large" }),
    ]),
  ]);

const referencesTab = (): RibbonTab =>
  tabNode("references", [
    group("toc", [
      split("toc", "toc", parsedItems(tocItems()), { size: "large" }),
      split("multilevel", "add-text", parsedItems(addTextItems()), { size: "large" }),
      split("sync", "update-toc", parsedItems(updateTocItems()), { size: "large" }),
    ]),
    group("footnotes", [
      split("footnote", "insert-footnote", parsedItems(footnoteItems()), { size: "large" }),
    ]),
    group("citations", [
      btn("comment-add", "insert-citation", { size: "large" }),
      btn("people", "manage-sources", { size: "large" }),
      btn("document-print", "bibliography", { size: "large" }),
    ]),
    group("captions", [
      btn("comment-add", "insert-caption", { size: "large" }),
      btn("document-print", "table-of-figures", { size: "large" }),
      btn("sync", "update-figures", { size: "large" }),
      btn("link", "cross-reference", { size: "large" }),
    ]),
    group("index", [
      btn("comment-add", "mark-entry", { size: "large" }),
      btn("document-print", "insert-index", { size: "large" }),
      btn("sync", "update-index", { size: "large" }),
    ]),
    group("toa", [
      btn("comment-add", "mark-citation", { size: "large" }),
      btn("document-print", "insert-toa", { size: "large" }),
    ]),
  ]);

const mailingsTab = (): RibbonTab =>
  tabNode("mailings", [
    group("create", [
      btn("mail", "envelopes", { size: "large" }),
      btn("mail", "labels", { size: "large" }),
    ]),
    group("start-merge", [
      split("document-print", "start-merge", parsedItems(startMergeItems()), { size: "large" }),
      btn("people", "select-recipients", { size: "large" }),
      btn("edit", "edit-recipients", { size: "large" }),
    ]),
    group("write-fields", [
      btn("document-print", "address-block", { size: "large" }),
      btn("comment-add", "greeting-line", { size: "large" }),
      btn("link", "merge-field", { size: "large" }),
      btn("highlight", "highlight-merge", { size: "large" }),
    ]),
    group("preview", [
      btn("search", "preview-results", { size: "large" }),
      col([grid([btn("align-left", "first-record"), btn("align-right", "last-record")])]),
    ]),
    group("finish", [
      split("document-print", "finish-merge", parsedItems(finishMergeItems()), { size: "large" }),
    ]),
  ]);

const reviewTab = (): RibbonTab =>
  tabNode("review", [
    group("proofing", [
      btn("spell-check", "spell-check", { size: "large" }),
      col([grid([btn("word-count", "word-count"), btn("search", "thesaurus")])]),
    ]),
    group("accessibility", [btn("checkmark-circle", "check-accessibility", { size: "large" })]),
    group("language", [
      btn("link", "translate", { size: "large" }),
      btn("text-font", "language", { size: "large" }),
    ]),
    group("comments", [
      btn("comment-add", "new-comment", { size: "large" }),
      col([grid([btn("edit", "edit-comment"), btn("close", "delete-comment")])]),
      col([grid([btn("align-left", "previous-comment"), btn("align-right", "next-comment")])]),
      btn("comment", "show-comments", { size: "large" }),
    ]),
    group("tracking", [
      btn("group-objects", "track-changes", { size: "large" }),
      split("accept", "accept-change", parsedItems(acceptItems()), { size: "large" }),
      split("close", "reject-change", parsedItems(rejectItems()), { size: "large" }),
      col([grid([btn("align-left", "previous-change"), btn("align-right", "next-change")])]),
      btn("reviewing-pane", "reviewing-pane", { size: "large" }),
    ]),
    group("compare", [
      btn("group-objects", "compare", { size: "large" }),
      btn("group-objects", "combine", { size: "large" }),
    ]),
    group("protect", [
      btn("protect", "restrict-editing", { size: "large" }),
      btn("protect", "protect-document", { size: "large" }),
    ]),
  ]);

const viewTab = (): RibbonTab =>
  tabNode("view", [
    group("views", [
      btn("document-print", "read-mode", { size: "large" }),
      btn("print", "print-layout", { size: "large" }),
      btn("document-print", "web-layout", { size: "large" }),
      btn("group-objects", "outline", { size: "large" }),
      btn("document-print", "draft", { size: "large" }),
    ]),
    group("show", [
      col([
        btn("ruler", "toggle-ruler"),
        btn("gridlines", "toggle-gridlines"),
        btn("replace", "toggle-navigation"),
      ]),
    ]),
    group("zoom", [
      btn("zoom-in", "zoom", { size: "large" }),
      split("zoom-in", "zoom-100", parsedItems(zoomItems()), { size: "large" }),
    ]),
    group("window", [
      btn("grid", "new-window", { size: "large" }),
      col([grid([btn("grid", "arrange-all"), btn("group-objects", "split-window")])]),
    ]),
    group("macros", [
      btn("group-objects", "view-macros", { size: "large" }),
      btn("edit", "record-macro", { size: "large" }),
    ]),
  ]);

// --- Contextual tabs (Word's Table Tools) ------------------------------------
// Values match the commands.ts value spaces: table-style presets, the
// table-borders sides (same as the Home border menu), and the align-cell
// 9-grid keys (top/middle/bottom × left/center/right).

// --- Table Design gallery -----------------------------------------------------
// Word renders each gallery entry as a mini-table thumbnail painted from the
// style's borders and conditional fills. We generate that thumbnail as an SVG
// from the same preset data the command stamps — one source, preview and
// result can't drift.

function tableStylePreviewSvg(preset: TableStylePreset): string {
  const stroke = `stroke="#595959" stroke-width="1"`;
  const cell = 10;
  const x3 = 1 + cell * 3;
  const y3 = 1 + cell * 3;
  const parts: string[] = [];
  if (preset.headerFill) {
    parts.push(
      `<rect x="1" y="1" width="${cell * 3}" height="${cell}" fill="#${preset.headerFill}"/>`,
    );
  }
  if (preset.bandFill) {
    parts.push(
      `<rect x="1" y="${1 + cell * 2}" width="${cell * 3}" height="${cell}" fill="#${preset.bandFill}"/>`,
    );
  }
  const b = preset.borders ?? {};
  const on = (side: { style: string } | undefined): boolean => !!side && side.style !== "none";
  if (on(b.top)) parts.push(`<line x1="1" y1="1" x2="${x3}" y2="1" ${stroke}/>`);
  if (on(b.bottom)) parts.push(`<line x1="1" y1="${y3}" x2="${x3}" y2="${y3}" ${stroke}/>`);
  if (on(b.left)) parts.push(`<line x1="1" y1="1" x2="1" y2="${y3}" ${stroke}/>`);
  if (on(b.right)) parts.push(`<line x1="${x3}" y1="1" x2="${x3}" y2="${y3}" ${stroke}/>`);
  if (on(b.insideHorizontal)) {
    for (const i of [1, 2]) {
      parts.push(`<line x1="1" y1="${1 + cell * i}" x2="${x3}" y2="${1 + cell * i}" ${stroke}/>`);
    }
  }
  if (on(b.insideVertical)) {
    for (const i of [1, 2]) {
      parts.push(`<line x1="${1 + cell * i}" y1="1" x2="${1 + cell * i}" y2="${y3}" ${stroke}/>`);
    }
  }
  // A borderless preset still shows a faint dashed cell grid, like Word's.
  if (parts.length === 0) {
    parts.push(
      `<rect x="1" y="1" width="${cell * 3}" height="${cell * 3}" fill="none" stroke="#C8C8C8" stroke-dasharray="2,2"/>`,
    );
  }
  return `<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">${parts.join("")}</svg>`;
}

// The tblLook flags behind Word's Table Style Options checkboxes, in Word's
// order, with their checkbox label keys.
const TABLE_LOOK_OPTIONS: readonly { flag: string; label: string }[] = [
  { flag: "firstRow", label: opt("look-header-row") },
  { flag: "lastRow", label: opt("look-total-row") },
  { flag: "bandRow", label: opt("look-banded-rows") },
  { flag: "firstCol", label: opt("look-first-column") },
  { flag: "lastCol", label: opt("look-last-column") },
  { flag: "bandCol", label: opt("look-banded-columns") },
];

let tableStyleIconsRegistered = false;
function ensureTableStyleIcons(): void {
  if (tableStyleIconsRegistered) return;
  for (const [id, preset] of Object.entries(TABLE_STYLE_PRESETS)) {
    registerIcon(`table-style-${id}`, tableStylePreviewSvg(preset));
  }
  tableStyleIconsRegistered = true;
}

/** Word's Table Styles gallery: the presets as icon-over-label thumbnails in a
 *  strip (first four visible), the More bar expanding every preset in the
 *  same compound shape as a drop-down grid. */
const tableStyleGallery = (): RibbonGallery => ({
  type: "gallery",
  event: "table-style",
  label: opt("more-table-styles"),
  items: Object.keys(TABLE_STYLE_PRESETS).map((id) => ({
    icon: `table-style-${id}`,
    text: opt(`style-${id}`),
    value: id,
  })),
  visibleCount: 4,
});

const cellAlignItems = (): string =>
  JSON.stringify(
    ["tl", "tc", "tr", "ml", "mc", "mr", "bl", "bc", "br"].map((value) => ({
      text: opt(`cell-align-${value}`),
      value,
    })),
  );

const tableSelectItems = (): string =>
  JSON.stringify([
    { text: opt("select-table-cell"), value: "cell", event: "select-table-cell" },
    { text: opt("select-table-column"), value: "column", event: "select-table-column" },
    { text: opt("select-table-row"), value: "row", event: "select-table-row" },
    { text: opt("select-table"), value: "table" },
  ]);

const tableDeleteItems = (): string =>
  JSON.stringify([
    // Word opens the Delete Cells dialog — needs its own dialog (not built).
    { text: opt("delete-cells"), value: "cells", event: "delete-table", disabled: true },
    { text: opt("delete-columns"), value: "columns", event: "delete-column" },
    { text: opt("delete-rows"), value: "rows", event: "delete-row" },
    { text: opt("delete-table"), value: "table" },
  ]);

/** The AutoFit split's drop-down: Word's three AutoFit modes. */
const autofitItems = (): string =>
  JSON.stringify([
    { text: opt("autofit-contents"), value: "contents", event: "autofit-contents" },
    { text: opt("autofit-window"), value: "window", event: "autofit-window" },
    { text: opt("fixed-column-width"), value: "fixed", event: "fixed-column-width" },
  ]);

/** The locale's Cell Size unit system (Word zh shows cm, en shows inches) —
 *  resolved against the ribbon's i18n scope (the workspace carries the
 *  effective `<docen-document lang>`), shared with the host's live combo
 *  sync (#syncCellSize). */
export const useCmUnits = (scope?: Element): boolean =>
  resolveLang(scope ?? document.documentElement)
    .toLowerCase()
    .startsWith("zh");

/** A length in twips as the Cell Size string ("1.50 厘米" / '1.00"'). */
export function formatMeasureTwip(tw: number, scope?: Element): string {
  return useCmUnits(scope) ? `${(tw / 567).toFixed(2)} 厘米` : `${(tw / 1440).toFixed(2)}"`;
}

/** Word's Cell Size spinner presets — the values are the twips the commands
 *  receive; the labels follow the locale's unit system. Free-typed entries
 *  accept the same UniversalMeasure units ("1.5cm" / "0.5in"). */
const cellWidthPresets = (scope?: Element): readonly (readonly [string, string])[] =>
  useCmUnits(scope)
    ? [
        ["0.5 厘米", "283"],
        ["1 厘米", "567"],
        ["1.5 厘米", "850"],
        ["2 厘米", "1134"],
        ["3 厘米", "1701"],
        ["4 厘米", "2268"],
      ]
    : [
        ['0.5"', "720"],
        ['0.75"', "1080"],
        ['1"', "1440"],
        ['1.5"', "2160"],
        ['2"', "2880"],
        ['3"', "4320"],
      ];
const cellHeightPresets = (scope?: Element): readonly (readonly [string, string])[] => [
  ...(useCmUnits(scope)
    ? ([
        ["0.5 厘米", "283"],
        ["1 厘米", "567"],
        ["1.5 厘米", "850"],
        ["2 厘米", "1134"],
      ] as const)
    : ([
        ['0.25"', "360"],
        ['0.5"', "720"],
        ['0.75"', "1080"],
        ['1"', "1440"],
      ] as const)),
  [useCmUnits(scope) ? "自动" : "auto", "0"],
];
const cellWidthItems = (scope?: Element): string =>
  JSON.stringify(cellWidthPresets(scope).map(([text, value]) => ({ text, value })));
const cellHeightItems = (scope?: Element): string =>
  JSON.stringify(cellHeightPresets(scope).map(([text, value]) => ({ text, value })));

/** Word's contextual Table Tools — the Table Design / Table Layout tabs that
 *  appear while the caret is inside a table. Marked `contextual` so
 *  {@link ribbonTabs} excludes them from the static render; the host appends
 *  them (via {@link buildContextualTab}) when the selection enters a table and
 *  removes them when it leaves. `scope` is the i18n scope the unit-system
 *  presets (Cell Size) resolve against. */
export function tableContextTabs(scope?: Element): RibbonTab[] {
  ensureTableStyleIcons();
  return [
    {
      id: "table-design",
      label: tab("table-design"),
      contextual: true,
      groups: [
        // Word's two leading groups, in Word's order: the Table Style Options
        // checkboxes (2×3), then the Table Styles gallery.
        group("table-style-options", [
          grid(
            TABLE_LOOK_OPTIONS.map((o) => ({
              type: "checkbox" as const,
              event: "toggle-table-look",
              value: o.flag,
              label: o.label,
            })),
          ),
        ]),
        group("table-styles", [tableStyleGallery()]),
        group("table-shading", [
          {
            type: "color-picker",
            icon: "shading",
            event: "cell-shading",
            label: cmd("cell-shading"),
            defaultColor: "FFFF00",
            size: "large",
          },
        ]),
        group("table-borders", [
          split("border", "table-borders", parsedItems(borderItems()), { size: "large" }),
        ]),
        group("draw-border", [
          // Word's Draw Border tools — the canvas has no border-painting
          // interaction yet, so the group greys until then.
          btn("pen", "pen-style", { size: "large" }),
          btn("font-size", "pen-size", { size: "large" }),
          btn("font-color", "pen-color", { size: "large" }),
          btn("format-painter", "border-painter", { size: "large" }),
          btn("gridlines", "toggle-gridlines", { size: "large" }),
        ]),
      ],
    },
    {
      id: "table-layout",
      label: tab("table-layout"),
      contextual: true,
      groups: [
        group("table", [
          // Table Properties opens Word's table dialog (read arrives via
          // #onCommand, the commit via table-properties:ok — same pair as the
          // context menu's entry).
          btn("table-properties", "table-properties", { size: "large" }),
          // Word stacks Select over Delete — two small menus, one column.
          col([
            menu("table-cursor", "select-table", parsedItems(tableSelectItems())),
            menu("table-delete", "delete-table", parsedItems(tableDeleteItems())),
          ]),
        ]),
        group("rows-columns", [
          row([
            col([
              btn("table-stack-above", "insert-row-above"),
              btn("table-stack-below", "insert-row-below"),
            ]),
            col([
              btn("table-stack-left", "insert-column-left"),
              btn("table-stack-right", "insert-column-right"),
            ]),
          ]),
        ]),
        group("merge", [
          btn("merge-cells", "merge-cells", { size: "large" }),
          btn("split-cells", "split-cell", { size: "large" }),
          btn("table-simple", "split-table", { size: "large" }),
        ]),
        group("cell-size", [
          split("autofit", "autofit", parsedItems(autofitItems()), { size: "large" }),
          col([
            combo("cell-height", "auto", parsedItems(cellHeightItems(scope)), {
              comboboxSize: "short",
            }),
            combo("cell-width", '1"', parsedItems(cellWidthItems(scope)), {
              comboboxSize: "short",
            }),
          ]),
          col([
            grid([
              btn("distribute-rows", "distribute-rows"),
              btn("distribute-columns", "distribute-columns"),
            ]),
          ]),
        ]),
        group("alignment", [
          split("align-center", "align-cell", parsedItems(cellAlignItems()), { size: "large" }),
          btn("text-direction", "text-direction", { size: "large" }),
          // Word's Cell Margins menu button: the presets stamp the caret
          // cell's tcMar; Custom opens the options dialog (not built).
          menu("table-properties", "cell-margins", parsedItems(cellMarginItems()), {
            size: "large",
          }),
        ]),
        group("data", [
          btn("table-repeat-headers", "repeat-header-rows", { size: "large" }),
          btn("list", "convert-to-text", { size: "large" }),
        ]),
      ],
    },
  ];
}

/** The math symbols of the equation context tab's Symbols grid — Word's Basic
 *  Math panel's frequent picks (operators, Greek letters, arrows). */
const EQUATION_SYMBOLS = "± ∞ ≠ ≈ ≤ ≥ ∝ × ÷ ∑ ∫ √ ∂ π ∇ α β γ δ ε θ λ σ μ ω Δ Ω → ⇒ ∈ ∀ ∃".split(
  " ",
);

/** Word's Equation Tools — the contextual tab while the selection sits on a
 *  math atom. Structures replays the Insert tab's equation templates (the
 *  engine's five OMML shapes plus the blank Alt+= insert), Symbols is a grid
 *  of characters inserted at the caret. Marked `contextual` so
 *  {@link ribbonTabs} excludes it from the static render; the host appends it
 *  via {@link buildContextualTab} as the selection enters/leaves a math atom. */
export function equationContextTab(): RibbonTab {
  return {
    id: "equation",
    label: tab("equation"),
    contextual: true,
    groups: [
      group("structures", [
        grid([
          { type: "button", label: opt("equation-blank"), event: "equation", value: "plain" },
          {
            type: "button",
            label: opt("equation-fraction"),
            event: "equation",
            value: "fraction",
          },
          {
            type: "button",
            label: opt("equation-script"),
            event: "equation",
            value: "superScript",
          },
          {
            type: "button",
            label: opt("equation-radical"),
            event: "equation",
            value: "radical",
          },
          { type: "button", label: opt("equation-sum"), event: "equation", value: "sum" },
          {
            type: "button",
            label: opt("equation-integral"),
            event: "equation",
            value: "integral",
          },
        ]),
      ]),
      group("symbols", [
        grid(
          EQUATION_SYMBOLS.map((char) => ({
            type: "button" as const,
            label: char,
            event: "insert-symbol",
            value: char,
          })),
        ),
      ]),
    ],
  };
}

/** Word's Header & Footer Tools — the contextual tab while a header/footer
 *  story is being edited. Marked `contextual` so {@link ribbonTabs} excludes
 *  it from the static render; the host appends it when a story opens and
 *  removes it when the story closes (same lifecycle as
 *  {@link tableContextTabs}). The Options checkboxes mirror the Insert tab's
 *  header/footer drop-down flags (first-page-different / odd-even-different). */
export function headerFooterContextTab(): RibbonTab {
  return {
    id: "header-footer-tab",
    label: tab("header-footer-tools"),
    contextual: true,
    groups: [
      group("navigation", [
        btn("header", "goto-header", { size: "large" }),
        btn("footer", "goto-footer", { size: "large" }),
      ]),
      group("options", [
        col([
          {
            type: "checkbox",
            event: "header-option",
            value: "title-page",
            label: opt("different-first"),
          },
          {
            type: "checkbox",
            event: "header-option",
            value: "odd-even",
            label: opt("odd-even"),
          },
        ]),
      ]),
      group("close", [btn("close", "close-header-footer", { size: "large" })]),
    ],
  };
}
