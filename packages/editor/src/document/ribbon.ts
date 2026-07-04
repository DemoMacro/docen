import { quickStyles, type StylesOptions } from "@docen/docx";

import type {
  RibbonButton,
  RibbonColorPicker,
  RibbonCombobox,
  RibbonControl,
  RibbonControlOrLayout,
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
 * `ribbonTabs()` builds the RibbonTab schema (resolving every visible string
 * through `t("ribbon.*")`, so it re-renders in the active locale);
 * `renderRibbonFromSchema()` consumes that tree to build the ribbon DOM. The
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
import { resolveLang, t } from "../ui";

// --- i18n shortcuts (ribbon.* keys, resolved at call time) -------------------

const tab = (id: string): string => t(`ribbon.tab.${id}`);
const grp = (id: string): string => t(`ribbon.group.${id}`);
const cmd = (event: string): string => t(`ribbon.cmd.${event}`);
const opt = (value: string): string => t(`ribbon.opt.${value}`);

// --- Option sets (menu/combobox items) ---------------------------------------
// Font names and point sizes are data, not UI copy — kept untranslated.

// Fallback font list shown when the Local Font Access API is unavailable or
// denied — includes common CJK faces so a zh host still sees familiar names.
const FONT_NAMES = [
  "Microsoft YaHei",
  "Calibri",
  "Arial",
  "Cambria",
  "Times New Roman",
  "Georgia",
  "Verdana",
  "Tahoma",
  "Courier New",
  "Segoe UI",
  "宋体",
  "黑体",
  "楷体",
  "仿宋",
  "等线",
];

// Chinese size names (Word's zh font-size names) mapped to point values, largest first
// (matching the Word zh size picker order).
const FONT_SIZES_CN: ReadonlyArray<readonly [string, number]> = [
  ["初号", 42],
  ["小初", 36],
  ["一号", 26],
  ["小一", 24],
  ["二号", 22],
  ["小二", 18],
  ["三号", 16],
  ["小三", 15],
  ["四号", 14],
  ["小四", 12],
  ["五号", 10.5],
  ["小五", 9],
  ["六号", 7.5],
  ["小六", 6.5],
  ["七号", 5.5],
  ["八号", 5],
];
// Point sizes listed below the Chinese sizes (ascending — Word's zh size
// picker orders the numeric sizes small-to-large under the Chinese names).
const FONT_SIZES_PT: ReadonlyArray<number> = [
  5, 5.5, 6.5, 7.5, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72,
];

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
 *  id, which round-trips via the Paragraph/Heading `styleId` attr. */
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

const highlightItems = (): string =>
  JSON.stringify([
    { text: opt("no-color"), value: "none" },
    { text: opt("yellow"), value: "yellow" },
    { text: opt("bright-green"), value: "bright-green" },
    { text: opt("turquoise"), value: "turquoise" },
    { text: opt("pink"), value: "pink" },
    { text: opt("red"), value: "red" },
    { text: opt("green"), value: "green" },
    { text: opt("blue"), value: "blue" },
  ]);

const caseItems = (): string =>
  JSON.stringify([
    { text: opt("sentence-case"), value: "sentence" },
    { text: opt("lowercase"), value: "lower" },
    { text: opt("uppercase"), value: "upper" },
    { text: opt("capitalize"), value: "capitalize" },
    { text: opt("toggle-case"), value: "toggle" },
  ]);

const bulletItems = (): string =>
  JSON.stringify([
    { text: opt("bullet"), value: "bullet" },
    { text: opt("circle"), value: "circle" },
    { text: opt("square"), value: "square" },
    { text: opt("change-list-level"), value: "level" },
  ]);

const numberItems = (): string =>
  JSON.stringify([
    { text: opt("decimal"), value: "decimal" },
    { text: opt("lower-alpha"), value: "lower-alpha" },
    { text: opt("lower-roman"), value: "lower-roman" },
    { text: opt("change-list-level"), value: "level" },
  ]);

const multilevelItems = (): string =>
  JSON.stringify([
    { text: opt("level-1"), value: "level-1" },
    { text: opt("level-2"), value: "level-2" },
    { text: opt("level-3"), value: "level-3" },
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
  ]);

const findItems = (): string =>
  JSON.stringify([
    { text: opt("find"), value: "find" },
    { text: opt("go-to"), value: "go-to" },
  ]);

const selectItems = (): string =>
  JSON.stringify([
    { text: opt("select-all"), value: "all" },
    { text: opt("select-objects"), value: "objects" },
    { text: opt("select-similar"), value: "similar" },
  ]);

const coverItems = (): string =>
  JSON.stringify([
    { text: cmd("page-break"), value: "page-break", event: "page-break" },
    { text: cmd("section-break"), value: "section-break", event: "section-break" },
  ]);

const tableItems = (): string =>
  JSON.stringify([
    { text: opt("insert-table"), value: "insert" },
    { text: opt("draw-table"), value: "draw" },
    { text: opt("convert-text"), value: "convert" },
    { text: opt("excel"), value: "excel" },
    { text: opt("quick-tables"), value: "quick" },
  ]);

const marginsItems = (): string =>
  JSON.stringify([
    { text: opt("normal-margin"), value: "normal" },
    { text: opt("narrow"), value: "narrow" },
    { text: opt("moderate"), value: "moderate" },
    { text: opt("wide"), value: "wide" },
    { text: opt("custom-margins"), value: "custom" },
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
    { text: opt("more-sizes"), value: "more" },
  ]);

const columnsItems = (): string =>
  JSON.stringify([
    { text: opt("one-col"), value: "1" },
    { text: opt("two-col"), value: "2" },
    { text: opt("three-col"), value: "3" },
    { text: opt("more-columns"), value: "more" },
  ]);

const breaksItems = (): string =>
  JSON.stringify([
    { text: cmd("page-break"), value: "page-break", event: "page-break" },
    { text: opt("column-break"), value: "column-break", event: "column-break" },
    { text: opt("text-wrapping"), value: "text-wrapping", event: "text-wrapping" },
    { text: cmd("section-break"), value: "section-break", event: "section-break" },
  ]);

const indentItems = (): string =>
  JSON.stringify([
    { text: opt("increase-indent"), value: "increase" },
    { text: opt("decrease-indent"), value: "decrease" },
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

const alignItems = (): string =>
  JSON.stringify([
    { text: cmd("align-left"), value: "left" },
    { text: opt("align-center"), value: "center" },
    { text: cmd("align-right"), value: "right" },
  ]);

const footnoteItems = (): string =>
  JSON.stringify([
    { text: cmd("insert-footnote"), value: "footnote" },
    { text: opt("endnote"), value: "endnote" },
    { text: opt("next-footnote"), value: "next" },
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
): DocumentFragment {
  const frag = document.createDocumentFragment();

  const tablist = document.createElement("fluent-tablist");
  tablist.setAttribute("slot", "tabs");
  tablist.setAttribute("appearance", "transparent");
  const activeId = tabs[0]?.id ?? "";
  if (activeId) tablist.setAttribute("activeid", activeId);
  for (const tab of tabs) {
    const t = document.createElement("docen-ribbon-tab");
    t.setAttribute("slot", "tab");
    t.id = tab.id;
    t.textContent = tab.label;
    tablist.append(t);
  }
  frag.append(tablist);

  for (const tab of tabs) {
    const panel = document.createElement("docen-ribbon-panel");
    panel.setAttribute("value", tab.id);
    for (const g of tab.groups) panel.append(buildGroup(g));
    frag.append(panel);
  }

  for (const c of actions) {
    const el = buildControl(c);
    el.setAttribute("slot", "actions");
    frag.append(el);
  }
  return frag;
}

function buildGroup(g: RibbonGroup): HTMLElement {
  const el = document.createElement("docen-ribbon-group");
  el.setAttribute("label", g.label);
  if (g.launcher) el.setAttribute("launcher", g.launcher);
  for (const c of g.controls) el.append(buildControlOrLayout(c));
  return el;
}

function buildControlOrLayout(c: RibbonControlOrLayout): HTMLElement {
  return c.type === "layout" ? buildLayout(c) : buildControl(c);
}

function buildLayout(l: RibbonLayout): HTMLElement {
  const el = document.createElement("div");
  el.className = l.layout === "column" ? "rb-col" : l.layout === "row" ? "rb-row" : "rb-grid";
  for (const c of l.controls) el.append(buildControlOrLayout(c));
  return el;
}

/** Stamp the shared base attrs (icon/label/event/iconOnly/size/disabled) every
 *  control component reads. */
function applyBase(
  el: HTMLElement,
  c: {
    icon?: string;
    label?: string;
    event?: string;
    iconOnly?: boolean;
    size?: "small" | "large";
    disabled?: boolean;
  },
): void {
  if (c.icon) el.setAttribute("icon", c.icon);
  if (c.label) el.setAttribute("label", c.label);
  if (c.event) el.setAttribute("event", c.event);
  if (c.iconOnly) el.setAttribute("icon-only", "");
  if (c.size === "large") el.setAttribute("size", "large");
  if (c.disabled) el.setAttribute("disabled", "");
}

function buildControl(c: RibbonControl): HTMLElement {
  switch (c.type) {
    case "separator": {
      const el = document.createElement("span");
      el.className = "rb-vsep";
      return el;
    }
    case "button": {
      const el = document.createElement("docen-ribbon-button");
      applyBase(el, c);
      return el;
    }
    case "menu": {
      const el = document.createElement("docen-ribbon-menu");
      applyBase(el, c);
      el.setAttribute("items", JSON.stringify(c.items ?? []));
      return el;
    }
    case "split": {
      const el = document.createElement("docen-ribbon-split-button");
      applyBase(el, c);
      el.setAttribute("items", JSON.stringify(c.items ?? []));
      return el;
    }
    case "combobox": {
      const el = document.createElement("docen-ribbon-combobox");
      applyBase(el, c);
      if (c.value != null) el.setAttribute("value", c.value);
      el.setAttribute("items", JSON.stringify(c.items ?? []));
      if (c.source) el.setAttribute("source", c.source);
      if (c.comboboxSize === "short") el.setAttribute("size", "short");
      return el;
    }
    case "color-picker": {
      const el = document.createElement("docen-color-picker");
      applyBase(el, c);
      if (c.defaultColor) el.setAttribute("default-color", c.defaultColor);
      return el;
    }
  }
}

// --- Data-driven ribbon (RibbonTab tree) -------------------------------------
// The 9 tabs expressed as data; renderRibbonFromSchema consumes this tree to
// build the ribbon DOM. i18n (t("ribbon.*")) resolves at call time, so
// re-calling on a locale change re-localizes the labels.

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
  label?: string,
): RibbonMenu => ({
  type: "menu",
  icon,
  event,
  label: label ?? cmd(event),
  items,
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

const picker = (icon: string, event: string, defaultColor: string): RibbonColorPicker => ({
  type: "color-picker",
  icon,
  event,
  label: cmd(event),
  defaultColor,
  iconOnly: true,
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
    menu("edit", "edit-mode", parsedItems(editItems()), cmd("editing")),
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
            btn("underline", "underline", { iconOnly: true }),
            btn("strike", "strike", { iconOnly: true }),
            btn("superscript", "superscript", { iconOnly: true }),
            btn("subscript", "subscript", { iconOnly: true }),
            sep(),
            split("highlight", "highlight", parsedItems(highlightItems()), { iconOnly: true }),
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
            btn("sort", "sort", { iconOnly: true }),
            btn("show-marks", "show-marks", { iconOnly: true }),
          ]),
          row([
            btn("align-left", "align-left", { iconOnly: true }),
            btn("align-center", "align-center", { iconOnly: true }),
            btn("align-right", "align-right", { iconOnly: true }),
            btn("justify", "justify", { iconOnly: true }),
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
      btn("shapes", "shapes", { size: "large" }),
      btn("icon-library", "icons", { size: "large" }),
      btn("3d-model", "3d-model", { size: "large" }),
      btn("smartart", "smartart", { size: "large" }),
      btn("chart", "chart", { size: "large" }),
      btn("insert-picture", "screenshot", { size: "large" }),
    ]),
    group("links", [
      btn("hyperlink", "link", { size: "large" }),
      btn("bookmark", "bookmark", { size: "large" }),
      btn("comment-add", "comment", { size: "large" }),
    ]),
    group("header-footer", [
      btn("header", "header", { size: "large" }),
      btn("footer", "footer", { size: "large" }),
      btn("page-number", "page-number", { size: "large" }),
    ]),
    group("text", [
      btn("text-box", "text-box", { size: "large" }),
      btn("wordart", "wordart", { size: "large" }),
    ]),
    group("symbols", [
      btn("equation", "equation", { size: "large" }),
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
        btn("theme", "theme", { size: "large" }),
        btn("font-color", "colors", { size: "large" }),
        btn("text-font", "fonts", { size: "large" }),
        btn("text-effects", "effects", { size: "large" }),
        col([grid([btn("line-spacing", "paragraph-spacing"), btn("page-border", "set-default")])]),
      ],
      "themes-dialog",
    ),
    group("page-background", [
      btn("page-color", "watermark", { size: "large" }),
      btn("page-color", "page-color", { size: "large" }),
      btn("page-border", "page-border", { size: "large" }),
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
        btn("number-symbol", "line-numbers", { size: "large" }),
      ],
      "page-setup-dialog",
    ),
    group(
      "paragraph",
      [
        split("indent-increase", "indent-left", parsedItems(indentItems()), { size: "large" }),
        split("line-spacing", "spacing", parsedItems(spacingItems()), { size: "large" }),
      ],
      "paragraph-dialog",
    ),
    group("arrange", [
      col([
        row([btn("orientation", "position"), btn("wrap", "wrap")]),
        row([btn("orientation", "bring-forward"), btn("orientation", "send-backward")]),
      ]),
      split("align-left", "align", parsedItems(alignItems()), { size: "large" }),
      split("group-objects", "group", parsedItems(groupItems()), { size: "large" }),
      split("rotate", "rotate", parsedItems(rotateItems()), { size: "large" }),
    ]),
  ]);

const referencesTab = (): RibbonTab =>
  tabNode("references", [
    group("toc", [
      btn("toc", "toc", { size: "large" }),
      btn("multilevel", "add-text", { size: "large" }),
      btn("sync", "update-toc", { size: "large" }),
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
    ]),
    group("tracking", [btn("group-objects", "track-changes", { size: "large" })]),
    group("changes", [
      btn("accept", "accept-change", { size: "large" }),
      btn("close", "reject-change", { size: "large" }),
      col([grid([btn("align-left", "previous-change"), btn("align-right", "next-change")])]),
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
      btn("print", "print-layout", { size: "large" }),
      btn("document-print", "web-layout", { size: "large" }),
      btn("document-print", "read-mode", { size: "large" }),
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
