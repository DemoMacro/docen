import { quickStyles, type StylesOptions } from "@docen/docx";

/**
 * Default MS Office Word ribbon for `<docen-document>` — all 8 standard tabs
 * (Home/Insert/Design/Layout/References/Mailings/Review/View) with the canonical
 * groups and primary commands.
 *
 * Built by `buildRibbonInnerHTML()`, which resolves every visible string through
 * `t("ribbon.*")` (see `./i18n.ts`), so the ribbon re-renders in the active
 * locale. The host `<docen-document>` stamps the returned markup into its
 * `<docen-ribbon>` element and re-runs it on language change. Callers wanting a
 * tailored ribbon compose the Fluent shell components directly instead.
 *
 * Layout helpers (`.rb-col` / `.rb-row` / `.rb-vsep`) are injected by the host
 * style — Office groups stack a large button beside rows/columns of small
 * `icon-only` buttons.
 *
 * Each command carries an `event` name. `commands.ts` wires the ones the Tiptap
 * engine supports today (marks, lists, alignment, styles, breaks, history); the
 * rest render as a complete visual skeleton and no-op on click until wired.
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

/**
 * Build the ribbon's inner markup (tabs + panels, without the wrapping
 * `<docen-ribbon>`) for the currently active locale. Re-call on language change.
 */
export function buildRibbonInnerHTML(styles?: StylesOptions | null): string {
  return `
      <fluent-tablist slot="tabs" appearance="transparent" activeid="${DEFAULT_RIBBON_TAB}">
        <docen-ribbon-tab slot="tab" id="home">${tab("home")}</docen-ribbon-tab>
        <docen-ribbon-tab slot="tab" id="insert">${tab("insert")}</docen-ribbon-tab>
        <docen-ribbon-tab slot="tab" id="draw">${tab("draw")}</docen-ribbon-tab>
        <docen-ribbon-tab slot="tab" id="design">${tab("design")}</docen-ribbon-tab>
        <docen-ribbon-tab slot="tab" id="layout">${tab("layout")}</docen-ribbon-tab>
        <docen-ribbon-tab slot="tab" id="references">${tab("references")}</docen-ribbon-tab>
        <docen-ribbon-tab slot="tab" id="mailings">${tab("mailings")}</docen-ribbon-tab>
        <docen-ribbon-tab slot="tab" id="review">${tab("review")}</docen-ribbon-tab>
        <docen-ribbon-tab slot="tab" id="view">${tab("view")}</docen-ribbon-tab>
      </fluent-tablist>
      ${homePanel(styles)}
      ${insertPanel()}
      ${drawPanel()}
      ${designPanel()}
      ${layoutPanel()}
      ${referencesPanel()}
      ${mailingsPanel()}
      ${reviewPanel()}
      ${viewPanel()}`;
}

// --- Home --------------------------------------------------------------------

const homePanel = (styles?: StylesOptions | null): string => `
      <docen-ribbon-panel value="home">
        <docen-ribbon-group label="${grp("clipboard")}" launcher="clipboard-dialog">
          <docen-ribbon-split-button icon="paste" label="${cmd("paste")}" event="paste" size="large" items='${pasteItems()}'></docen-ribbon-split-button>
          <div class="rb-col">
            <docen-ribbon-button icon="cut" label="${cmd("cut")}" event="cut" icon-only></docen-ribbon-button>
            <docen-ribbon-button icon="copy" label="${cmd("copy")}" event="copy" icon-only></docen-ribbon-button>
            <docen-ribbon-button icon="format-painter" label="${cmd("format-painter")}" event="format-painter" icon-only></docen-ribbon-button>
          </div>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("font")}" launcher="font-dialog">
          <div class="rb-col">
            <div class="rb-row">
              <docen-ribbon-combobox value="Microsoft YaHei" event="font-name" source="local-fonts" items='${fontItems()}'></docen-ribbon-combobox>
              <docen-ribbon-combobox value="14" event="font-size" size="short" items='${sizeItems()}'></docen-ribbon-combobox>
              <docen-ribbon-button icon="font-size" label="${cmd("grow-font")}" event="grow-font" icon-only></docen-ribbon-button>
              <docen-ribbon-button icon="font-size" label="${cmd("shrink-font")}" event="shrink-font" icon-only></docen-ribbon-button>
              <docen-ribbon-split-button icon="case" label="${cmd("change-case")}" event="change-case" icon-only items='${caseItems()}'></docen-ribbon-split-button>
              <docen-ribbon-button icon="clear-format" label="${cmd("clear-format")}" event="clear-format" icon-only></docen-ribbon-button>
            </div>
            <div class="rb-row">
              <docen-ribbon-button icon="bold" label="${cmd("bold")}" event="bold" icon-only></docen-ribbon-button>
              <docen-ribbon-button icon="italic" label="${cmd("italic")}" event="italic" icon-only></docen-ribbon-button>
              <docen-ribbon-button icon="underline" label="${cmd("underline")}" event="underline" icon-only></docen-ribbon-button>
              <docen-ribbon-button icon="strike" label="${cmd("strike")}" event="strike" icon-only></docen-ribbon-button>
              <docen-ribbon-button icon="superscript" label="${cmd("superscript")}" event="superscript" icon-only></docen-ribbon-button>
              <docen-ribbon-button icon="subscript" label="${cmd("subscript")}" event="subscript" icon-only></docen-ribbon-button>
              <span class="rb-vsep"></span>
              <docen-ribbon-split-button icon="highlight" label="${cmd("highlight")}" event="highlight" icon-only items='${highlightItems()}'></docen-ribbon-split-button>
              <docen-color-picker icon="font-color" label="${cmd("font-color")}" event="font-color" default-color="000000" icon-only></docen-color-picker>
            </div>
          </div>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("paragraph")}" launcher="paragraph-dialog">
          <div class="rb-col">
            <div class="rb-row">
              <docen-ribbon-split-button icon="list" label="${cmd("bullet-list")}" event="bullet-list" icon-only items='${bulletItems()}'></docen-ribbon-split-button>
              <docen-ribbon-split-button icon="numbering" label="${cmd("ordered-list")}" event="ordered-list" icon-only items='${numberItems()}'></docen-ribbon-split-button>
              <docen-ribbon-split-button icon="multilevel" label="${cmd("multilevel-list")}" event="multilevel-list" icon-only items='${multilevelItems()}'></docen-ribbon-split-button>
              <docen-ribbon-button icon="indent-decrease" label="${cmd("indent-decrease")}" event="indent-decrease" icon-only></docen-ribbon-button>
              <docen-ribbon-button icon="indent-increase" label="${cmd("indent-increase")}" event="indent-increase" icon-only></docen-ribbon-button>
              <docen-ribbon-button icon="sort" label="${cmd("sort")}" event="sort" icon-only></docen-ribbon-button>
              <docen-ribbon-button icon="show-marks" label="${cmd("show-marks")}" event="show-marks" icon-only></docen-ribbon-button>
            </div>
            <div class="rb-row">
              <docen-ribbon-button icon="align-left" label="${cmd("align-left")}" event="align-left" icon-only></docen-ribbon-button>
              <docen-ribbon-button icon="align-center" label="${cmd("align-center")}" event="align-center" icon-only></docen-ribbon-button>
              <docen-ribbon-button icon="align-right" label="${cmd("align-right")}" event="align-right" icon-only></docen-ribbon-button>
              <docen-ribbon-button icon="justify" label="${cmd("justify")}" event="justify" icon-only></docen-ribbon-button>
              <span class="rb-vsep"></span>
              <docen-ribbon-split-button icon="line-spacing" label="${cmd("line-spacing")}" event="line-spacing" icon-only items='${spacingItems()}'></docen-ribbon-split-button>
              <docen-color-picker icon="shading" label="${cmd("shading")}" event="shading" default-color="FFFF00" icon-only></docen-color-picker>
              <docen-ribbon-split-button icon="border" label="${cmd("border")}" event="border" icon-only items='${borderItems()}'></docen-ribbon-split-button>
            </div>
          </div>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("styles")}" launcher="styles-pane">
          <div class="rb-col">
            <docen-ribbon-combobox value="Normal" event="style" items='${styleItems(styles)}'></docen-ribbon-combobox>
          </div>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("editing")}" launcher="find-dialog">
          <docen-ribbon-split-button icon="search" label="${cmd("search")}" event="search" size="large" items='${findItems()}'></docen-ribbon-split-button>
          <div class="rb-col">
            <docen-ribbon-button icon="replace" label="${cmd("replace")}" event="replace" icon-only></docen-ribbon-button>
            <docen-ribbon-split-button icon="board" label="${cmd("select")}" event="select" icon-only items='${selectItems()}'></docen-ribbon-split-button>
          </div>
        </docen-ribbon-group>
      </docen-ribbon-panel>`;

// --- Insert ------------------------------------------------------------------

const insertPanel = (): string => `
      <docen-ribbon-panel value="insert">
        <docen-ribbon-group label="${grp("pages")}">
          <docen-ribbon-split-button icon="page-break" label="${cmd("page-break")}" event="page-break" size="large" items='${coverItems()}'></docen-ribbon-split-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("tables")}">
          <docen-ribbon-split-button icon="table-add" label="${cmd("insert-table")}" event="insert-table" size="large" items='${tableItems()}'></docen-ribbon-split-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("illustrations")}">
          <docen-ribbon-button icon="picture" label="${cmd("insert-picture")}" event="insert-picture" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="online-picture" label="${cmd("online-picture")}" event="online-picture" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="shapes" label="${cmd("shapes")}" event="shapes" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="icon-library" label="${cmd("icons")}" event="icons" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="3d-model" label="${cmd("3d-model")}" event="3d-model" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="smartart" label="${cmd("smartart")}" event="smartart" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="chart" label="${cmd("chart")}" event="chart" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="insert-picture" label="${cmd("screenshot")}" event="screenshot" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("links")}">
          <docen-ribbon-button icon="hyperlink" label="${cmd("link")}" event="link" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="bookmark" label="${cmd("bookmark")}" event="bookmark" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="comment-add" label="${cmd("comment")}" event="comment" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("header-footer")}">
          <docen-ribbon-button icon="header" label="${cmd("header")}" event="header" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="footer" label="${cmd("footer")}" event="footer" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="page-number" label="${cmd("page-number")}" event="page-number" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("text")}">
          <docen-ribbon-button icon="text-box" label="${cmd("text-box")}" event="text-box" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="wordart" label="${cmd("wordart")}" event="wordart" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("symbols")}">
          <docen-ribbon-button icon="equation" label="${cmd("equation")}" event="equation" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="symbol" label="${cmd("symbol")}" event="symbol" size="large"></docen-ribbon-button>
        </docen-ribbon-group>
      </docen-ribbon-panel>`;

// --- Draw --------------------------------------------------------------------

const drawPanel = (): string => `
      <docen-ribbon-panel value="draw">
        <docen-ribbon-group label="${grp("pens")}">
          <docen-ribbon-button icon="pen" label="${cmd("draw-pen")}" event="draw-pen" size="large"></docen-ribbon-button>
          <div class="rb-col">
            <div class="rb-grid">
              <docen-ribbon-button icon="pencil" label="${cmd("draw-pencil")}" event="draw-pencil"></docen-ribbon-button>
              <docen-ribbon-button icon="highlight" label="${cmd("draw-highlighter")}" event="draw-highlighter"></docen-ribbon-button>
              <docen-ribbon-button icon="eraser" label="${cmd("draw-eraser")}" event="draw-eraser"></docen-ribbon-button>
            </div>
          </div>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("draw-tools")}">
          <docen-ribbon-button icon="lasso" label="${cmd("lasso-select")}" event="lasso-select" size="large"></docen-ribbon-button>
          <div class="rb-col">
            <div class="rb-grid">
              <docen-ribbon-button icon="board" label="${cmd("select-objects")}" event="select-objects"></docen-ribbon-button>
              <docen-ribbon-button icon="action-pen" label="${cmd("action-pen")}" event="action-pen"></docen-ribbon-button>
            </div>
          </div>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("ink-convert")}">
          <docen-ribbon-button icon="ink-shape" label="${cmd("ink-to-shape")}" event="ink-to-shape" size="large"></docen-ribbon-button>
          <div class="rb-col">
            <div class="rb-grid">
              <docen-ribbon-button icon="equation" label="${cmd("ink-to-math")}" event="ink-to-math"></docen-ribbon-button>
              <docen-ribbon-button icon="sync" label="${cmd("replay-ink")}" event="replay-ink"></docen-ribbon-button>
            </div>
          </div>
        </docen-ribbon-group>
      </docen-ribbon-panel>`;

// --- Design ------------------------------------------------------------------

const designPanel = (): string => `
      <docen-ribbon-panel value="design">
        <docen-ribbon-group label="${grp("document-formatting")}" launcher="themes-dialog">
          <docen-ribbon-button icon="theme" label="${cmd("theme")}" event="theme" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="font-color" label="${cmd("colors")}" event="colors" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="text-font" label="${cmd("fonts")}" event="fonts" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="text-effects" label="${cmd("effects")}" event="effects" size="large"></docen-ribbon-button>
          <div class="rb-col">
            <div class="rb-grid">
              <docen-ribbon-button icon="line-spacing" label="${cmd("paragraph-spacing")}" event="paragraph-spacing"></docen-ribbon-button>
              <docen-ribbon-button icon="page-border" label="${cmd("set-default")}" event="set-default"></docen-ribbon-button>
            </div>
          </div>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("page-background")}">
          <docen-ribbon-button icon="page-color" label="${cmd("watermark")}" event="watermark" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="page-color" label="${cmd("page-color")}" event="page-color" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="page-border" label="${cmd("page-border")}" event="page-border" size="large"></docen-ribbon-button>
        </docen-ribbon-group>
      </docen-ribbon-panel>`;

// --- Layout ------------------------------------------------------------------

const layoutPanel = (): string => `
      <docen-ribbon-panel value="layout">
        <docen-ribbon-group label="${grp("page-setup")}" launcher="page-setup-dialog">
          <docen-ribbon-split-button icon="page-color" label="${cmd("margins")}" event="margins" size="large" items='${marginsItems()}'></docen-ribbon-split-button>
          <docen-ribbon-split-button icon="orientation" label="${cmd("orientation")}" event="orientation" size="large" items='${orientationItems()}'></docen-ribbon-split-button>
          <docen-ribbon-split-button icon="page-color" label="${cmd("page-size")}" event="page-size" size="large" items='${sizePaperItems()}'></docen-ribbon-split-button>
          <docen-ribbon-split-button icon="multilevel" label="${cmd("columns")}" event="columns" size="large" items='${columnsItems()}'></docen-ribbon-split-button>
          <docen-ribbon-split-button icon="page-break" label="${cmd("breaks")}" event="page-break" size="large" items='${breaksItems()}'></docen-ribbon-split-button>
          <docen-ribbon-button icon="number-symbol" label="${cmd("line-numbers")}" event="line-numbers" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("paragraph")}" launcher="paragraph-dialog">
          <docen-ribbon-split-button icon="indent-increase" label="${cmd("indent-left")}" event="indent-left" size="large" items='${indentItems()}'></docen-ribbon-split-button>
          <docen-ribbon-split-button icon="line-spacing" label="${cmd("spacing")}" event="spacing" size="large" items='${spacingItems()}'></docen-ribbon-split-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("arrange")}">
          <div class="rb-col">
            <div class="rb-row">
              <docen-ribbon-button icon="orientation" label="${cmd("position")}" event="position"></docen-ribbon-button>
              <docen-ribbon-button icon="wrap" label="${cmd("wrap")}" event="wrap"></docen-ribbon-button>
            </div>
            <div class="rb-row">
              <docen-ribbon-button icon="orientation" label="${cmd("bring-forward")}" event="bring-forward"></docen-ribbon-button>
              <docen-ribbon-button icon="orientation" label="${cmd("send-backward")}" event="send-backward"></docen-ribbon-button>
            </div>
          </div>
          <docen-ribbon-split-button icon="align-left" label="${cmd("align")}" event="align" size="large" items='${alignItems()}'></docen-ribbon-split-button>
          <docen-ribbon-split-button icon="group-objects" label="${cmd("group")}" event="group" size="large" items='${groupItems()}'></docen-ribbon-split-button>
          <docen-ribbon-split-button icon="rotate" label="${cmd("rotate")}" event="rotate" size="large" items='${rotateItems()}'></docen-ribbon-split-button>
        </docen-ribbon-group>
      </docen-ribbon-panel>`;

// --- References --------------------------------------------------------------

const referencesPanel = (): string => `
      <docen-ribbon-panel value="references">
        <docen-ribbon-group label="${grp("toc")}">
          <docen-ribbon-button icon="toc" label="${cmd("toc")}" event="toc" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="multilevel" label="${cmd("add-text")}" event="add-text" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="sync" label="${cmd("update-toc")}" event="update-toc" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("footnotes")}">
          <docen-ribbon-split-button icon="footnote" label="${cmd("insert-footnote")}" event="insert-footnote" size="large" items='${footnoteItems()}'></docen-ribbon-split-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("citations")}">
          <docen-ribbon-button icon="comment-add" label="${cmd("insert-citation")}" event="insert-citation" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="people" label="${cmd("manage-sources")}" event="manage-sources" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="document-print" label="${cmd("bibliography")}" event="bibliography" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("captions")}">
          <docen-ribbon-button icon="comment-add" label="${cmd("insert-caption")}" event="insert-caption" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="document-print" label="${cmd("table-of-figures")}" event="table-of-figures" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="sync" label="${cmd("update-figures")}" event="update-figures" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="link" label="${cmd("cross-reference")}" event="cross-reference" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("index")}">
          <docen-ribbon-button icon="comment-add" label="${cmd("mark-entry")}" event="mark-entry" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="document-print" label="${cmd("insert-index")}" event="insert-index" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="sync" label="${cmd("update-index")}" event="update-index" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("toa")}">
          <docen-ribbon-button icon="comment-add" label="${cmd("mark-citation")}" event="mark-citation" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="document-print" label="${cmd("insert-toa")}" event="insert-toa" size="large"></docen-ribbon-button>
        </docen-ribbon-group>
      </docen-ribbon-panel>`;

// --- Mailings ----------------------------------------------------------------

const mailingsPanel = (): string => `
      <docen-ribbon-panel value="mailings">
        <docen-ribbon-group label="${grp("create")}">
          <docen-ribbon-button icon="mail" label="${cmd("envelopes")}" event="envelopes" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="mail" label="${cmd("labels")}" event="labels" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("start-merge")}">
          <docen-ribbon-split-button icon="document-print" label="${cmd("start-merge")}" event="start-merge" size="large" items='${startMergeItems()}'></docen-ribbon-split-button>
          <docen-ribbon-button icon="people" label="${cmd("select-recipients")}" event="select-recipients" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="edit" label="${cmd("edit-recipients")}" event="edit-recipients" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("write-fields")}">
          <docen-ribbon-button icon="document-print" label="${cmd("address-block")}" event="address-block" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="comment-add" label="${cmd("greeting-line")}" event="greeting-line" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="link" label="${cmd("merge-field")}" event="merge-field" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="highlight" label="${cmd("highlight-merge")}" event="highlight-merge" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("preview")}">
          <docen-ribbon-button icon="search" label="${cmd("preview-results")}" event="preview-results" size="large"></docen-ribbon-button>
          <div class="rb-col">
            <div class="rb-grid">
              <docen-ribbon-button icon="align-left" label="${cmd("first-record")}" event="first-record"></docen-ribbon-button>
              <docen-ribbon-button icon="align-right" label="${cmd("last-record")}" event="last-record"></docen-ribbon-button>
            </div>
          </div>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("finish")}">
          <docen-ribbon-split-button icon="document-print" label="${cmd("finish-merge")}" event="finish-merge" size="large" items='${finishMergeItems()}'></docen-ribbon-split-button>
        </docen-ribbon-group>
      </docen-ribbon-panel>`;

// --- Review ------------------------------------------------------------------

const reviewPanel = (): string => `
      <docen-ribbon-panel value="review">
        <docen-ribbon-group label="${grp("proofing")}">
          <docen-ribbon-button icon="spell-check" label="${cmd("spell-check")}" event="spell-check" size="large"></docen-ribbon-button>
          <div class="rb-col">
            <div class="rb-grid">
              <docen-ribbon-button icon="word-count" label="${cmd("word-count")}" event="word-count"></docen-ribbon-button>
              <docen-ribbon-button icon="search" label="${cmd("thesaurus")}" event="thesaurus"></docen-ribbon-button>
            </div>
          </div>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("accessibility")}">
          <docen-ribbon-button icon="checkmark-circle" label="${cmd("check-accessibility")}" event="check-accessibility" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("language")}">
          <docen-ribbon-button icon="link" label="${cmd("translate")}" event="translate" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="text-font" label="${cmd("language")}" event="language" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("comments")}">
          <docen-ribbon-button icon="comment-add" label="${cmd("new-comment")}" event="new-comment" size="large"></docen-ribbon-button>
          <div class="rb-col">
            <div class="rb-grid">
              <docen-ribbon-button icon="edit" label="${cmd("edit-comment")}" event="edit-comment"></docen-ribbon-button>
              <docen-ribbon-button icon="close" label="${cmd("delete-comment")}" event="delete-comment"></docen-ribbon-button>
            </div>
          </div>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("tracking")}">
          <docen-ribbon-button icon="group-objects" label="${cmd("track-changes")}" event="track-changes" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("changes")}">
          <docen-ribbon-button icon="accept" label="${cmd("accept-change")}" event="accept-change" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="close" label="${cmd("reject-change")}" event="reject-change" size="large"></docen-ribbon-button>
          <div class="rb-col">
            <div class="rb-grid">
              <docen-ribbon-button icon="align-left" label="${cmd("previous-change")}" event="previous-change"></docen-ribbon-button>
              <docen-ribbon-button icon="align-right" label="${cmd("next-change")}" event="next-change"></docen-ribbon-button>
            </div>
          </div>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("compare")}">
          <docen-ribbon-button icon="group-objects" label="${cmd("compare")}" event="compare" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="group-objects" label="${cmd("combine")}" event="combine" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("protect")}">
          <docen-ribbon-button icon="protect" label="${cmd("restrict-editing")}" event="restrict-editing" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="protect" label="${cmd("protect-document")}" event="protect-document" size="large"></docen-ribbon-button>
        </docen-ribbon-group>
      </docen-ribbon-panel>`;

// --- View --------------------------------------------------------------------

const viewPanel = (): string => `
      <docen-ribbon-panel value="view">
        <docen-ribbon-group label="${grp("views")}">
          <docen-ribbon-button icon="print" label="${cmd("print-layout")}" event="print-layout" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="document-print" label="${cmd("web-layout")}" event="web-layout" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="document-print" label="${cmd("read-mode")}" event="read-mode" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="group-objects" label="${cmd("outline")}" event="outline" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="document-print" label="${cmd("draft")}" event="draft" size="large"></docen-ribbon-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("show")}">
          <div class="rb-col">
            <docen-ribbon-button icon="ruler" label="${cmd("toggle-ruler")}" event="toggle-ruler"></docen-ribbon-button>
            <docen-ribbon-button icon="gridlines" label="${cmd("toggle-gridlines")}" event="toggle-gridlines"></docen-ribbon-button>
            <docen-ribbon-button icon="replace" label="${cmd("toggle-navigation")}" event="toggle-navigation"></docen-ribbon-button>
          </div>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("zoom")}">
          <docen-ribbon-button icon="zoom-in" label="${cmd("zoom")}" event="zoom" size="large"></docen-ribbon-button>
          <docen-ribbon-split-button icon="zoom-in" label="${cmd("zoom-100")}" event="zoom-100" size="large" items='${zoomItems()}'></docen-ribbon-split-button>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("window")}">
          <docen-ribbon-button icon="grid" label="${cmd("new-window")}" event="new-window" size="large"></docen-ribbon-button>
          <div class="rb-col">
            <div class="rb-grid">
              <docen-ribbon-button icon="grid" label="${cmd("arrange-all")}" event="arrange-all"></docen-ribbon-button>
              <docen-ribbon-button icon="group-objects" label="${cmd("split-window")}" event="split-window"></docen-ribbon-button>
            </div>
          </div>
        </docen-ribbon-group>

        <docen-ribbon-group label="${grp("macros")}">
          <docen-ribbon-button icon="group-objects" label="${cmd("view-macros")}" event="view-macros" size="large"></docen-ribbon-button>
          <docen-ribbon-button icon="edit" label="${cmd("record-macro")}" event="record-macro" size="large"></docen-ribbon-button>
        </docen-ribbon-group>
      </docen-ribbon-panel>`;
