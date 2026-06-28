import {
  convertMillimetersToTwip,
  createDocxEditor,
  effectiveRunProps,
  generateDOCX,
  parseDOCX,
  parseHTML,
  resolveFontName,
  sectionPageSizeDefaults,
  stylesToCss,
  twipsToMm,
  type JSONContent,
  type SectionPropertiesOptions,
  type StylesOptions,
} from "@docen/docx";
import { Extension, type Editor } from "@docen/docx/core";
import { ListKeymap } from "@tiptap/extension-list";
import { TableOfContents } from "@tiptap/extension-table-of-contents";
import {
  CharacterCount,
  Dropcursor,
  Focus,
  Gapcursor,
  Placeholder,
  Selection,
  TrailingNode,
  UndoRedo,
} from "@tiptap/extensions";
import type { Mark } from "@tiptap/pm/model";
import { EditorState } from "@tiptap/pm/state";
import {
  findNext,
  findPrev,
  getMatchHighlights,
  replaceAll,
  replaceNext,
  search,
  setSearchState,
  SearchQuery,
} from "prosemirror-search";

import { applyTheme, observeLang, registerComponents, t } from "../ui";
import type { OutlineItem } from "../ui/components/workspace/outline";
import { dispatchRibbonCommand, WIRED_DISPATCH } from "./commands";
// Side-effect import: registers the ribbon/header translation tables.
import "./i18n";
import { FontMetricDecoration } from "./extensions/font-metric";
import { ImageCap } from "./extensions/image-cap";
import { DocenKeymap } from "./extensions/keymap";
import { PageBreakView } from "./extensions/page-break";
import { Page, PageDocument } from "./extensions/page-node";
import { PagePlugin, pageStorageOf } from "./extensions/page-plugin";
import { SectionBreakMarks } from "./extensions/section-break";
import { SplitMarks } from "./extensions/split-paragraph";
import { SplitTable, SplitTableRow } from "./extensions/split-table";
import { buildRibbonInnerHTML, RIBBON_TAB_IDS, type RibbonTabId } from "./ribbon-default";
import { fontNormalRatio } from "./utils/font-metric";
import { clearMeasureCache } from "./utils/measure";
import { unwrapPages, wrapPages } from "./utils/wrap";

/** Ribbon commands handled locally in #onCommand/#onChange (not via
 *  dispatchRibbonCommand). Together with {@link WIRED_DISPATCH} this is the
 *  "wired" set used to grey out unwired skeleton commands. lang-zh/lang-en are
 *  header menu items, not ribbon commands, so excluded. */
const LOCAL_HANDLED: ReadonlySet<string> = new Set([
  // #onCommand
  "toggle-navigation",
  "search",
  "replace",
  "page-size",
  "orientation",
  "margins",
  "zoom",
  "zoom-100",
  "save",
  "insert-picture",
  "show-marks",
  "copy",
  "cut",
  "paste",
  "select",
  "format-painter",
  // #onChange (data-event)
  "open",
  "save-as",
  "print",
]);

/** Parse a `tabs="home,review"` attribute into a validated whitelist of tab ids;
 *  undefined when unset/empty/invalid (=> render all tabs). */
function parseTabs(attr: string | null): readonly RibbonTabId[] | undefined {
  if (!attr) return undefined;
  const known = RIBBON_TAB_IDS as readonly string[];
  const ids = attr
    .split(",")
    .map((s) => s.trim())
    .filter((s): s is RibbonTabId => known.includes(s));
  return ids.length > 0 ? ids : undefined;
}

const TEMPLATE = `
  <style>
    /* Cascade layers — declared ONCE up front so layer ORDER (not specificity)
       governs priority below. reset strips UA defaults; docxStyles holds the
       document's named styles (styles.xml). Both lose to unlayered rules (page
       geometry, crop marks, inline OOXML styles). */
    @layer reset, docxStyles;
    :host { display: flex; flex-direction: column; height: 100%; }
    /* Office ribbon group layout helpers — a large button beside stacked rows of
       small icon-only buttons. Applied to light-DOM wrappers in the ribbon. */
    .rb-col { display: flex; flex-direction: column; gap: 2px; }
    .rb-row { display: flex; flex-direction: row; align-items: center; gap: 2px; flex-wrap: wrap; }
    /* Small icon-only buttons as a 3-row column-flow grid: buttons stack into
       columns of ≤3 (Word's compact group layout), not a flat single row. */
    .rb-grid { display: grid; grid-template-rows: repeat(3, auto); grid-auto-flow: column; gap: 2px; align-content: start; }
    .rb-vsep { width: 1px; align-self: stretch; background: var(--docen-color-divider, #e1e1e1); margin: 0 2px; }
    .avatar {
      display: inline-flex; align-items: center; justify-content: center;
      width: 20px; height: 20px; border-radius: 50%;
      background: var(--docen-color-brand, #0078d4); color: #fff;
      font-size: 10px; font-weight: 600; margin-inline-end: 4px;
    }
    .avatar-img { object-fit: cover; background: none; }
    /* The editor wrapper (.docen-pages) hosts the Tiptap .ProseMirror, which
       renders one .docen-page NODE per page. The wrapper just centers the
       flow; each page node is its own fixed paper sheet (C-route — see
       CLAUDE.md). */
    .docen-pages .ProseMirror { padding: 0; }
    /* CSS reset inside the editor surface — strip the browser's UA defaults so
       OOXML properties (inline styles + named styles in the docxStyles layer)
       are the SOLE source of truth. The UA stylesheet gives <p>/<h1-6> a 1em
       margin, resets heading font-size/weight, pads <ul>/<ol>, etc. — all of
       which corrupt pagination (a one-page cover spills to two). These rules
       live in @layer reset, declared BEFORE docxStyles, so the document's named
       styles (.docx-style-*) ALWAYS win regardless of specificity — that is the
       real fix for the Heading bold/centering bug (the old specificity juggling
       is gone). Table borders and list markers are handled in their own rules
       below. */
    @layer reset {
      .docen-pages .ProseMirror p,
      .docen-pages .ProseMirror h1, .docen-pages .ProseMirror h2,
      .docen-pages .ProseMirror h3, .docen-pages .ProseMirror h4,
      .docen-pages .ProseMirror h5, .docen-pages .ProseMirror h6,
      .docen-pages .ProseMirror blockquote, .docen-pages .ProseMirror figure,
      .docen-pages .ProseMirror pre,
      .docen-pages .ProseMirror ul, .docen-pages .ProseMirror ol {
        margin: 0; padding: 0;
      }
      /* Clear the UA heading defaults (2em font-size, bold weight) so headings
         take the doc default unless a named style overrides — same reset layer,
         so a .docx-style-Heading* font-size/font-weight always wins. */
      .docen-page h1, .docen-page h2,
      .docen-page h3, .docen-page h4,
      .docen-page h5, .docen-page h6 {
        font-size: inherit; font-weight: inherit;
      }
    }
    /* .ProseMirror's default focus outline paints a black border on every
       click — drop it (the caret + selection still mark focus). */
    .docen-pages .ProseMirror:focus { outline: none; }
    /* Each page node = a fixed paper sheet. 'height' (NOT min-height) +
       overflow: hidden forces overflow into the next page instead of
       stretching the sheet — the C-route invariant. Geometry comes from
       <docen-canvas> CSS vars inherited through the shadow boundary. */
    .docen-pages .docen-page {
      width: var(--docen-page-width, 210mm);
      height: var(--docen-page-min-height, 297mm);
      overflow: hidden;
      box-sizing: border-box;
      padding: var(--docen-page-margin, 25.4mm);
      margin: 0 auto var(--docen-page-gap, 24px);
      background-color: var(--docen-color-page, #ffffff);
      box-shadow: 0 1px 4px rgba(0, 0, 0, 0.12);
      position: relative;
      /* content-visibility:auto is DISABLED here: it skips layout for off-screen
         pages, but the paginator measures block/row heights by reading the DOM
         (measureFlatItems — Pretext covers text blocks; table rows still use
         getBoundingClientRect). Reading a rect on a cv:auto-skipped page returns
         a placeholder/0, NOT a forced layout, so reflow measures a different
         layout every round and never converges — it re-flows ~1.5s/round
         indefinitely (verified on a 1000+ page doc). Re-enable only once row
         measurement is deterministic (Pretext) or reflow forces layout while
         measuring. */
      /* content-visibility: auto; */
      /* contain-intrinsic-size: var(--docen-page-width, 210mm) var(--docen-page-min-height, 297mm); */
    }
    /* A wrapNone floating drawing (image / wpg group) anchored to its paragraph
       (verticalPosition.relative = paragraph, the OOXML default) renders
       position:absolute. Its offsetParent must be that paragraph, not the page
       box, or top/left resolve from the page top and the drawing floats over
       the heading/body text. Making the anchor <p> position:relative pins the
       drawing to the blank line it belongs on (matches Word: a floating group
       overlays its own empty paragraph, over the body text below it). */
    .docen-pages .docen-page p:has([data-float-anchor="paragraph"]) {
      position: relative;
    }
    /* Images cap to the section content width the way Word caps them: a wider
       image scales DOWN to fit, never upscales. The ImageCap extension sets the
       real width on data-URL images it can sync-decode; this rule is the visual
       fallback for what it skips (http(s) URLs, docs without section geometry)
       so nothing overflows the fixed page box. crop images render an enlarged
       inner <img>, so opt them out or the cap collapses the crop geometry. */
    .docen-pages .docen-page img {
      max-width: 100%;
      height: auto;
    }
    .docen-pages .docen-page span[data-image="crop"] > img {
      max-width: none;
    }
    /* EMF/WMF (Office GDI vector) images can't be decoded by the browser, so
       Image.renderHTML emits a div[data-image=vector] placeholder carrying the
       real art in data-vector-src. Paint it as a dashed, hatched, labeled box
       so the gap reads as "unsupported vector art" instead of a silent empty
       rectangle. Inline width/height (from the image extent) size the box;
       inline-block (set in renderImageStyles) keeps it in-flow. */
    .docen-pages .docen-page [data-image="vector"] {
      box-sizing: border-box;
      border: 1px dashed #b8b8b8;
      background: repeating-linear-gradient(
        45deg,
        #f6f6f6,
        #f6f6f6 9px,
        #efefef 9px,
        #efefef 18px
      );
      color: #9a9a9a;
      font-size: 12px;
      text-align: center;
      padding-top: 6px;
      overflow: hidden;
    }
    /* prosemirror-tables CellSelection tags each selected cell with class
       "selectedCell"; paint it with Word's translucent selection blue so a
       multi-cell selection is actually visible. box-shadow (not
       background-color) overlays a cell's own fill without being beaten by
       the cell's inline background-color. */
    .ProseMirror .selectedCell {
      box-shadow: inset 0 0 0 9999px rgba(0, 120, 215, 0.18);
    }
    /* Crop marks — four L-brackets OUTSIDE the content box, in the page-margin
       gutter, each L's corner pointing AT the editable area: the vertex sits
       just outside a content-box corner and the two 23px legs reach into the
       margin. Drawn on a ::before that covers the whole page (inset: 0) and
       carries the page margin as its own padding, so its content-box == the
       page's content box — that is the origin the background-position
       offsets are measured from (negative for the top legs, positive for the
       bottom), landing each leg in the margin gutter (not over text). The
       ::before is used (not the page node's own background) because a
       ProseMirror-managed node does not paint its own background-image, but its
       pseudo-element does. */
    .docen-pages .docen-page::before {
      content: "";
      position: absolute;
      inset: 0;
      padding: var(--docen-page-margin, 25.4mm);
      pointer-events: none;
      background-image:
        linear-gradient(var(--docen-color-crop, #c0c0c0), var(--docen-color-crop, #c0c0c0)),
        linear-gradient(var(--docen-color-crop, #c0c0c0), var(--docen-color-crop, #c0c0c0)),
        linear-gradient(var(--docen-color-crop, #c0c0c0), var(--docen-color-crop, #c0c0c0)),
        linear-gradient(var(--docen-color-crop, #c0c0c0), var(--docen-color-crop, #c0c0c0)),
        linear-gradient(var(--docen-color-crop, #c0c0c0), var(--docen-color-crop, #c0c0c0)),
        linear-gradient(var(--docen-color-crop, #c0c0c0), var(--docen-color-crop, #c0c0c0)),
        linear-gradient(var(--docen-color-crop, #c0c0c0), var(--docen-color-crop, #c0c0c0)),
        linear-gradient(var(--docen-color-crop, #c0c0c0), var(--docen-color-crop, #c0c0c0));
      background-position:
        -24px -2px, -2px -24px,
        calc(100% + 24px) -2px, calc(100% + 2px) -24px,
        -24px calc(100% + 2px), -2px calc(100% + 24px),
        calc(100% + 24px) calc(100% + 2px), calc(100% + 2px) calc(100% + 24px);
      background-size: 23px 1px, 1px 23px, 23px 1px, 1px 23px, 23px 1px, 1px 23px, 23px 1px, 1px 23px;
      background-origin: content-box;
      background-repeat: no-repeat;
    }
    /* Grey the "Auto-save" label to match its disabled switch (skeleton
       feature), so the label + switch read as one unavailable control, like
       ribbon skeleton buttons. Lifts automatically once the switch loses
       disabled. */
    .autosave-label:has(+ fluent-switch[disabled]) {
      color: var(--docen-color-text-3, #8a8a8a);
    }
    /* While focus is in a ribbon combobox dropdown, the editor is blurred and
       the browser stops painting its selection. The Tiptap Selection extension
       stamps a .selection class on the range so it stays visible. Uses the
       system selection colors (Highlight/HighlightText) to match the browser's
       native ::selection, including high-contrast and custom OS themes. */
    .ProseMirror .selection {
      background: Highlight;
      color: HighlightText;
    }
    /* Placeholder — the Tiptap Placeholder extension only stamps an is-empty
       class + data-placeholder attribute on empty nodes; this CSS paints the
       label. :first-child shows it just on the document's first empty paragraph
       (Word shows the prompt once, at the top); showOnlyCurrent (default) means
       only the caret's empty node carries the attribute. float + height:0 keep
       the label in-flow at the paragraph start without consuming layout, so it
       never pushes real content or skews pagination measurement. */
    .docen-pages .ProseMirror p.is-empty:first-child::before {
      content: attr(data-placeholder);
      color: var(--docen-color-text-3, #8a8a8a);
      pointer-events: none;
      float: left;
      height: 0;
    }
    /* Status bar (slotted into docen-workspace's .rb-shell-status, already a
       flex row). Lay the footer as a left cluster (page count + word count)
       and a right (zoom %), matching Word's bottom status bar. */
    footer[slot="status"] {
      display: flex;
      justify-content: space-between;
      align-items: center;
      width: 100%;
    }
    .docen-status-left { display: flex; gap: 14px; }
    /* Right cluster — Word's zoom control: a minus / plus button flanking a
       draggable slider, then the percent. The slider is a native range input
       styled to a Fluent track + accent thumb. */
    .docen-status-zoom { display: flex; align-items: center; gap: 4px; }
    .docen-zoom-step {
      width: 18px; height: 18px; padding: 0;
      border: 1px solid var(--docen-color-stroke-1, #c7c7c7);
      border-radius: 3px; background: transparent;
      color: var(--docen-color-text-1, #242424);
      font-size: 13px; line-height: 1; cursor: pointer;
      display: inline-flex; align-items: center; justify-content: center;
    }
    .docen-zoom-step:hover { background: var(--docen-color-subtle-background-hover, #f5f5f5); }
    .docen-zoom-slider {
      -webkit-appearance: none; appearance: none;
      width: 90px; height: 3px; margin: 0;
      background: var(--docen-color-stroke-1, #c7c7c7);
      border-radius: 2px; cursor: pointer;
    }
    .docen-zoom-slider::-webkit-slider-thumb {
      -webkit-appearance: none; appearance: none;
      width: 11px; height: 11px; border: none; border-radius: 50%;
      background: var(--docen-color-accent, #0f6cbd); cursor: pointer;
    }
    .docen-zoom-slider::-moz-range-thumb {
      width: 11px; height: 11px; border: none; border-radius: 50%;
      background: var(--docen-color-accent, #0f6cbd); cursor: pointer;
    }
    .docen-zoom-pct { min-width: 38px; text-align: right; }
    /* Tables — Word inserts tables in the Table Grid style (a single black
       border on every cell). Without this a freshly inserted table is
       invisible: the docx table extension emits a border only when the node
       carries border attrs, and insertTable creates none. */
    .docen-pages table { border-collapse: collapse; }
    .docen-pages table td,
    .docen-pages table th {
      border: 1px solid #000;
      /* OOXML defaults w:tcMar top/bottom to 0 (TableNormal); the UA td padding
         (1px) would inflate every row ~2px vs Word. Horizontal padding stays at
         the cell's w:tcMar (set inline by renderTableCellStyles) or 0. */
      padding-block: 0;
    }
    /* Formatting marks (Show/Hide ¶) — Word shows these only while editing
       (non-printing). The show-marks command flips the host [show-marks]
       attribute; the marks themselves live entirely in CSS. */
    /* Pilcrow ¶ is painted by the FormattingMarks extension as a widget
       decoration — CSS ::after on a ProseMirror-managed <p> does not render. */
    /* Zero-width inline-block: the mark hugs the last character and never
       breaks to its own line on a full line. text-indent:0 cancels the
       inherited paragraph indent (an inline-block is a block container), or
       the glyph drifts right. */
    .docen-pages .docen-para-mark {
      color: var(--docen-color-marks, #6e6e6e);
      user-select: none;
      pointer-events: none;
      margin-inline-start: 1px;
      display: inline-block;
      width: 0;
      overflow: visible;
      vertical-align: baseline;
      text-indent: 0;
    }
    /* A page break renders as a Fluent divider with a centered label (Word).
       The NodeView (PageBreakView) supplies the fluent-divider; it is hidden
       unless show-marks is on. */
    .docen-pages [data-type="pageBreak"] { display: block; line-height: 0; }
    .docen-pages [data-type="pageBreak"] fluent-divider { display: none; }
    :host([show-marks]) .docen-pages [data-type="pageBreak"] {
      margin: 8px 0;
      line-height: normal;
    }
    :host([show-marks]) .docen-pages [data-type="pageBreak"] fluent-divider {
      display: flex;
      font-size: 0.8em;
    }
    /* A section break renders as a Fluent divider after the section-carrying
       paragraph (Word: the boundary only shows the marker while editing). The
       SectionBreakMarks widget supplies the fluent-divider; hidden unless
       show-marks is on — same mechanism as the page-break marker. */
    .docen-pages [data-section-break] { line-height: 0; }
    .docen-pages [data-section-break] fluent-divider { display: none; }
    :host([show-marks]) .docen-pages [data-section-break] {
      margin: 8px 0;
      line-height: normal;
    }
    :host([show-marks]) .docen-pages [data-section-break] fluent-divider {
      display: flex;
      font-size: 0.8em;
    }
    /* Find highlights (prosemirror-search) — Word's yellow-match / orange-active. */
    .docen-pages .ProseMirror-search-match {
      background: rgba(255, 235, 59, 0.55);
    }
    .docen-pages .ProseMirror-active-search-match {
      background: rgba(255, 145, 0, 0.75);
    }
    /* Find Results — Word-style match list: each hit rendered with surrounding
       context and a data-from/to for click-to-jump. Padding keeps items off the
       pane edge (the previous "N matches" text butted right against it). */
    .search-results {
      padding: 6px 8px;
      box-sizing: border-box;
    }
    .search-results .result-count {
      font-size: 12px;
      color: var(--docen-color-marks, #6e6e6e);
      padding: 2px 4px 8px;
    }
    .search-results .result-item {
      display: block;
      width: 100%;
      text-align: start;
      border: none;
      background: transparent;
      padding: 5px 8px;
      margin-block-end: 2px;
      border-radius: 4px;
      font-family: inherit;
      font-size: 12px;
      line-height: 1.45;
      color: #3b3b3b;
      cursor: pointer;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .search-results .result-item:hover {
      background: var(--docen-color-hover, rgba(0, 0, 0, 0.06));
    }
    .search-results .result-item mark {
      background: rgba(255, 235, 59, 0.85);
      color: inherit;
      font-weight: 600;
    }
    /* Print: drop shadow/gap/marks; each page node is its own printed sheet. */
    @media print {
      .docen-pages .docen-page {
        content-visibility: visible;
        box-shadow: none;
        margin: 0;
        break-after: page;
      }
      .docen-pages .docen-page::before { display: none; }
      /* Formatting marks + search highlights never print (editing-only). */
      .docen-pages .docen-para-mark,
      .docen-pages [data-type="pageBreak"],
      .docen-pages [data-section-break],
      .docen-pages .ProseMirror-search-match,
      .docen-pages .ProseMirror-active-search-match {
        display: none !important;
      }
    }
  </style>
  <docen-workspace>
    <docen-app-header slot="header" part="header"></docen-app-header>
    <docen-ribbon slot="ribbon" part="ribbon"></docen-ribbon>
    <docen-task-pane slot="task-pane-start" position="start" open part="nav-pane">
      <docen-navigation-pane>
        <docen-outline slot="headings"></docen-outline>
        <div class="search-results" slot="results" part="search-results"></div>
      </docen-navigation-pane>
    </docen-task-pane>
    <docen-canvas>
      <div class="docen-pages" part="page"></div>
    </docen-canvas>
    <docen-task-pane slot="task-pane-end" position="end" part="props-pane">
      <docen-properties-panel></docen-properties-panel>
    </docen-task-pane>
    <footer slot="status" part="status">
      <span class="docen-status-left">
        <span class="docen-status-section"></span>
        <span class="docen-status-pages"></span>
        <span class="docen-status-words"></span>
      </span>
      <span class="docen-status-zoom">
        <button
          type="button"
          class="docen-zoom-step"
          data-zoom-step="-1"
          part="zoom-out"
          aria-label="Zoom out"
        >−</button>
        <input
          type="range"
          class="docen-zoom-slider"
          min="10"
          max="500"
          step="1"
          value="100"
          part="zoom-slider"
          aria-label="Zoom level"
        />
        <button
          type="button"
          class="docen-zoom-step"
          data-zoom-step="1"
          part="zoom-in"
          aria-label="Zoom in"
        >+</button>
        <span class="docen-zoom-pct"></span>
      </span>
    </footer>
  </docen-workspace>
  <docen-find-replace-dialog></docen-find-replace-dialog>
  <input type="file" id="file-input" accept=".docx" hidden />
  <input type="file" id="image-input" accept="image/*" hidden />`;

/** A TOC anchor as emitted by the TableOfContents extension — the subset of
 *  fields we consume (id for selection, pos for jump, textContent + level for
 *  the outline tree). */
interface TocAnchor {
  id: string;
  pos: number;
  textContent: string;
  originalLevel: number;
}

/** Build a nested OutlineItem tree from the flat TOC anchor list: each heading
 *  nests under the nearest preceding heading with a smaller level. */
function buildOutlineTree(anchors: readonly TocAnchor[]): OutlineItem[] {
  type Node = { id: string; title: string; level: number; children?: Node[] };
  const roots: Node[] = [];
  const stack: Node[] = [];
  for (const a of anchors) {
    const node: Node = { id: a.id, title: a.textContent, level: a.originalLevel };
    while (stack.length && stack[stack.length - 1].level >= a.originalLevel) stack.pop();
    const parent = stack[stack.length - 1];
    if (parent) (parent.children ??= []).push(node);
    else roots.push(node);
    stack.push(node);
  }
  return roots as OutlineItem[];
}

/** prosemirror-search's plugin, wrapped as a Tiptap extension so it loads with
 *  the editor. It stores the active SearchQuery and highlights its matches
 *  (classes ProseMirror-search-match / -active-search-match); the host drives
 *  it via setSearchState and the findNext / findPrev commands. */
const Search = Extension.create({
  name: "docenSearch",
  addProseMirrorPlugins() {
    return [search()];
  },
});

/** MS Office standard paper sizes (mm, portrait width × height). Page-setup
 *  presets resolve to raw mm here; <docen-canvas> takes only raw page-width /
 *  page-height, so presets stay in this document layer, not the UI component. */
const PAPER_SIZES: Readonly<Record<string, readonly [number, number]>> = {
  letter: [215.9, 279.4],
  legal: [215.9, 355.6],
  statement: [139.7, 215.9],
  executive: [184.15, 266.7],
  tabloid: [279.4, 431.8],
  a3: [297, 420],
  a4: [210, 297],
  a5: [148, 210],
  a6: [105, 148],
  b5: [182, 257],
};

/** MS Office margin presets (mm, CSS padding list for the page content box:
 *  one value = uniform; two = top/bottom then left/right). `normal` matches the
 *  engine default (@office-open/docx sectionMarginDefaults: top/bottom 25.4mm,
 *  left/right 31.75mm = MS Office zh-CN "Normal") so the canvas fallback and the
 *  document-model sectionProperties agree. */
const MARGINS: Readonly<Record<string, string>> = {
  normal: "25.4mm 31.75mm",
  narrow: "12.7mm",
  moderate: "25.4mm 19.05mm",
  wide: "25.4mm 50.8mm",
};

/** Parse a CSS padding list (mm, 1–4 values) into OOXML page margins (twips),
 *  via the engine's convertMillimetersToTwip (mm → twips). */
function marginTwipsFromCss(css: string): {
  top: number;
  right: number;
  bottom: number;
  left: number;
} {
  const mm = css.split(/\s+/).map((s) => parseFloat(s));
  const [t, r, b, l] =
    mm.length === 1
      ? [mm[0], mm[0], mm[0], mm[0]]
      : mm.length === 2
        ? [mm[0], mm[1], mm[0], mm[1]]
        : [mm[0], mm[1], mm[2] ?? mm[1], mm[3] ?? mm[1]];
  return {
    top: convertMillimetersToTwip(t),
    right: convertMillimetersToTwip(r),
    bottom: convertMillimetersToTwip(b),
    left: convertMillimetersToTwip(l),
  };
}

/** Deep-merge a sectionProperties patch (page.size / page.margin) into a base,
 *  preserving sides/dims the patch omits — so e.g. changing only the margins
 *  keeps the page size. Reuses the engine's SectionPropertiesOptions type. */
function mergeSectionProperties(
  base: SectionPropertiesOptions | null | undefined,
  patch: SectionPropertiesOptions,
): SectionPropertiesOptions {
  const bp = base?.page;
  const pp = patch.page;
  return {
    ...base,
    page: {
      ...bp,
      ...pp,
      size: { ...bp?.size, ...pp?.size },
      margin: { ...bp?.margin, ...pp?.margin },
    },
  };
}

/**
 * `<docen-document>` — a turnkey DOCX editor super-component.
 *
 * Wires the Fluent UI shell (app-header + ribbon + canvas) to the `@docen/docx`
 * Tiptap engine, with Pretext-driven offline pagination. Drop it in for an
 * editable, paginated Word surface: the app header drives file I/O (open/save)
 * and language switching, ribbon commands route to Tiptap, embedded objects open
 * their editors on double-click, and file I/O goes through `parseDOCX`/
 * `generateDOCX`. The header + ribbon re-render on locale change.
 */
class DocenDocument extends HTMLElement {
  #editor?: Editor;
  #fileInput?: HTMLInputElement;
  #imageInput?: HTMLInputElement;
  /** Latest TOC anchors, refreshed by TableOfContents.onUpdate; used to resolve
   *  an outline click back to a document position (pos). */
  #anchors: readonly TocAnchor[] = [];
  /** Cached doc nodeSize + Word-style word count so caret-move transactions
   *  don't re-walk the whole document (recomputed only when content changes). */
  #lastDocSize = -1;
  #lastWords = 0;
  /** Semantic fingerprint of the last outline tree — id/level/title only. `pos`
   *  shifts on every pagination re-flow but never changes what the pane shows,
   *  so it's excluded; the fingerprint is built from per-anchor arrays (not the
   *  serialized tree) so object key order can never cause a spurious mismatch. */
  #outlineSig = "";
  #unobserveLang?: () => void;
  /** Tears down the transaction listener mirroring caret font/size → comboboxes. */
  #fontSyncCleanup?: () => void;
  // Format Painter captured marks + the pointerup listener that applies them.
  #painterMarks: readonly Mark[] | null = null;
  #painterOff?: () => void;
  /** Current zoom level (percent) applied to the canvas via CSS `zoom`. */
  #zoom = 100;

  /** Attributes whose runtime changes re-configure the component (chrome
   *  visibility, ribbon tabs, editable, identity). Without this list, attribute
   *  changes after connect are silently ignored. */
  static get observedAttributes(): string[] {
    return [
      "toolbar",
      "tabs",
      "header",
      "navigation-pane",
      "properties-pane",
      "status-bar",
      "editable",
      "filename",
      "closable",
      "user",
      "avatar",
    ];
  }

  attributeChangedCallback(name: string, _old: string, value: string): void {
    switch (name) {
      case "editable":
        this.#editor?.setEditable(value !== "false");
        break;
      case "filename":
      case "user":
      case "avatar":
      case "closable":
      case "tabs":
        // These only change rendered chrome — re-stamp header/ribbon.
        this.#renderChrome();
        break;
      case "toolbar":
      case "header":
      case "status-bar":
      case "navigation-pane":
      case "properties-pane":
        this.#applyChromeVisibility();
        break;
    }
  }

  /** Esc fallback: restore the ribbon to "always shown" after the browser leaves fullscreen. */
  readonly #onFullscreenChange = (): void => {
    if (document.fullscreenElement) return;
    const ribbon = this.shadowRoot?.querySelector("docen-ribbon");
    const workspace = this.shadowRoot?.querySelector("docen-workspace");
    if (ribbon?.getAttribute("data-ribbon-mode") === "auto-hide") {
      ribbon.removeAttribute("data-ribbon-mode");
      workspace?.removeAttribute("data-fullscreen");
    }
  };

  /** Ctrl+= / Ctrl+- / Ctrl+0 zoom, Ctrl+F find (Word behavior). Zoom is
   *  ignored inside ribbon comboboxes and other inputs (so the keystroke reaches
   *  them); Ctrl+F is global. preventDefault blocks the browser's native zoom/find. */
  readonly #onZoomKey = (event: KeyboardEvent): void => {
    if (!(event.ctrlKey || event.metaKey)) return;
    // Ctrl+F opens Find, Ctrl+H opens Find & Replace (Word behavior).
    if (event.key === "f" || event.key === "F") {
      event.preventDefault();
      this.#openSearch();
      return;
    }
    if (event.key === "h" || event.key === "H") {
      event.preventDefault();
      this.#openFindReplace();
      return;
    }
    // composedPath()[0] is the real target inside the shadow DOM (e.g. a combobox input).
    const target = event.composedPath()[0] as HTMLElement | null;
    if (target instanceof HTMLElement && target.closest("input, textarea, docen-ribbon-combobox"))
      return;
    const key = event.key;
    if (key === "+" || key === "=") {
      event.preventDefault();
      this.#setZoom(this.#zoom + 10);
    } else if (key === "-" || key === "_") {
      event.preventDefault();
      this.#setZoom(this.#zoom - 10);
    } else if (key === "0") {
      event.preventDefault();
      this.#setZoom(100);
    }
  };

  /** The scroll container for the editor — <docen-canvas> (overflow:auto).
   *  TableOfContents uses it to track which heading is active on scroll. */
  #scrollParent(): HTMLElement | Window {
    return this.shadowRoot?.querySelector("docen-canvas") ?? window;
  }

  /** TableOfContents.onUpdate → <docen-outline>. Cache the anchors (so an
   *  outline click resolves to a position) and rebuild the nested tree. */
  #renderOutline(anchors: readonly TocAnchor[]): void {
    this.#anchors = anchors;
    const outline = this.shadowRoot?.querySelector("docen-outline");
    if (!outline) return;
    // Fingerprint only what the pane shows (id/level/title). `pos` moves on
    // every pagination re-flow but never changes the outline, so excluding it
    // avoids rebuilding — and flickering — the fluent-tree each pass. Built
    // from per-anchor arrays rather than the serialized tree, so object key
    // order is irrelevant (no dependency on buildOutlineTree's literal field
    // order, unlike a plain JSON.stringify(tree) comparison).
    const sig = anchors
      .map((a) => JSON.stringify([a.id, a.originalLevel, a.textContent]))
      .join("\n");
    if (this.#outlineSig === sig) return;
    this.#outlineSig = sig;
    outline.setAttribute("items", JSON.stringify(buildOutlineTree(anchors)));
  }

  /** Outline click → select the heading at its position and scroll it into view. */
  readonly #onOutlineSelect = (event: CustomEvent<{ id?: string }>): void => {
    const id = event.detail?.id;
    const editor = this.#editor;
    if (!id || !editor) return;
    const anchor = this.#anchors.find((a) => a.id === id);
    if (!anchor) return;
    editor.chain().focus().setTextSelection(anchor.pos).run();
    editor.view.dispatch(editor.view.state.tr.scrollIntoView());
  };

  /** navigation:search → set the active query; matches highlight live. */
  readonly #onSearch = (event: CustomEvent<{ query?: string }>): void => {
    const editor = this.#editor;
    if (!editor) return;
    const query = new SearchQuery({ search: event.detail?.query ?? "", caseSensitive: false });
    editor.view.dispatch(setSearchState(editor.state.tr, query));
    this.#updateSearchResults();
  };

  /** navigation:find → jump to the next/previous match (prosemirror-search). */
  readonly #onFind = (event: CustomEvent<{ direction: "next" | "prev" }>): void => {
    const editor = this.#editor;
    if (!editor) return;
    (event.detail.direction === "prev" ? findPrev : findNext)(editor.state, editor.view.dispatch);
  };

  /** Stamp the Results slot with the live match list — each hit rendered with
   *  surrounding context and a data-from/to for click-to-jump (Word's Results
   *  pane lists every match with context, not just a count). */
  #updateSearchResults(): void {
    const editor = this.#editor;
    const slot = this.shadowRoot?.querySelector(".search-results");
    if (!slot) return;
    const decos = editor ? getMatchHighlights(editor.state).find() : [];
    slot.replaceChildren();
    const header = document.createElement("div");
    header.className = "result-count";
    header.textContent =
      decos.length > 0
        ? `${decos.length} ${t("search.matches", this)}`
        : t("search.noResults", this);
    slot.append(header);
    if (!editor || decos.length === 0) return;
    const doc = editor.state.doc;
    const RADIUS = 24;
    for (const deco of decos) {
      const { from, to } = deco as { from: number; to: number };
      const before = doc.textBetween(Math.max(0, from - RADIUS), from, " ");
      const after = doc.textBetween(to, Math.min(doc.content.size, to + RADIUS), " ");
      const item = document.createElement("button");
      item.type = "button";
      item.className = "result-item";
      item.dataset.from = String(from);
      item.dataset.to = String(to);
      if (before) {
        const span = document.createElement("span");
        span.textContent = `…${before}`;
        item.append(span);
      }
      const hit = document.createElement("mark");
      hit.textContent = doc.textBetween(from, to, " ");
      item.append(hit);
      if (after) {
        const span = document.createElement("span");
        span.textContent = `${after}…`;
        item.append(span);
      }
      slot.append(item);
    }
  }

  /** Click a Results entry → select that match range and scroll it into view. */
  readonly #onSearchResultClick = (event: Event): void => {
    const editor = this.#editor;
    if (!editor) return;
    const item = (event.target as HTMLElement | null)?.closest(".result-item");
    if (!(item instanceof HTMLElement)) return;
    const from = Number(item.dataset.from);
    const to = Number(item.dataset.to);
    if (!Number.isFinite(from) || !Number.isFinite(to)) return;
    editor.chain().focus().setTextSelection({ from, to }).run();
    editor.view.dispatch(editor.view.state.tr.scrollIntoView());
  };

  /** Ctrl+F → open the nav pane and focus its search box (Word behavior). */
  #openSearch(): void {
    const taskPane = this.shadowRoot?.querySelector('docen-task-pane[position="start"]') as
      | (HTMLElement & { open: boolean })
      | null;
    if (taskPane) taskPane.open = true;
    const input = this.shadowRoot
      ?.querySelector("docen-navigation-pane")
      ?.shadowRoot?.querySelector("[part='search-input']") as
      | (HTMLElement & { select: () => void })
      | null;
    input?.focus();
    input?.select?.();
  }

  /** Ctrl+H / ribbon Replace → open the Find & Replace dialog. */
  #openFindReplace(): void {
    const dialog = this.shadowRoot?.querySelector("docen-find-replace-dialog") as
      | (HTMLElement & { show: () => void })
      | null;
    dialog?.show();
  }

  /** find-replace:action → drive prosemirror-search (query highlights, find-next,
   *  replace-next = replace + advance, replace-all). Each action re-stamps the
   *  query so Find/Replace/options are always current. */
  readonly #onFindReplace = (
    event: CustomEvent<{
      action: string;
      find: string;
      replace: string;
      caseSensitive: boolean;
      wholeWord: boolean;
    }>,
  ): void => {
    const editor = this.#editor;
    if (!editor) return;
    const { action, find, replace, caseSensitive, wholeWord } = event.detail ?? {};
    const query = new SearchQuery({ search: find, replace, caseSensitive, wholeWord });
    editor.view.dispatch(setSearchState(editor.state.tr, query));
    if (action === "find-next") findNext(editor.state, editor.view.dispatch);
    else if (action === "replace-next") replaceNext(editor.state, editor.view.dispatch);
    else if (action === "replace-all") replaceAll(editor.state, editor.view.dispatch);
  };

  /** Paste from the system clipboard as plain text. navigator.clipboard is the
   *  reliable path; execCommand("paste") is the fallback (often blocked). */
  async #paste(): Promise<void> {
    const editor = this.#editor;
    if (!editor) return;
    let text: string | null = null;
    try {
      text = await navigator.clipboard.readText();
    } catch {
      editor.commands.focus();
      document.execCommand("paste");
      return;
    }
    if (text) editor.chain().focus().insertContent(text).run();
  }

  /** Editing → Select menu. "all" uses the official selectAll() command (an
   *  AllSelection that crosses page isolating boundaries); "objects"/"similar"
   *  are placeholders. */
  #select(value?: string): void {
    const editor = this.#editor;
    if (!editor) return;
    if ((value ?? "all") !== "all") return;
    editor.chain().focus().selectAll().run();
  }

  /** Editing → Find drop-down → Go To: prompt for a page number and move the
   *  caret to that page, scrolling it into view. */
  #goToPage(): void {
    const editor = this.#editor;
    if (!editor) return;
    const input = window.prompt(t("ribbon.opt.go-to-prompt", this));
    if (input == null) return;
    const page = parseInt(input, 10);
    if (!Number.isFinite(page) || page < 1) return;
    let count = 0;
    let target = -1;
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === "page") {
        count++;
        if (count === page) {
          target = pos + 1;
          return false;
        }
      }
      return true;
    });
    if (target < 0) return;
    editor.chain().focus().setTextSelection(target).run();
    editor.view.dispatch(editor.view.state.tr.scrollIntoView());
  }

  /** Format Painter: on first click, capture the current selection's marks and
   *  arm a one-shot pointerup listener; the next non-empty selection receives
   *  those marks and disarms the painter. A second click cancels. */
  #toggleFormatPainter(): void {
    if (this.#painterMarks) {
      this.#stopFormatPainter();
      return;
    }
    const editor = this.#editor;
    if (!editor || editor.state.selection.empty) return;
    this.#painterMarks = editor.state.selection.$from.marks();
    this.toggleAttribute("format-painter", true);
    const dom = editor.view.dom;
    const onUp = (): void => {
      const ed = this.#editor;
      if (!ed) return;
      const { from, to, empty } = ed.state.selection;
      if (!empty && this.#painterMarks) {
        const tr = ed.state.tr;
        for (const mark of this.#painterMarks) tr.addMark(from, to, mark);
        ed.view.dispatch(tr);
      }
      this.#stopFormatPainter();
    };
    dom.addEventListener("pointerup", onUp, { once: true });
    this.#painterOff = () => dom.removeEventListener("pointerup", onUp);
  }

  #stopFormatPainter(): void {
    this.#painterMarks = null;
    this.removeAttribute("format-painter");
    this.#painterOff?.();
    this.#painterOff = undefined;
  }

  /** Mirror the font name / size and paragraph style at the caret into the
   *  ribbon comboboxes — Word behavior: the boxes report the formatting at the
   *  cursor, not a fixed default. Re-runs on every editor transaction (caret
   *  moves, marks change). Font/size read the browser-resolved computed style
   *  so the full style inheritance chain (direct run props → paragraph-style →
   *  document defaults) is reflected. */
  #setupFontSync(): void {
    const editor = this.#editor;
    if (!editor) return;
    const sync = (): void => {
      this.#syncFontControls();
      this.#syncStyleControl();
      this.#updateStatus();
    };
    editor.on("transaction", sync);
    sync();
    this.#fontSyncCleanup = (): void => {
      editor.off("transaction", sync);
    };
  }

  #syncFontControls(): void {
    const editor = this.#editor;
    if (!editor) return;
    const fontDisplay = this.#effectiveFontAt(editor);
    const sizeDisplay = this.#effectiveSizeAt(editor);
    const fontCb = this.shadowRoot?.querySelector<HTMLElement>(
      'docen-ribbon-combobox[event="font-name"]',
    );
    const sizeCb = this.shadowRoot?.querySelector<HTMLElement>(
      'docen-ribbon-combobox[event="font-size"]',
    );
    if (fontCb && fontCb.getAttribute("value") !== fontDisplay) {
      fontCb.setAttribute("value", fontDisplay);
    }
    if (sizeCb && sizeCb.getAttribute("value") !== sizeDisplay) {
      sizeCb.setAttribute("value", sizeDisplay);
    }
  }

  /** The font name at the caret, resolved through the style inheritance chain
   *  (direct run props → paragraph style → basedOn → document defaults) in the
   *  document's own units — no px conversion. */
  #effectiveFontAt(editor: Editor): string {
    const direct = editor.getAttributes("textStyle");
    const { font } = effectiveRunProps(
      this.#docStyles(editor),
      this.#currentStyleId(editor),
      direct,
    );
    return font ?? "";
  }

  /** The font size at the caret in points (Word's unit), resolved through the
   *  style inheritance chain — no px conversion. */
  #effectiveSizeAt(editor: Editor): string {
    const direct = editor.getAttributes("textStyle");
    const { size } = effectiveRunProps(
      this.#docStyles(editor),
      this.#currentStyleId(editor),
      direct,
    );
    return size != null ? String(size) : "";
  }

  /** The loaded document's styles model (doc.attrs.styles), or null. */
  #docStyles(editor: Editor): StylesOptions | null {
    return (editor.state.doc.attrs?.styles as StylesOptions | undefined) ?? null;
  }

  /** The paragraph-style id at the caret (HeadingLevel literal for headings,
   *  pStyle id for paragraphs). */
  #currentStyleId(editor: Editor): string | null {
    const h = editor.getAttributes("heading") as { styleId?: unknown };
    const p = editor.getAttributes("paragraph") as { styleId?: unknown };
    if (typeof h?.styleId === "string" && h.styleId) return h.styleId;
    if (typeof p?.styleId === "string" && p.styleId) return p.styleId;
    return null;
  }

  /** Mirror the paragraph style at the caret into the Styles gallery combobox —
   *  its value is the current paragraph's styleId (the HeadingLevel literal for
   *  headings, the pStyle id for paragraphs, or "Normal" when the paragraph
   *  carries none). The combobox matches the value against its gallery items to
   *  show the style's display name. */
  #syncStyleControl(): void {
    const editor = this.#editor;
    if (!editor) return;
    const headingAttrs = editor.getAttributes("heading") as { styleId?: unknown };
    const paraAttrs = editor.getAttributes("paragraph") as { styleId?: unknown };
    const styleId =
      typeof headingAttrs?.styleId === "string" && headingAttrs.styleId
        ? headingAttrs.styleId
        : typeof paraAttrs?.styleId === "string" && paraAttrs.styleId
          ? paraAttrs.styleId
          : "";
    const value = styleId || "Normal";
    const cb = this.shadowRoot?.querySelector<HTMLElement>('docen-ribbon-combobox[event="style"]');
    if (cb && cb.getAttribute("value") !== value) cb.setAttribute("value", value);
  }

  async connectedCallback(): Promise<void> {
    if (!this.shadowRoot) {
      this.attachShadow({ mode: "open" }).innerHTML = TEMPLATE;
    }
    registerComponents();
    applyTheme("light");

    this.#fileInput = this.shadowRoot!.querySelector<HTMLInputElement>("#file-input")!;
    this.#imageInput = this.shadowRoot!.querySelector<HTMLInputElement>("#image-input")!;
    this.#renderChrome();
    this.#applyChromeVisibility();
    this.#setupZoomControls();

    const page = this.shadowRoot!.querySelector<HTMLDivElement>(".docen-pages");
    if (!page) return;

    // Fonts must be loaded before pagination measures, else Pretext drifts
    // from the browser's actual line layout (see rendering-engine-choices).
    await document.fonts?.ready;

    // Wrap the initial content into doc > page+ (the editing schema). The
    // page node never enters DOCX — wrapPages/unwrapPages bridge it at the
    // editor layer, so DOCX round-trip stays transparent.
    const contentAttr = this.getAttribute("content");
    this.#editor = createDocxEditor({
      element: page,
      content: wrapPages(contentAttr ? parseHTML(contentAttr) : undefined),
      // Spellcheck defaults OFF — Chromium's spellcheck is a major perf cost on
      // large documents (ProseMirror community-confirmed). Opt in via the
      // Review ribbon's spell-check button (spellcheck="true" attribute).
      spellcheck: this.getAttribute("spellcheck") === "true",
      editable: this.getAttribute("editable") !== "false",
      // PageDocument overrides the doc schema to doc > page+; Page is the
      // fixed-height sheet node; PagePlugin physically reflows blocks across
      // pages so nothing overflows a sheet (C-route — see CLAUDE.md).
      extensions: [
        PageDocument,
        Page,
        // ImageCap scales over-wide images down to the section content width
        // (Word behavior) and runs before PagePlugin so the reflow measures the
        // capped dimensions, not the pre-cap overflow.
        ImageCap,
        PagePlugin,
        FontMetricDecoration,
        SplitTable,
        SplitTableRow,
        // Paragraph/heading split support: adds editor-only splitGroup/splitPart
        // attrs so the paginator can split a tall paragraph across pages at a
        // line boundary (head on the current page, tail on the next). Both
        // halves share the splitGroup id; unwrapPages merges them on export.
        SplitMarks,
        // TableOfContents injects id/data-toc-id on each heading and reports the
        // anchor list via onUpdate → <docen-outline>. scrollParent is the canvas
        // (our scroll surface), not the window.
        TableOfContents.configure({
          scrollParent: () => this.#scrollParent(),
          onUpdate: (anchors) => this.#renderOutline(anchors as TocAnchor[]),
        }),
        // Search (prosemirror-search) — stores the active query and highlights
        // its matches; driven by the nav-pane search box (#onSearch).
        Search,
        // pageBreak NodeView — renders a Fluent divider with a centered label
        // while show-marks is on. Schema comes from the engine's PageBreak;
        // this only overrides the editor rendering.
        PageBreakView,
        // Centralized MS Office editing keymap (Ctrl+Enter page break, etc.) —
        // see extensions/keymap.ts. Outranks HardBreak via priority.
        DocenKeymap,
        // sectionBreak widget — a section boundary is paragraph attrs (not a
        // node), so it has no NodeView; a widget decoration paints the Fluent
        // divider marker after each section-carrying paragraph.
        SectionBreakMarks,
        Placeholder.configure({
          // A function (not a string) so the prompt re-reads the active locale
          // each time the decoration set is rebuilt — a locale switch then
          // refreshes the placeholder text without recreating the extension.
          placeholder: () => t("editor.placeholder"),
          // C-route nests paragraphs inside a non-textblock page node
          // (doc > page > p). Placeholder's decoration walk returns false at the
          // page unless includeChildren is set, so the prompt never reaches the
          // paragraphs — turn it on so the empty first paragraph gets the
          // is-empty class + data-placeholder attribute our CSS paints.
          includeChildren: true,
        }),
        // Editing-behavior set — the engine carries schema only.
        UndoRedo,
        Dropcursor,
        Gapcursor,
        TrailingNode,
        ListKeymap,
        CharacterCount.configure({
          // Word-style count: each CJK character counts as one, non-CJK runs
          // split on whitespace — matches Word for mixed CJK/Latin (the default
          // split(' ').length counts a whole CJK paragraph as a single word).
          wordCounter: (text: string): number => {
            const cjkRe = /[一-鿿぀-ヿ가-힯]/g;
            const cjk = (text.match(cjkRe) ?? []).length;
            const western = text.replace(cjkRe, " ").split(/\s+/).filter(Boolean).length;
            return cjk + western;
          },
          // Count characters by grapheme cluster so emoji / combining marks /
          // surrogate pairs count as one (default text.length undercounts them).
          textCounter: (text: string): number => {
            const seg = new Intl.Segmenter("en", { granularity: "grapheme" });
            let n = 0;
            for (const _ of seg.segment(text)) n++;
            return n;
          },
        }),
        Focus,
        Selection,
      ],
    });
    this.#applyDocStyles();

    // Mirror the caret's font/size into the ribbon comboboxes (Word behavior).
    this.#setupFontSync();
    // Stamp the status bar (page count / caret page / zoom) once laid out.
    this.#updateStatus();

    // Default page setup (Word defaults): A4 portrait + Normal margins. The
    // canvas already defaults to 210×297; apply the margin preset so the
    // content box matches Word and pagination measures correctly.
    this.#setMargins("normal");

    // command = ribbon buttons; change = menu items + auto-save switch. Listen
    // on the shadow root so non-composed Fluent events (menu-item "change")
    // reach us, not just composed ones (ribbon "command").
    this.shadowRoot!.addEventListener("command", this.#onCommand as EventListener);
    this.shadowRoot!.addEventListener("change", this.#onChange as EventListener);
    this.#fileInput.addEventListener("change", this.#onFileChange);
    this.#imageInput.addEventListener("change", this.#onImageChange);
    // Outline (Headings tab) → jump to the clicked heading.
    this.shadowRoot!.querySelector("docen-outline")?.addEventListener(
      "outline:select",
      this.#onOutlineSelect as EventListener,
    );
    // Nav-pane search → Find (live highlight, next/prev, results count).
    this.addEventListener("navigation:search", this.#onSearch as EventListener);
    this.addEventListener("navigation:find", this.#onFind as EventListener);
    // Click a Results entry → jump to that match (delegated on the container).
    this.shadowRoot!.querySelector(".search-results")?.addEventListener(
      "click",
      this.#onSearchResultClick as EventListener,
    );
    // Find & Replace dialog → Replace / Replace All (prosemirror-search).
    this.shadowRoot!.querySelector("docen-find-replace-dialog")?.addEventListener(
      "find-replace:action",
      this.#onFindReplace as EventListener,
    );

    // Re-render header + ribbon when the page locale (<html lang>) changes.
    this.#unobserveLang = observeLang(() => this.#renderChrome());

    // Ribbon Display Options → drive browser fullscreen + status-bar hide.
    // auto-hide = Full Screen (Office); any other mode exits it.
    const ribbon = this.shadowRoot!.querySelector("docen-ribbon");
    const workspace = this.shadowRoot!.querySelector("docen-workspace");
    if (ribbon && workspace) {
      ribbon.addEventListener("ribbon-mode-change", (event) => {
        const mode = (event as CustomEvent<{ mode: string }>).detail.mode;
        if (mode === "auto-hide") {
          void this.requestFullscreen?.().catch(() => {});
          workspace.setAttribute("data-fullscreen", "");
        } else {
          if (document.fullscreenElement) void document.exitFullscreen?.().catch(() => {});
          workspace.removeAttribute("data-fullscreen");
        }
      });
    }
    // Emit docen:change on every content change (autosave driver) and docen:ready
    // once the editor is live — both bubble out so a host can react.
    this.#editor?.on("transaction", this.#onTransaction);
    document.addEventListener("fullscreenchange", this.#onFullscreenChange);
    this.addEventListener("keydown", this.#onZoomKey);
    this.dispatchEvent(new CustomEvent("docen:ready", { bubbles: true, composed: true }));
  }

  disconnectedCallback(): void {
    this.shadowRoot?.removeEventListener("command", this.#onCommand as EventListener);
    this.shadowRoot?.removeEventListener("change", this.#onChange as EventListener);
    this.#fileInput?.removeEventListener("change", this.#onFileChange);
    this.#imageInput?.removeEventListener("change", this.#onImageChange);
    this.shadowRoot
      ?.querySelector("docen-outline")
      ?.removeEventListener("outline:select", this.#onOutlineSelect as EventListener);
    this.removeEventListener("navigation:search", this.#onSearch as EventListener);
    this.removeEventListener("navigation:find", this.#onFind as EventListener);
    this.shadowRoot
      ?.querySelector(".search-results")
      ?.removeEventListener("click", this.#onSearchResultClick as EventListener);
    this.shadowRoot
      ?.querySelector("docen-find-replace-dialog")
      ?.removeEventListener("find-replace:action", this.#onFindReplace as EventListener);
    this.#unobserveLang?.();
    this.#editor?.off("transaction", this.#onTransaction);
    document.removeEventListener("fullscreenchange", this.#onFullscreenChange);
    this.removeEventListener("keydown", this.#onZoomKey);
    this.#fontSyncCleanup?.();
    this.#editor?.destroy();
  }

  /** App header markup — i18n labels plus the host-supplied `user` / `filename`
   *  (not translated: identity and the file name come from the app via
   *  attributes, not the locale table). The auto-save label sits to the left of
   *  its toggle (Word layout), so the switch carries only an aria-label. */
  #renderHeader(): string {
    const user = this.getAttribute("user") ?? "";
    const avatar = this.getAttribute("avatar") ?? "";
    const filename = this.getAttribute("filename") ?? t("header.doc-name");
    const initial = user.trim().charAt(0).toUpperCase();
    const avatarMarkup = avatar
      ? `<img class="avatar avatar-img" src="${avatar}" alt="" />`
      : initial
        ? `<span class="avatar">${initial}</span>`
        : "";
    const autosave = t("header.autosave");
    return `
          <div slot="start" style="display:flex;align-items:center;gap:4px">
            <span style="font-weight:600;font-size:13px;padding-inline:6px">${t("header.brand")}</span>
            <span class="autosave-label">${autosave}</span>
            <!-- auto-save is skeleton-only — disabled (greyed, non-interactive)
                 until the feature is wired; the autosave case in onChange is a
                 no-op. Remove disabled (and re-add checked) when it lands. -->
            <fluent-switch data-event="autosave" disabled aria-label="${autosave}"></fluent-switch>
            <docen-ribbon-button icon="save" label="${t("header.save")}" event="save" icon-only></docen-ribbon-button>
            <docen-ribbon-button icon="undo" label="${t("header.undo")}" event="undo" icon-only></docen-ribbon-button>
            <docen-ribbon-button icon="redo" label="${t("header.redo")}" event="redo" icon-only></docen-ribbon-button>
            <fluent-menu>
              <fluent-menu-button slot="trigger" appearance="subtle">${filename}</fluent-menu-button>
              <fluent-menu-list>
                <fluent-menu-item data-event="new">${t("header.new")}</fluent-menu-item>
                <fluent-menu-item data-event="open">${t("header.open")}</fluent-menu-item>
                <fluent-menu-item data-event="save-as">${t("header.save-as")}</fluent-menu-item>
                <fluent-menu-item data-event="print">${t("header.print")}</fluent-menu-item>
              </fluent-menu-list>
            </fluent-menu>
          </div>
          <fluent-text-input slot="search" placeholder="${t("header.search")}"></fluent-text-input>
          <div slot="end" style="display:flex;align-items:center;gap:4px">
            <fluent-menu>
              <fluent-menu-button slot="trigger" appearance="subtle">${avatarMarkup}${user}</fluent-menu-button>
              <fluent-menu-list>
                <fluent-menu-item data-event="lang-zh">${t("header.lang.zh")}</fluent-menu-item>
                <fluent-menu-item data-event="lang-en">${t("header.lang.en")}</fluent-menu-item>
              </fluent-menu-list>
            </fluent-menu>
            ${this.hasAttribute("closable") ? `<docen-ribbon-button icon="close" label="${t("header.close")}" event="close" icon-only></docen-ribbon-button>` : ""}
          </div>`;
  }

  /** Stamp the header + ribbon markup for the active locale (re-run on lang change). */
  #renderChrome(): void {
    const root = this.shadowRoot;
    if (!root) return;
    const styles = this.#editor?.state.doc.attrs?.styles ?? null;
    root.querySelector("docen-app-header")!.innerHTML = this.#renderHeader();
    root.querySelector("docen-ribbon")!.innerHTML = buildRibbonInnerHTML(styles, {
      tabs: this.#tabs(),
    });
    this.#applyRibbonGreying();
    this.#renderPanes();
  }

  /** Stamp pane titles + status text for the active locale (re-run on lang change). */
  #renderPanes(): void {
    const root = this.shadowRoot;
    if (!root) return;
    const navPane = root.querySelector('docen-task-pane[position="start"]');
    if (navPane) navPane.setAttribute("title", t("pane.navigation"));
    const propsPane = root.querySelector('docen-task-pane[position="end"]');
    if (propsPane) propsPane.setAttribute("title", t("pane.properties"));
    // Page-break / section-break divider labels live in the editor view
    // (NodeView / widget decoration), not the ribbon — update them alongside
    // the chrome so a locale change relabels them.
    root.querySelectorAll<HTMLElement>("fluent-divider[data-pb]").forEach((d) => {
      d.textContent = t("ribbon.cmd.page-break");
    });
    root.querySelectorAll<HTMLElement>("fluent-divider[data-sb]").forEach((d) => {
      d.textContent = t("ribbon.cmd.section-break");
    });
    // Status bar is dynamic (page count / caret page / zoom) — re-stamp it so a
    // locale change re-localizes the text too.
    this.#updateStatus();
    // The placeholder prompt comes from a closure over t(); its decoration set
    // is only rebuilt on a view update, so dispatch a meta-only transaction to
    // force a rebuild against the now-current locale (empty doc → new prompt).
    const view = this.#editor?.view;
    if (view) view.dispatch(view.state.tr.setMeta("docen-i18n", true));
  }

  /** Apply the chrome-visibility attributes (toolbar/header/status-bar/panes)
   *  by toggling the native HTML `hidden` attribute + the task-pane `open`
   *  property. UI components need no changes — `hidden` is native. */
  #applyChromeVisibility(): void {
    const root = this.shadowRoot;
    if (!root) return;
    const hide = (selector: string, attr: string): void => {
      const el = root.querySelector(selector);
      if (el) (el as HTMLElement).toggleAttribute("hidden", this.getAttribute(attr) === "false");
    };
    hide("docen-ribbon", "toolbar");
    hide("docen-app-header", "header");
    hide('footer[slot="status"]', "status-bar");
    this.#applyPane("start", "navigation-pane");
    this.#applyPane("end", "properties-pane");
  }

  /** Apply one side pane's attribute: "hidden" removes it, "open"/"closed"
   *  toggle its `open` state (the task-pane component observes `open`). */
  #applyPane(side: "start" | "end", attr: string): void {
    const pane = this.shadowRoot?.querySelector(`docen-task-pane[position="${side}"]`) as
      | (HTMLElement & { open: boolean })
      | null;
    if (!pane) return;
    const value = this.getAttribute(attr);
    if (value === "hidden") pane.setAttribute("hidden", "");
    else {
      pane.removeAttribute("hidden");
      if (value === "closed") pane.open = false;
      else if (value === "open") pane.open = true;
    }
  }

  /** Grey out ribbon commands that have no handler (skeleton buttons). Runs
   *  after every ribbon re-stamp; fresh elements start un-disabled, so this is
   *  the single place `disabled` is applied. Only controls that support
   *  `disabled` (button/split-button/toggle-button) are greyed — combobox /
   *  color-picker lack it and live in wired tabs anyway. */
  #applyRibbonGreying(): void {
    const ribbon = this.shadowRoot?.querySelector("docen-ribbon");
    if (!ribbon) return;
    const wired = this.#wiredCommands();
    ribbon
      .querySelectorAll<HTMLElement>(
        "docen-ribbon-button[event], docen-ribbon-split-button[event], docen-ribbon-toggle-button[event]",
      )
      .forEach((el) => {
        const event = el.getAttribute("event");
        if (event && !wired.has(event)) el.setAttribute("disabled", "");
      });
  }

  /** The full set of wired command names (dispatch + locally handled). */
  #wiredCommands(): Set<string> {
    return new Set<string>([...WIRED_DISPATCH, ...LOCAL_HANDLED]);
  }

  /** The `tabs` whitelist, or undefined for all tabs. */
  #tabs(): readonly RibbonTabId[] | undefined {
    return parseTabs(this.getAttribute("tabs"));
  }

  /** Dispatch a cancelable event; returns true when a host preventDefaulted it
   *  (i.e. took over the action). Lets save/open/print/new work out-of-box yet
   *  stay overridable. */
  #emitCancelable(
    name: "docen:save" | "docen:save-as" | "docen:open" | "docen:new" | "docen:print",
  ): boolean {
    const event = new CustomEvent(name, { bubbles: true, composed: true, cancelable: true });
    this.dispatchEvent(event);
    return event.defaultPrevented;
  }

  /** The close (×) button asks the host to close the editor — the component
   *  itself never unmounts (it doesn't know the host's context). Only rendered
   *  when the host sets `closable`. */
  #emitRequestClose(): void {
    this.dispatchEvent(new CustomEvent("docen:request-close", { bubbles: true, composed: true }));
  }

  /** docen:change — fired on every doc-changing transaction (autosave driver,
   *  mirroring OnlyOffice's onDocumentStateChange). Selection-only transactions
   *  are skipped. */
  readonly #onTransaction = (props: { transaction: { docChanged: boolean } }): void => {
    if (props.transaction.docChanged) {
      this.dispatchEvent(
        new CustomEvent("docen:change", { bubbles: true, composed: true, detail: { dirty: true } }),
      );
    }
  };

  /** Toggle a side task pane open/closed (ribbon View → toggle-navigation). */
  #togglePane(side: "start" | "end"): void {
    const pane = this.shadowRoot?.querySelector(`docen-task-pane[position="${side}"]`) as
      | (HTMLElement & { open: boolean })
      | null;
    if (pane) pane.open = !pane.open;
  }

  /** Apply a paper-size preset (a4/letter/…) to the canvas as raw mm, then
   *  re-paginate. Also writes the size into the document-model sectionProperties
   *  (Word stores page setup in the sectPr) so render/measure/image-cap/export
   *  share one geometry source; the canvas attrs are now the rendering fallback
   *  + the zoom surface's page-width source. */
  #setPageSize(value?: string): void {
    const canvas = this.shadowRoot?.querySelector("docen-canvas");
    const size = value ? PAPER_SIZES[value] : undefined;
    if (canvas && size) {
      canvas.setAttribute("page-width", String(size[0]));
      canvas.setAttribute("page-height", String(size[1]));
    }
    if (size) {
      this.#updateSectionGeometry({
        page: {
          size: {
            width: convertMillimetersToTwip(size[0]),
            height: convertMillimetersToTwip(size[1]),
          },
        },
      });
    }
    this.#refreshGeometry();
  }

  /** Apply orientation (portrait/landscape) to the canvas, then re-paginate.
   *  portrait clears the attribute (the canvas default); landscape sets it and
   *  the canvas swaps width/min-height via :host([orientation]). Also writes
   *  orientation onto page.size, deep-merged with the current (or engine-default)
   *  size so resolvePageSize can swap edges for landscape. */
  #setOrientation(value?: string): void {
    const canvas = this.shadowRoot?.querySelector("docen-canvas");
    if (canvas && value) {
      if (value === "landscape") canvas.setAttribute("orientation", "landscape");
      else canvas.removeAttribute("orientation");
    }
    if (value) {
      const cur = (
        this.#editor?.state.doc.attrs as { sectionProperties?: SectionPropertiesOptions }
      )?.sectionProperties?.page?.size;
      const size =
        cur && typeof cur.width === "number" && typeof cur.height === "number"
          ? cur
          : { width: sectionPageSizeDefaults.WIDTH, height: sectionPageSizeDefaults.HEIGHT };
      this.#updateSectionGeometry({
        page: { size: { ...size, orientation: value as "portrait" | "landscape" } },
      });
    }
    this.#refreshGeometry();
  }

  /** Apply a margin preset (normal/narrow/…) to the canvas as a CSS padding
   *  list, then re-paginate. Also writes the margins into the document-model
   *  sectionProperties so a page-setup change actually re-caps images and
   *  re-renders (the canvas CSS alone wouldn't, once a sectPr is inlined). */
  #setMargins(value?: string): void {
    const canvas = this.shadowRoot?.querySelector("docen-canvas");
    if (canvas && value && MARGINS[value]) canvas.setAttribute("margin", MARGINS[value]);
    if (value && MARGINS[value]) {
      this.#updateSectionGeometry({ page: { margin: marginTwipsFromCss(MARGINS[value]) } });
    }
    this.#refreshGeometry();
  }

  /** Deep-merge a sectionProperties patch into the CURRENT section's sectPr and
   *  dispatch it — Word's "this section" semantics. The current section is the
   *  one holding the caret: its sectPr rides on its last paragraph (the first
   *  section-carrying paragraph at/after the caret), or, when the caret is in the
   *  final section, on doc.attrs.sectionProperties (the body-level sectPr).
   *  Reflow re-stamps each page's geometry from its section's sectPr, so an edit
   *  only affects the caret's section — multi-section docs keep per-section page
   *  setups, and render/measure/image-cap/export all see the change. */
  #updateSectionGeometry(patch: SectionPropertiesOptions): void {
    const editor = this.#editor;
    if (!editor) return;
    const { doc, tr } = editor.state;
    const from = editor.state.selection.from;
    // First section-carrying paragraph at/after the caret = the current
    // section's last paragraph (OOXML: its sectPr ends that section).
    let targetPos: number | null = null;
    doc.descendants((node, nodePos) => {
      if (targetPos != null || nodePos < from) return true;
      if (
        node.type.name === "paragraph" &&
        (node.attrs as { sectionProperties?: unknown }).sectionProperties != null
      ) {
        targetPos = nodePos;
        return false;
      }
      return true;
    });
    if (targetPos != null) {
      const node = doc.nodeAt(targetPos);
      if (node) {
        const cur = (node.attrs as { sectionProperties?: SectionPropertiesOptions })
          .sectionProperties;
        tr.setNodeMarkup(targetPos, undefined, {
          ...node.attrs,
          sectionProperties: mergeSectionProperties(cur, patch),
        });
      }
    } else {
      // Caret in the final section (no section-carrying paragraph at/after it) —
      // its sectPr is body-level (doc.attrs.sectionProperties).
      const cur = (doc.attrs as { sectionProperties?: SectionPropertiesOptions }).sectionProperties;
      tr.setDocAttribute("sectionProperties", mergeSectionProperties(cur, patch));
    }
    editor.view.dispatch(tr);
  }

  /** Re-paginate after a page-size / orientation / margin change. RAF so the
   *  new canvas geometry applies first; PagePlugin re-reads the page height
   *  from the DOM and re-solves the breaks. */
  #refreshGeometry(): void {
    requestAnimationFrame(() => {
      if (this.#editor) pageStorageOf(this.#editor).repaginate();
    });
  }

  /** Apply a zoom level (percent, clamped 10–500) to the canvas and refresh the
   *  status bar. CSS `zoom` rescales the pages and reflows the scroll surface. */
  #setZoom(pct: number): void {
    this.#zoom = Math.max(10, Math.min(500, Math.round(pct)));
    this.shadowRoot?.querySelector("docen-canvas")?.setAttribute("zoom", String(this.#zoom));
    this.#updateStatus();
  }

  /** Wire the status-bar zoom cluster (Word's bottom-right control): dragging
   *  the slider sets the zoom live, and the minus / plus buttons step by 10%.
   *  Bound once — the footer is static template DOM, never re-stamped. */
  #setupZoomControls(): void {
    const root = this.shadowRoot;
    if (!root) return;
    const slider = root.querySelector<HTMLInputElement>(".docen-zoom-slider");
    if (slider) {
      slider.addEventListener("input", () => {
        this.#setZoom(Number(slider.value));
      });
    }
    root.querySelectorAll<HTMLElement>(".docen-zoom-step").forEach((btn) => {
      btn.addEventListener("click", () => {
        const step = Number(btn.getAttribute("data-zoom-step") ?? "0");
        this.#setZoom(this.#zoom + step * 10);
      });
    });
  }

  /** Resolve a ribbon zoom preset to a percent. Numeric presets ("200", "100",
   *  "75", "50") map directly to that zoom level; "page-width" scales the page
   *  to fill the canvas width (mm → px @96dpi). */
  #zoomPreset(preset: string): void {
    if (/^\d+$/.test(preset)) return this.#setZoom(Number(preset));
    if (preset !== "page-width") return;
    const canvas = this.shadowRoot?.querySelector("docen-canvas");
    if (!canvas) return;
    const MM2PX = 96 / 25.4;
    const pw = parseFloat(canvas.getAttribute("page-width") ?? "210") * MM2PX;
    this.#setZoom((canvas.clientWidth / pw) * 100);
  }

  /** Refresh the status bar to mirror Word's bottom row: the left cluster is
   *  the caret's section, then "Page X of Y", then the word count; the right
   *  cluster is the zoom slider value + percent. Runs on every transaction
   *  (caret moves, a re-flow changes the page count) and on zoom / locale change.
   *
   *  The section number is O(pages): each page node carries its section's
   *  sectionProperties attrs (one shared reference across a section's pages),
   *  so the count increments only when that reference changes — no paragraph
   *  walk. The word count is cached by doc nodeSize so caret moves skip
   *  re-walking the full document (CharacterCount.words() regexes all text). */
  #updateStatus(): void {
    const root = this.shadowRoot;
    if (!root) return;
    const sectionEl = root.querySelector<HTMLElement>(".docen-status-section");
    const pagesEl = root.querySelector<HTMLElement>(".docen-status-pages");
    const wordsEl = root.querySelector<HTMLElement>(".docen-status-words");
    const sliderEl = root.querySelector<HTMLInputElement>(".docen-zoom-slider");
    const pctEl = root.querySelector<HTMLElement>(".docen-zoom-pct");
    const editor = this.#editor;
    let page = 0;
    let total = 0;
    let section = 1;
    if (editor) {
      const from = editor.state.selection.from;
      let prevSp: unknown;
      let sectionCount = 0;
      let firstPage = true;
      editor.state.doc.forEach((node, offset) => {
        if (node.type.name !== "page") return;
        total++;
        const sp = (node.attrs as { sectionProperties?: unknown }).sectionProperties ?? null;
        // A section's pages share one sectionProperties reference, so the count
        // rises only at a real section boundary (or on the very first page).
        if (firstPage || sp !== prevSp) sectionCount++;
        firstPage = false;
        prevSp = sp;
        if (page === 0 && from > offset && from <= offset + node.nodeSize) {
          page = total;
          section = sectionCount;
        }
      });
    }
    if (sectionEl) {
      sectionEl.textContent = t("status.section").replace("{n}", String(section));
    }
    if (pagesEl) {
      pagesEl.textContent = t("status.page-of")
        .replace("{page}", String(page || 1))
        .replace("{total}", String(total || 1));
    }
    if (wordsEl) {
      const docSize = editor?.state.doc.nodeSize ?? 0;
      if (docSize !== this.#lastDocSize) {
        const cc = editor?.storage.characterCount as { words?: () => number } | undefined;
        this.#lastWords = cc?.words?.() ?? 0;
        this.#lastDocSize = docSize;
      }
      wordsEl.textContent = t("status.words").replace("{n}", String(this.#lastWords));
    }
    // Sync the slider without retriggering its own input handler — only write
    // when the value drifted (keyboard / ribbon zoom changed it out of band).
    if (sliderEl && Number(sliderEl.value) !== this.#zoom) sliderEl.value = String(this.#zoom);
    if (pctEl) pctEl.textContent = `${this.#zoom}%`;
  }

  readonly #onCommand = (event: CustomEvent<{ event?: string; value?: string }>): void => {
    const { event: name, value } = event.detail ?? {};
    if (typeof name !== "string") return;
    // UI chrome actions are handled locally and need no Tiptap editor.
    if (name === "toggle-navigation") {
      this.#togglePane("start");
      return;
    }
    // Close (×) — only rendered when `closable`; ask the host to close.
    if (name === "close") {
      this.#emitRequestClose();
      return;
    }
    // Find (ribbon Home → Editing → Find, or Ctrl+F) → open the nav-pane search.
    if (name === "search") {
      // Find drop-down → Go To jumps to a page; the main button and Find
      // open the nav-pane search box.
      if (value === "go-to") this.#goToPage();
      else this.#openSearch();
      return;
    }
    // Replace (ribbon Home → Editing → Replace, or Ctrl+H) → Find & Replace dialog.
    if (name === "replace") {
      this.#openFindReplace();
      return;
    }
    // Page setup actions are handled locally (they change the canvas/page, not
    // the Tiptap doc) and need no editor command.
    if (name === "page-size") {
      this.#setPageSize(value);
      return;
    }
    if (name === "orientation") {
      this.#setOrientation(value);
      return;
    }
    if (name === "margins") {
      this.#setMargins(value);
      return;
    }
    // Zoom is a canvas action (not a Tiptap command): step in, or apply a
    // preset from the split menu (200/100/75/50/page-width); the split's
    // main button sets 100%.
    if (name === "zoom") {
      this.#setZoom(this.#zoom + 10);
      return;
    }
    if (name === "zoom-100") {
      if (value) this.#zoomPreset(value);
      else this.#setZoom(100);
      return;
    }
    if (!this.#editor) return;
    // "save" is a document action, not a Tiptap command — handle locally,
    // unless the host took over via docen:save (preventDefault).
    if (name === "save") {
      if (!this.#emitCancelable("docen:save")) void this.#saveAs();
      return;
    }
    // Picture needs a file picker — open it, then insert the chosen image.
    if (name === "insert-picture") {
      this.#imageInput?.click();
      return;
    }
    // Formatting marks toggle — FormattingMarks paints the paragraph mark via a
    // widget decoration; the host [show-marks] attr drives the page-break
    // divider CSS.
    if (name === "show-marks") {
      this.toggleAttribute("show-marks");
      this.#editor?.commands.toggleFormattingMarks();
      return;
    }
    // Clipboard — execCommand copy/cut acts on the editor's DOM selection;
    // paste reads the system clipboard (contenteditable execCommand paste is
    // blocked in most browsers).
    if (name === "copy" || name === "cut") {
      this.#editor.commands.focus();
      document.execCommand(name);
      return;
    }
    if (name === "paste") {
      void this.#paste();
      return;
    }
    // Editing → Select: selectAll() spans every page (bypassing page
    // isolating); objects/similar are not yet wired.
    if (name === "select") {
      this.#select(value);
      return;
    }
    // Format Painter — toggle capture/apply of the current run's marks.
    if (name === "format-painter") {
      this.#toggleFormatPainter();
      return;
    }
    dispatchRibbonCommand(this.#editor, name, value);
  };

  /** Menu items and the auto-save switch carry their action in `data-event`. */
  readonly #onChange = (event: Event): void => {
    const name = (event.target as HTMLElement)?.dataset?.event;
    if (!name) return;
    switch (name) {
      case "open":
        // Host can take over via docen:open (preventDefault); else open the picker.
        if (!this.#emitCancelable("docen:open")) this.#fileInput?.click();
        break;
      case "save-as":
        if (!this.#emitCancelable("docen:save-as")) void this.#saveAs();
        break;
      case "print":
        if (!this.#emitCancelable("docen:print")) this.#print();
        break;
      case "new":
        // No built-in "new" — always hand to the host (docen:new).
        this.#emitCancelable("docen:new");
        break;
      case "lang-zh":
        document.documentElement.lang = "zh-CN";
        break;
      case "lang-en":
        document.documentElement.lang = "en";
        break;
      // autosave: skeleton — wired when that feature lands.
    }
  };

  readonly #onFileChange = (event: Event): void => {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (file) void this.openDOCX(file);
    // Reset so picking the same file twice still fires `change`.
    (event.target as HTMLInputElement).value = "";
  };

  /** Insert the picked image as a data URL. Width/height are left unset — the
   *  browser shows natural size, and prepareImages fills them on DOCX export. */
  readonly #onImageChange = (event: Event): void => {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = "";
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (): void => {
      this.#editor
        ?.chain()
        .focus()
        .insertContent({ type: "image", attrs: { src: reader.result } })
        .run();
    };
    reader.readAsDataURL(file);
  };

  /** Save the document as .docx. Uses the native Save As dialog
   *  (showSaveFilePicker) when available so the user picks the location and
   *  name; falls back to a plain download otherwise. The header filename is
   *  updated to match the saved name. */
  async #saveAs(): Promise<void> {
    const buffer = await this.saveDOCX();
    const blob = new Blob([buffer as BlobPart], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    const suggestedName = this.getAttribute("filename")?.trim() || t("header.doc-name");
    const picker = (
      window as unknown as {
        showSaveFilePicker?: (opts: {
          suggestedName?: string;
          types?: Array<{ description?: string; accept: Record<string, string[]> }>;
        }) => Promise<{
          name: string;
          createWritable: () => Promise<{
            write: (data: Blob | BufferSource | string) => Promise<void>;
            close: () => Promise<void>;
          }>;
        }>;
      }
    ).showSaveFilePicker;
    if (picker) {
      try {
        const handle = await picker({
          suggestedName,
          types: [
            {
              description: "Word Document",
              accept: {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [
                  ".docx",
                ],
              },
            },
          ],
        });
        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
        this.setAttribute("filename", handle.name);
        this.#renderChrome();
        return;
      } catch {
        // User cancelled, or the picker is unavailable — fall back to a download.
      }
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = suggestedName;
    a.click();
    URL.revokeObjectURL(url);
  }

  /** Print the document: window.print() — the @media print rules in this
   *  template, the workspace, and the canvas hide the chrome and render only
   *  the page content. */
  #print(): void {
    window.print();
  }

  /** Load a .docx into the editor from a File or a buffer (ArrayBuffer /
   *  Uint8Array). A File also adopts its name as the filename; a bare buffer
   *  carries no name. parseDOCX is synchronous, but this is async so a File's
   *  bytes can be awaited. */
  async openDOCX(input: File | ArrayBuffer | Uint8Array): Promise<void> {
    const buffer = input instanceof File ? await input.arrayBuffer() : input;
    const json = parseDOCX(buffer);
    const editor = this.#editor;
    if (!editor) return;
    // Replace the whole doc node (content + doc-level attrs) — see #loadDoc.
    this.#loadDoc(editor, wrapPages(json));
    this.#applyDocStyles();
    // Paginate once the new document has laid out. setContent renders the DOM
    // synchronously, but the browser commits layout on the NEXT frame —
    // measuring in the same frame reads half-laid-out blocks and the breaks
    // come out wrong (the first page overflows, so the pages look uneven).
    this.#repaginateAfterLoad(editor);
    if (input instanceof File) {
      this.setAttribute("filename", input.name);
      this.#renderChrome();
    }
  }

  /** Re-paginate after loading a document, once its layout has settled.
   *
   *  Two passes: the first rAF runs as soon as the browser has committed layout
   *  for the just-set content (correct for text and tables — the common case,
   *  and fast enough to read as "instant"). Fonts and images load after that
   *  frame and change block heights, so a second pass re-measures once they're
   *  ready. No debounce: an import must paginate within a frame, not after
   *  300ms — the old deferred pass lingered on stale seams for seconds. */
  #repaginateAfterLoad(editor: Editor): void {
    const run = (): void => {
      if (this.#editor === editor && !editor.isDestroyed) pageStorageOf(editor).repaginate();
    };
    requestAnimationFrame(run);
    const fonts = document.fonts?.ready ?? Promise.resolve();
    const imgs = Array.from(editor.view.dom.querySelectorAll<HTMLImageElement>("img"));
    const imgReady = (img: HTMLImageElement): Promise<void> =>
      img.complete
        ? Promise.resolve()
        : new Promise((resolve) => {
            img.addEventListener("load", () => resolve(), { once: true });
            img.addEventListener("error", () => resolve(), { once: true });
          });
    void Promise.all([fonts, ...imgs.map(imgReady)]).then(() => {
      // Fonts loading after the first pass change canvas measureText widths,
      // so the Pretext prepare cache (keyed on text+font+letterSpacing) is now
      // stale — clear it before the second pass re-measures with the real fonts.
      clearMeasureCache();
      requestAnimationFrame(run);
    });
  }

  /** Serialize the current document to a DOCX buffer. */
  async saveDOCX(): Promise<Uint8Array> {
    const json = unwrapPages(this.#editor!.getJSON() as JSONContent);
    const buffer = await generateDOCX(json);
    return buffer as unknown as Uint8Array;
  }

  /** Current document as Tiptap JSON. */
  getJSON(): JSONContent {
    // Unwrap page nodes — external consumers see flat doc > block+ (page
    // nodes are editor-only; transparent in the public API).
    const editor = this.#editor;
    return editor ? unwrapPages(editor.getJSON()) : ({} as JSONContent);
  }

  /** Replace the document with Tiptap JSON. */
  setJSON(json: JSONContent): void {
    // Wrap on the way in — external callers pass flat doc > block+. Tiptap's
    // setContent only swaps content (it drops doc-level attrs like styles/core),
    // so replace the whole doc node via a fresh EditorState to preserve them.
    const editor = this.#editor;
    if (editor) this.#loadDoc(editor, wrapPages(json));
    this.#applyDocStyles();
  }

  /** Replace the whole doc node (content + doc-level attrs) via a fresh
   *  EditorState. Tiptap's setContent only swaps content and drops doc-level
   *  attrs; this carries them (styles/core/sectionProperties). updateState
   *  bypasses appendTransaction/onTransaction, so extensions that react to doc
   *  changes wouldn't wake — dispatch a docChanged tr (re-stamp the first
   *  block's attrs, a no-op visually) to trigger them: TableOfContents injects
   *  heading ids + fires onUpdate, PagePlugin schedules a re-flow. */
  #loadDoc(editor: Editor, doc: JSONContent): void {
    editor.view.updateState(
      EditorState.create({ doc: editor.schema.nodeFromJSON(doc), plugins: editor.state.plugins }),
    );
    if (editor.isDestroyed) return;
    // updateState bypasses appendTransaction, so extensions that react to doc
    // changes wouldn't wake. Dispatch a docChanged tr to fire them: TableOfContents
    // injects heading ids + emits onUpdate, PagePlugin schedules a re-flow. The
    // tr re-stamps the LAST leaf block's OWN attrs — a true no-op (same node,
    // same attrs) — so nothing is clobbered. Targeting the last leaf (not pos 1)
    // sidesteps the page node's `isolating` boundary entirely: the old
    // setNodeMarkup(1) resolved INTO the page's first child and could overwrite
    // its attrs, which is how the first heading lost its styleId (and with it
    // its Heading1 bold/centering) on load.
    const state = editor.state;
    // Collect leaf blocks; the last one (deepest, rightmost) is the re-stamp
    // target. A const array (mutated, not reassigned) keeps TS's control-flow
    // analysis happy — a `let` reassigned inside the callback reads as `never`
    // outside it (TS can't see the callback runs synchronously).
    const leaves: { pos: number; attrs: Record<string, unknown> }[] = [];
    state.doc.nodesBetween(0, state.doc.content.size, (node, pos) => {
      if (node.isText) return;
      if (node.isTextblock || node.isLeaf) {
        leaves.push({ pos, attrs: node.attrs as Record<string, unknown> });
      }
      // Don't descend into textblocks (their text isn't a markup target).
      return node.isTextblock ? false : undefined;
    });
    const last = leaves[leaves.length - 1];
    if (last) {
      editor.view.dispatch(state.tr.setNodeMarkup(last.pos, undefined, last.attrs));
    }
  }

  /** Inject the document's named styles (styles.xml) as scoped CSS so imported
   *  headings/body text render with their real font/size/color instead of the
   *  browser default. Called after every content load (create / import / setJSON). */
  #applyDocStyles(): void {
    const editor = this.#editor;
    if (!editor) return;
    const styles = editor.state.doc.attrs?.styles;
    // Re-stamp the ribbon so the Styles gallery reflects the loaded document's
    // style library (named + custom paragraph styles) — see styleItems().
    this.#renderChrome();
    const css = stylesToCss(styles, ".docen-page");
    const root = this.shadowRoot!;
    const existing = root.querySelector("#docen-doc-styles");
    if (css) {
      const styleEl = (existing ?? document.createElement("style")) as HTMLStyleElement;
      styleEl.id = "docen-doc-styles";
      // Wrap in @layer docxStyles so these named styles beat the reset layer
      // (layer order, not specificity) yet stay below unlayered inline styles.
      styleEl.textContent = "@layer docxStyles {\n" + css + "\n}";
      if (!existing) root.append(styleEl);
    } else {
      existing?.remove();
    }
    // Font metric + section geometry apply regardless of named styles — both
    // read attrs (default font / sectionProperties) that a styles-less document
    // still carries. The previous early `return` skipped them, leaving
    // --docen-font-metric and the section's page-size/pitch unset on freshly
    // loaded (styles-less) documents.
    this.#applyDefaultFontMetric();
    this.#applySectionGeometry();
  }

  /** Set a document-wide --docen-font-metric fallback on .docen-pages — the
   *  default font's `normal` ratio — so the page container's line-height
   *  (inherited by paragraphs without their own spacing) resolves to a real
   *  metric instead of the 1.2 fallback. Per-paragraph decorations override
   *  this for paragraphs that carry their own line-spacing. */
  #applyDefaultFontMetric(): void {
    const editor = this.#editor;
    const pages = this.shadowRoot?.querySelector<HTMLElement>(".docen-pages");
    if (!editor || !pages) return;
    const styles = (editor.state.doc.attrs?.styles ?? null) as StylesOptions | null;
    const { font } = effectiveRunProps(styles, null, {});
    const family = resolveFontName(font) ?? "serif";
    const ratio = fontNormalRatio({ family, bold: false, italic: false }).toFixed(4);
    pages.style.setProperty("--docen-font-metric", ratio);
  }

  /** Apply the document's section geometry — page size, orientation, and
   *  margins from `sectionProperties` (twips) — to the canvas, then
   *  re-paginate. Without this an imported document renders on the default
   *  A4 portrait + Normal margins instead of its real page setup, so the page
   *  count and breaks drift far from Word. twips → mm (1in = 1440tw = 25.4mm);
   *  canvas flips width/height for landscape via :host([orientation]). */
  #applySectionGeometry(): void {
    const editor = this.#editor;
    const canvas = this.shadowRoot?.querySelector("docen-canvas");
    if (!editor || !canvas) return;
    const sp = (editor.state.doc.attrs as Record<string, unknown> | undefined)
      ?.sectionProperties as
      | {
          page?: {
            size?: { width?: number; height?: number; orientation?: string };
            margin?: { top?: number; right?: number; bottom?: number; left?: number };
          };
          grid?: { linePitch?: number; type?: string } | null;
        }
      | undefined;
    const page = sp?.page;
    if (!page) return;
    // Page geometry in millimeters — the paper's natural unit (docx stores page
    // size/margins in twips; 1in = 1440tw = 25.4mm maps cleanly to mm). Font
    // sizes stay in pt (OOXML's unit); pt and mm are BOTH absolute CSS units
    // anchored to the same 96px/in reference pixel, so they render on one
    // consistent pixel grid — mm and pt do NOT need to be unified.
    if (page.size) {
      if (page.size.width) canvas.setAttribute("page-width", twipsToMm(page.size.width));
      if (page.size.height) canvas.setAttribute("page-height", twipsToMm(page.size.height));
    }
    // Render page-width × page-height directly. office-open's `orientation`
    // flag is unreliable (it can read "landscape" on portrait dimensions — this
    // very file is A4 portrait), so clear any prior landscape swap and let the
    // physical width/height decide the orientation.
    canvas.removeAttribute("orientation");
    if (page.margin) {
      const m = page.margin;
      const sides = [m.top, m.right, m.bottom, m.left];
      if (sides.every((s) => s != null))
        canvas.setAttribute("margin", sides.map(twipsToMm).join(" "));
    }
    // Document grid (w:docGrid) is applied PER PAGE: each page node renders its
    // section's linePitch as an inline line-height (page-node renderHTML), so
    // multi-section docs with different grids each render correctly. Nothing to
    // inject globally here — the per-page inline style cascades to the page's
    // paragraphs (line-height is inherited).
    this.#refreshGeometry();
  }

  /** The underlying Tiptap editor (for advanced, direct control). */
  getEditor(): Editor | undefined {
    return this.#editor;
  }

  /** Force a pagination re-measure now (bypasses the debounce). */
  repaginate(): void {
    if (this.#editor) pageStorageOf(this.#editor).repaginate();
  }
}

customElements.define("docen-document", DocenDocument);

export default DocenDocument;
