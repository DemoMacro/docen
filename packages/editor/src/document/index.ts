import {
  createDocxEditor,
  effectiveRunProps,
  generateDOCX,
  parseDOCX,
  parseHTML,
  stylesToCss,
  type JSONContent,
  type StylesOptions,
} from "@docen/docx";
import { Extension, type Editor } from "@docen/docx/core";
import { TableOfContents } from "@tiptap/extension-table-of-contents";
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
import { dispatchRibbonCommand } from "./commands";
// Side-effect import: registers the ribbon/header translation tables.
import "./i18n";
import { PageBreakView } from "./extensions/page-break";
import { Page, PageDocument } from "./extensions/page-node";
import { PagePlugin, pageStorageOf } from "./extensions/page-plugin";
import { SectionBreakMarks } from "./extensions/section-break";
import { SplitMarks } from "./extensions/split-paragraph";
import { SplitTable, SplitTableRow } from "./extensions/split-table";
import { buildRibbonInnerHTML } from "./ribbon-default";
import { clearMeasureCache } from "./utils/measure";
import { unwrapPages, wrapPages } from "./utils/wrap";

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
       page's content box — that is the origin the negative background-position
       offsets from, landing each leg in the margin gutter (not over text). The
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
        -24px 100%, -2px 100%,
        calc(100% + 24px) 100%, calc(100% + 2px) 100%;
      background-size: 23px 1px, 1px 23px, 23px 1px, 1px 23px, 23px 1px, 1px 23px, 23px 1px, 1px 23px;
      background-origin: content-box;
      background-repeat: no-repeat;
    }
    /* While focus is in a ribbon combobox dropdown, the editor is blurred and
       the browser stops painting its selection. Paint a CSS Custom Highlight
       over the selection so it stays visible — otherwise clicking the font/size
       box reads as "the selection was cancelled". Registered/cleared on editor
       blur/focus (see #setupBlurSelection); ignored where the Highlight API is
       unavailable. */
    ::highlight(docen-blur-selection) {
      background: var(--docen-color-selection-blur, rgba(0, 120, 215, 0.35));
    }
    /* Tables — Word inserts tables in the Table Grid style (a single black
       border on every cell). Without this a freshly inserted table is
       invisible: the docx table extension emits a border only when the node
       carries border attrs, and insertTable creates none. */
    .docen-pages table { border-collapse: collapse; }
    .docen-pages table td,
    .docen-pages table th {
      border: 1px solid #000;
    }
    /* Formatting marks (Show/Hide ¶) — Word shows these only while editing
       (non-printing). The show-marks command flips the host [show-marks]
       attribute; the marks themselves live entirely in CSS. */
    /* Pilcrow ¶ is painted by the FormattingMarks extension as a widget
       decoration — CSS ::after on a ProseMirror-managed <p> does not render. */
    .docen-pages .docen-para-mark {
      color: var(--docen-color-marks, #6e6e6e);
      user-select: none;
      pointer-events: none;
      margin-inline-start: 1px;
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
    <footer slot="status" part="status"></footer>
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
 *  one value = uniform; two = top/bottom then left/right). */
const MARGINS: Readonly<Record<string, string>> = {
  normal: "25.4mm",
  narrow: "12.7mm",
  moderate: "25.4mm 19.05mm",
  wide: "25.4mm 50.8mm",
};

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
/** CSS Custom Highlight painted over the editor selection while focus is
 *  elsewhere — the ribbon font-name / font-size comboboxes steal focus to open
 *  their dropdowns (fluent-dropdown calls `input.focus()` on click, which no
 *  `preventDefault` on pointer events can stop). Without this, the browser
 *  stops painting the selection the moment the box is clicked, so it looks
 *  "cancelled" even though `state.selection` survives and the chosen font/size
 *  is applied correctly. Browsers without the Highlight API no-op. */
const BLUR_SELECTION_HIGHLIGHT = "docen-blur-selection";

function paintBlurSelection(editor: Editor, name: string): void {
  const registry = (
    CSS as unknown as {
      highlights?: { set(name: string, highlight: unknown): void; delete(name: string): void };
    }
  ).highlights;
  const HighlightCtor = (
    globalThis as unknown as {
      Highlight?: new (...ranges: AbstractRange[]) => unknown;
    }
  ).Highlight;
  if (!registry || !HighlightCtor) return;
  const { from, to, empty } = editor.state.selection;
  if (empty) {
    registry.delete(name);
    return;
  }
  try {
    const start = editor.view.domAtPos(from);
    const end = editor.view.domAtPos(to);
    const range = document.createRange();
    range.setStart(start.node, start.offset);
    range.setEnd(end.node, end.offset);
    registry.set(name, new HighlightCtor(range));
  } catch {
    registry.delete(name);
  }
}

function clearBlurSelection(name: string): void {
  (CSS as unknown as { highlights?: { delete(name: string): void } }).highlights?.delete(name);
}

class DocenDocument extends HTMLElement {
  #editor?: Editor;
  #fileInput?: HTMLInputElement;
  #imageInput?: HTMLInputElement;
  /** Latest TOC anchors, refreshed by TableOfContents.onUpdate; used to resolve
   *  an outline click back to a document position (pos). */
  #anchors: readonly TocAnchor[] = [];
  /** Semantic fingerprint of the last outline tree — id/level/title only. `pos`
   *  shifts on every pagination re-flow but never changes what the pane shows,
   *  so it's excluded; the fingerprint is built from per-anchor arrays (not the
   *  serialized tree) so object key order can never cause a spurious mismatch. */
  #outlineSig = "";
  #unobserveLang?: () => void;
  /** Tears down the blur/focus listeners driving the blur-selection highlight. */
  #blurSelCleanup?: () => void;
  /** Tears down the transaction listener mirroring caret font/size → comboboxes. */
  #fontSyncCleanup?: () => void;
  // Format Painter captured marks + the pointerup listener that applies them.
  #painterMarks: readonly Mark[] | null = null;
  #painterOff?: () => void;
  /** Current zoom level (percent) applied to the canvas via CSS `zoom`. */
  #zoom = 100;
  /** Esc 兜底：浏览器全屏被 Esc 退出后，把 ribbon 还原为「始终显示」。 */
  readonly #onFullscreenChange = (): void => {
    if (document.fullscreenElement) return;
    const ribbon = this.shadowRoot?.querySelector("docen-ribbon");
    const workspace = this.shadowRoot?.querySelector("docen-workspace");
    if (ribbon?.getAttribute("data-ribbon-mode") === "auto-hide") {
      ribbon.removeAttribute("data-ribbon-mode");
      workspace?.removeAttribute("data-fullscreen");
    }
  };

  /** Ctrl+= / Ctrl+- / Ctrl+0 缩放、Ctrl+F 查找（Word 行为）。缩放在 ribbon
   *  combobox 等输入框中忽略（让按键到达输入框）；Ctrl+F 全局生效。preventDefault
   *  阻止浏览器原生的页面缩放/查找。 */
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
    // composedPath()[0] 是 shadow 内的真实 target（如 combobox 的 input）。
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

  /** Keep the editor selection visible while a ribbon combobox (font-name /
   *  font-size) holds focus. On editor blur, paint a CSS Custom Highlight over
   *  the current selection; on focus, clear it so the browser's native
   *  selection painting resumes. */
  #setupBlurSelection(): void {
    const editor = this.#editor;
    if (!editor) return;
    const paint = (): void => paintBlurSelection(editor, BLUR_SELECTION_HIGHLIGHT);
    const clear = (): void => clearBlurSelection(BLUR_SELECTION_HIGHLIGHT);
    editor.on("blur", paint);
    editor.on("focus", clear);
    this.#blurSelCleanup = (): void => {
      editor.off("blur", paint);
      editor.off("focus", clear);
      clear();
    };
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
        PagePlugin,
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
        // sectionBreak widget — a section boundary is paragraph attrs (not a
        // node), so it has no NodeView; a widget decoration paints the Fluent
        // divider marker after each section-carrying paragraph.
        SectionBreakMarks,
      ],
    });
    this.#applyDocStyles();

    // Keep the editor selection painted while a ribbon combobox holds focus.
    this.#setupBlurSelection();
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
    document.addEventListener("fullscreenchange", this.#onFullscreenChange);
    this.addEventListener("keydown", this.#onZoomKey);
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
    document.removeEventListener("fullscreenchange", this.#onFullscreenChange);
    this.removeEventListener("keydown", this.#onZoomKey);
    this.#blurSelCleanup?.();
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
            <fluent-switch data-event="autosave" checked aria-label="${autosave}"></fluent-switch>
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
            <docen-ribbon-button icon="close" label="${t("header.close")}" event="close" icon-only></docen-ribbon-button>
          </div>`;
  }

  /** Stamp the header + ribbon markup for the active locale (re-run on lang change). */
  #renderChrome(): void {
    const root = this.shadowRoot;
    if (!root) return;
    const styles = this.#editor?.state.doc.attrs?.styles ?? null;
    root.querySelector("docen-app-header")!.innerHTML = this.#renderHeader();
    root.querySelector("docen-ribbon")!.innerHTML = buildRibbonInnerHTML(styles);
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
  }

  /** Toggle a side task pane open/closed (ribbon View → toggle-navigation). */
  #togglePane(side: "start" | "end"): void {
    const pane = this.shadowRoot?.querySelector(`docen-task-pane[position="${side}"]`) as
      | (HTMLElement & { open: boolean })
      | null;
    if (pane) pane.open = !pane.open;
  }

  /** Apply a paper-size preset (a4/letter/…) to the canvas as raw mm, then
   *  re-paginate. The canvas takes only page-width/page-height; presets live
   *  here (PAPER_SIZES). */
  #setPageSize(value?: string): void {
    const canvas = this.shadowRoot?.querySelector("docen-canvas");
    const size = value ? PAPER_SIZES[value] : undefined;
    if (canvas && size) {
      canvas.setAttribute("page-width", String(size[0]));
      canvas.setAttribute("page-height", String(size[1]));
    }
    this.#refreshGeometry();
  }

  /** Apply orientation (portrait/landscape) to the canvas, then re-paginate.
   *  portrait clears the attribute (the canvas default); landscape sets it and
   *  the canvas swaps width/min-height via :host([orientation]). */
  #setOrientation(value?: string): void {
    const canvas = this.shadowRoot?.querySelector("docen-canvas");
    if (canvas && value) {
      if (value === "landscape") canvas.setAttribute("orientation", "landscape");
      else canvas.removeAttribute("orientation");
    }
    this.#refreshGeometry();
  }

  /** Apply a margin preset (normal/narrow/…) to the canvas as a CSS padding
   *  list, then re-paginate. Presets live here (MARGINS); the canvas takes the
   *  raw length list. */
  #setMargins(value?: string): void {
    const canvas = this.shadowRoot?.querySelector("docen-canvas");
    if (canvas && value && MARGINS[value]) canvas.setAttribute("margin", MARGINS[value]);
    this.#refreshGeometry();
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

  /** Refresh the status bar: "Page X of Y · Z%" — the caret's page over the
   *  page count, plus the zoom percent. Runs on every transaction (caret moves,
   *  a re-flow changes the page count) and on zoom / locale change. */
  #updateStatus(): void {
    const status = this.shadowRoot?.querySelector('footer[slot="status"]');
    if (!status) return;
    const editor = this.#editor;
    let page = 0;
    let total = 0;
    if (editor) {
      const from = editor.state.selection.from;
      editor.state.doc.forEach((node, offset) => {
        if (node.type.name !== "page") return;
        total++;
        if (page === 0 && from > offset && from <= offset + node.nodeSize) page = total;
      });
    }
    status.textContent = t("status.page-of")
      .replace("{page}", String(page || 1))
      .replace("{total}", String(total || 1))
      .replace("{zoom}", String(this.#zoom));
  }

  readonly #onCommand = (event: CustomEvent<{ event?: string; value?: string }>): void => {
    const { event: name, value } = event.detail ?? {};
    if (typeof name !== "string") return;
    // UI chrome actions are handled locally and need no Tiptap editor.
    if (name === "toggle-navigation") {
      this.#togglePane("start");
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
    // "save" is a document action, not a Tiptap command — handle locally.
    if (name === "save") {
      void this.#saveAs();
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
        this.#fileInput?.click();
        break;
      case "save-as":
        void this.#saveAs();
        break;
      case "print":
        this.#print();
        break;
      case "lang-zh":
        document.documentElement.lang = "zh-CN";
        break;
      case "lang-en":
        document.documentElement.lang = "en";
        break;
      // new / autosave: skeleton — wired when those features land.
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

  /** Load a .docx file into the editor and adopt its name as the filename. */
  async openDOCX(file: File): Promise<void> {
    this.openDOCXFromBuffer(await file.arrayBuffer());
    this.setAttribute("filename", file.name);
    this.#renderChrome();
  }

  /** Load a DOCX from an ArrayBuffer or Uint8Array (parseDOCX is synchronous). */
  openDOCXFromBuffer(buffer: ArrayBuffer | Uint8Array): void {
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
    if (!css) {
      existing?.remove();
      return;
    }
    const styleEl = (existing ?? document.createElement("style")) as HTMLStyleElement;
    styleEl.id = "docen-doc-styles";
    // Wrap in @layer docxStyles so these named styles beat the reset layer
    // (layer order, not specificity) yet stay below unlayered inline styles.
    styleEl.textContent = "@layer docxStyles {\n" + css + "\n}";
    if (!existing) root.append(styleEl);
    this.#applySectionGeometry();
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
    const mm = (twips: number): string => `${((twips / 1440) * 25.4).toFixed(2)}mm`;
    if (page.size) {
      if (page.size.width) canvas.setAttribute("page-width", mm(page.size.width));
      if (page.size.height) canvas.setAttribute("page-height", mm(page.size.height));
    }
    // Render page-width × page-height directly. office-open's `orientation`
    // flag is unreliable (it can read "landscape" on portrait dimensions — this
    // very file is A4 portrait), so clear any prior landscape swap and let the
    // physical width/height decide the orientation.
    canvas.removeAttribute("orientation");
    if (page.margin) {
      const m = page.margin;
      const sides = [m.top, m.right, m.bottom, m.left];
      if (sides.every((s) => s != null)) canvas.setAttribute("margin", sides.map(mm).join(" "));
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
