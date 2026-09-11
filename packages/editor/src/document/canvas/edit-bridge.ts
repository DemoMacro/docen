// The canvas route's editing base — a viewless Tiptap editor (element: null)
// driving the layout pipeline. The PM state is the single source of truth:
// every transaction re-flows the document (raf-merged) and the stage
// repaints; no DOM render of the doc exists at all. Text input arrives
// through an invisible textarea bridge — the canvas owns the visual surface,
// the textarea owns the browser's keyboard/IME machinery, and beforeinput
// events translate to state transactions. The view Proxy Tiptap installs for
// element:null supports command dispatch (insertContentAt/deleteRange), but
// anything touching real view internals (focus(), someProp via
// captureTransaction) crashes — so every translation here goes through pure
// PM state commands.

import type { ShapeTextStack } from "@docen/core";
import {
  docxExtensions,
  parseHTMLBody,
  DOCEN_CLIP_MIME,
  selectionSlicePayload,
  HEADING_COMPILE_MAP,
  nextOrderedReference,
  type JSONContent,
} from "@docen/docx";
import { Editor } from "@docen/docx/core";
import type { FlowPage } from "@docen/layout";
import paintBrushSvg from "@fluentui/svg-icons/icons/paint_brush_24_regular.svg?raw";
import { UndoRedo } from "@tiptap/extensions";
import {
  joinBackward,
  joinForward,
  selectAll,
  selectNodeBackward,
  selectNodeForward,
  splitBlock,
} from "@tiptap/pm/commands";
import { Fragment, type Node as PMNode, Slice } from "@tiptap/pm/model";
import { NodeSelection, TextSelection } from "@tiptap/pm/state";
import { getMatchHighlights } from "prosemirror-search";

import { DrawingGestures, type DrawingHit } from "../../drawing";
import { t } from "../../ui/i18n/localize";
import { collectListReferences, listLevelStepPatch } from "../extensions/commands";
import { KEYBOARD_SHORTCUTS } from "../extensions/keymap";
import { autocorrectOf } from "./autocorrect";
import { CaretMap, type TableZone } from "./caret-map";
import { CellSelection, cellAt, inSameTable } from "./cell-selection";
import { installChartHover, type ChartTip } from "./chart-hover";
import { followLink, installLinkHover, type LinkHit } from "./link-hover";
import { blockRuleOf, enterRuleOf, inlineRuleOf, isHyphenRun } from "./markdown-input";
import { sameChildPath } from "./stage";

/** Word's format-painter cursor: the text I-beam with the paint brush riding
 *  its lower right, as an SVG data URL (hot spot on the I-beam's insertion
 *  point). Fluent's own brush path scales into the corner. */
const FORMAT_PAINTER_CURSOR = (() => {
  const path = /d="([^"]+)"/.exec(paintBrushSvg)?.[1] ?? "";
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24">` +
    `<path d="M5 1.5v9M3.2 1.5h3.6M3.2 10.5h3.6" stroke="#000" stroke-width="1.5"` +
    ` fill="none" stroke-linecap="round"/>` +
    `<g transform="translate(9 9) scale(.58)"><path d="${path}" fill="#000"/></g>` +
    `</svg>`;
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}") 5 2, text`;
})();

/** A grapheme-boundary segmenter shared by the delete translations — surrogate
 *  pairs, combining marks, and emoji must delete as one user-perceived
 *  character. */
const segmenter = new Intl.Segmenter(undefined, { granularity: "grapheme" });

/** The code-unit length of `text`'s last grapheme (0 when empty) — the
 *  backspace cut. */
function lastGraphemeUnits(text: string): number {
  let last = 0;
  for (const { segment } of segmenter.segment(text)) last = segment.length;
  return last;
}

/** The code-unit length of `text`'s first grapheme (0 when empty) — the
 *  forward-delete cut. */
function firstGraphemeUnits(text: string): number {
  for (const { segment } of segmenter.segment(text)) return segment.length;
  return 0;
}

/** The code-unit cut of a word-delete backward from `offset` (Ctrl+Backspace):
 *  skip the caret's preceding whitespace, then take the run of non-whitespace
 *  up to the next boundary — Word's deleteWordLeft. */
function wordUnitsBackward(text: string, offset: number): number {
  let i = offset;
  while (i > 0 && /\s/.test(text[i - 1]!)) i--;
  const wordEnd = i;
  while (i > 0 && !/\s/.test(text[i - 1]!)) i--;
  return wordEnd - i;
}

/** The forward mirror of wordUnitsBackward (Ctrl+Delete). */
function wordUnitsForward(text: string, offset: number): number {
  const end = text.length;
  let i = offset;
  while (i < end && /\s/.test(text[i]!)) i++;
  const wordStart = i;
  while (i < end && !/\s/.test(text[i]!)) i++;
  return i - wordStart;
}

/** A furniture edit story — the header/footer editing mode. One story at a
 *  time: its editor, map, and render schedule live beside the main story's,
 *  and every input handler routes through the active one. */
export type StoryKind = "header" | "footer";
export type StorySlot = "default" | "first" | "even";

export interface EditBridgeStory {
  /** Live furniture geometry for a page — null when the doc has no
   *  furniture (no story is enterable). */
  geometry(
    kind: StoryKind,
    page: number,
  ): {
    stack: readonly import("@docen/layout").LaidOutStackItem[] | null;
    band: { top: number; bottom: number; paintY: number } | null;
    slot: StorySlot;
  } | null;
  /** The story's source JSON (empty when the slot has no content yet) —
   *  resolved for the section the anchor `page` belongs to (Word: double-
   *  clicking a band edits THAT page's section's furniture, whatever body
   *  position the caret happens to sit at). */
  read: (kind: StoryKind, slot: StorySlot, page: number) => JSONContent[];
  /** A story entered — drop in the stage chrome (grayed body, boundary).
   *  `page` is the anchor page the story edits in place on. */
  entered: (kind: StoryKind, slot: StorySlot, page: number) => void;
  /** A story transaction landed — re-project the furniture from this JSON
   *  (the body flow must not re-lay) and call back updateStoryMap. */
  onDoc: (kind: StoryKind, slot: StorySlot, json: JSONContent[]) => void;
  /** The story exited (a body click / Esc) — persist `json` when `dirty`,
   *  drop the stage chrome. */
  exit: (story: { kind: StoryKind; slot: StorySlot; json: JSONContent[]; dirty: boolean }) => void;
}

export interface EditBridgeOptions {
  /** A positioned host covering the canvas surface — the bridge's overlays
   *  mount here and it captures clicks to take focus. */
  host: HTMLElement;
  /** The input textarea's mount point — MUST sit outside any menu component:
   *  an ancestor fluent-menu treats Space/Enter as menu activation keys and
   *  preventDefaults them, which kills the textarea's default insertion (no
   *  beforeinput → spaces and Enter are silently dropped). Positioned, so the
   *  textarea's caret-anchored coordinates resolve against it. */
  inputHost: HTMLElement;
  /** Initial document (Tiptap JSON, e.g. parseDOCX's result). */
  content: JSONContent | Record<string, unknown>;
  /** raf-merged document callback — one per frame at most, with the fresh
   *  editor JSON for the full re-flow (compile → project → layout → paint). */
  onDoc: (json: JSONContent) => void;
  /** The positioned page frame for a page index (the caret overlay mounts
   *  inside it, page-local). Absent pages report null. */
  pageHost?: (page: number) => HTMLElement | null;
  /** Engine extensions for the viewless editor (defaults to the docx schema
   *  set). The host layers its own (commands, outline, search, …) here. */
  extensions?: Editor["extensionManager"]["extensions"];
  /** The stage's zoom factor (semantic page px → screen px). Overlays are
   *  written in screen px inside zoom-sized frames; hit-testing converts the
   *  other way. Defaults to 1 (unzoomed). */
  scale?: () => number;
  /** The page's in-front float boxes (page-local px) — squiggles clip against
   *  them: a front-of-text picture covers the text and its spelling wave
   *  (Word keeps only the selection and caret above front floats). */
  frontFloats?: (page: number) => Array<{ x: number; y: number; width: number; height: number }>;
  /** Header/footer edit stories — absent, the furniture bands are inert. */
  story?: EditBridgeStory;
  /** Drawing hit-test (page-local px) — the stage's painted-box table. A hit
   *  selects the drawing (Word: clicking a picture grabs it) instead of
   *  placing the caret behind it; absent, every click is text. */
  drawingAt?: (page: number, lx: number, ly: number) => DrawingHit | null;
  /** The PM node selection for a drawing hit — resolves the host paragraph
   *  position and the index-th drawing node inside it (null when the map
   *  cannot pair the host, e.g. a furniture story paragraph). A hit with a
   *  `childPath` targets a group member: the resolver decides between the
   *  member node (the group is entered) and the group itself (Word: a click
   *  selects the group until then). `enter` marks the entry double click —
   *  the member resolves even when the group was not entered yet. */
  drawingSelection?: (
    hit: {
      para: unknown;
      index: number;
      kind: "drawing" | "inline";
      childPath?: readonly number[];
    },
    enter?: boolean,
  ) => number | null;
  /** Re-resolves a selected drawing's painted box after a re-render (the
   *  selection drops when the drawing no longer paints). `childPath` matches
   *  a group member's box. */
  drawingBoxOf?: (
    para: unknown,
    index: number,
    kind: "drawing" | "inline",
    childPath?: readonly number[],
  ) => DrawingHit | null;
  /** A framed chart's sub-element boxes with fresh geometry — the
   *  sub-selection highlight re-reads them on every placement (a re-render
   *  re-objects the boxes, so a deleted series drops out and moved ones
   *  follow). */
  chartPartBoxes?: (para: unknown, index: number, kind: "drawing" | "inline") => DrawingHit[];
  /** The paint pass's editable text-box stacks — registered with each fresh
   *  caret map so a double click edits the shape's text in place. */
  shapeTextStacks?: () => readonly ShapeTextStack[];
  /** Re-finds the wpsShape node a text-box stack belongs to (the stack's
   *  host paragraph + drawing index — the same identity a hit box carries;
   *  `childPath` drills into a group's interior member). */
  shapeResolve?: (host: {
    para: unknown;
    index: number;
    childPath?: readonly number[];
  }) => { pos: number; node: PMNode } | null;
  /** Ctrl+Click on a `#name` link — the host resolves the bookmark anchor and
   *  scrolls it into view (followLink only opens external URLs). Absent,
   *  internal anchors are inert. */
  onInternalAnchor?: (name: string) => void;
  /** Every in-editor copy/cut (keyboard + ribbon) reports its content here —
   *  the Office Clipboard pane's collection feed. */
  onClipboardCollect?: (item: { text: string; payload: string | null }) => void;
  /** A keyboard paste landed rich content (the docen slice or styled HTML
   *  lane) — the host shows Word's paste-options bar over the pasted text. */
  onRichPaste?: (source: { kind: "slice" | "html"; raw: string; text: string }) => void;
  /** The border painter's armed state (Table Design → Draw Border). While
   *  active the canvas presses start edge sweeps instead of text selection. */
  borderPaint?: () => { active: boolean; eraser: boolean };
  /** The format painter's armed state (Home → Format Painter) — owns the
   *  cursor (Word's brush I-beam) until the paint lands or Esc disarms. */
  formatPaint?: () => boolean;
  /** The Markdown input mode (Options → Markdown input). While on, the typed
   *  leg applies the markdown block/inline conversions before autocorrect
   *  sees the character. Read per keystroke, so the toggle lands mid-session. */
  markdown?: () => boolean;
  /** A finished sweep — every crossed table edge (cell pos + side, both
   *  collapse halves of an interior line included) for the host to commit as
   *  one paint/erase command. */
  applyBorderPaint?: (sides: { pos: number; side: "top" | "bottom" | "left" | "right" }[]) => void;
}

export interface EditBridge {
  editor: Editor;
  /** Feed each render's flow result — rebuilds the pixel↔position map and
   *  re-places the caret against the fresh geometry. `pageOrigin` resolves a
   *  page to its content-box origin (multi-section documents: each page's
   *  own section's margins). */
  updatePages(
    pages: readonly FlowPage[],
    pageOrigin: (page: number) => { contentLeftPx: number; contentTopPx: number },
  ): void;
  /** Feed the active furniture story's fresh stack + band after the host
   *  re-projected it (each story keystroke rebuilds the story map). The
   *  map anchors at the band's `paintY` — the stack's own draw y. */
  updateStoryMap(
    stack: readonly import("@docen/layout").LaidOutStackItem[] | null,
    band: { top: number; bottom: number; paintY: number },
  ): void;
  /** Whether a furniture story is active (its kind, else null). */
  storyKind(): StoryKind | null;
  /** Open a furniture story programmatically (Insert → Header/Footer/Page
   *  Number) — the same lifecycle as the band double-click. `page` defaults
   *  to the caret's page; `seed` is inserted at the story's end after entry
   *  (the Page Number drop's PAGE field). False when blocked (no furniture,
   *  read-only, or a story already active). */
  enterStory(kind: StoryKind, page?: number, seed?: JSONContent): boolean;
  /** Leave the furniture story — tears down the story editor and restores
   *  the main story's overlays. Returns the story's final JSON; persistence
   *  stays with the host (it already ran `exit` for band/Esc routes). */
  exitStory(): { kind: StoryKind; slot: StorySlot; json: JSONContent[]; dirty: boolean } | null;
  /** Scroll the page holding a doc position into view (null when unmappable). */
  scrollIntoView(pos: number): void;
  /** The page index a doc position renders on (null when unmappable). */
  pageOf(pos: number): number | null;
  /** The first doc position rendered on a page (null when unmappable). */
  firstPosOfPage(page: number): number | null;
  /** The PM position just inside the laid paragraph (null when the map
   *  cannot pair it — render-only or unmapped). */
  posOfPara(para: unknown): number | null;
  /** A viewport point → the active story's doc position (null off-page or
   *  when the map is stale) — the context menu moves the caret to where it
   *  was right-clicked, like Word. */
  posAtClient(clientX: number, clientY: number): number | null;
  /** Right-click on a drawing: select the drawing under the viewport point
   *  (Word selects a picture before its context menu shows). True when a
   *  drawing was hit and selected. */
  selectDrawingAtClient(clientX: number, clientY: number): boolean;
  /** Enter crop mode on the selected image — the source shows in full with
   *  crop handles; Enter / a press outside commits, Esc cancels. False when
   *  the selection isn't a source-carrying image. */
  enterCropMode(): boolean;
  /** Arm Set Transparent Color: the next canvas press on a drawing samples
   *  the pixel under the pointer (display-normalized 0..1) and calls back
   *  instead of running the select chains; a press off any drawing disarms
   *  and clicks through. Pass null to disarm (Esc does too). */
  setTransparentPick(onPick: ((hit: DrawingHit, nx: number, ny: number) => void) | null): void;
  /** The multi-selection's members (primary + Shift+Click set) with their PM
   *  positions and page boxes — the host assembles the group/distribute
   *  payloads from it. Null when fewer than two resolve. */
  drawingMulti():
    | { pos: number; box: { x: number; y: number; width: number; height: number } }[]
    | null;
  /** The selection's last-line rect against its page frame — frame-relative
   * screen px (zoom applied), the anchor a floating comment compose positions
   * at (Word hangs the reply box in the margin beside the anchored text).
   * Main story only; null when unmappable or in a furniture story. */
  commentAnchorRect(
    from: number,
    to: number,
  ): { frame: HTMLElement; left: number; top: number; height: number } | null;
  /** The caret's rect against its page frame — frame-relative screen px
   *  (zoom applied). The paste-options bar hangs it beside the pasted
   *  content. Main story only; null when unmappable or in a story. */
  pasteAnchorRect(
    pos: number,
  ): { frame: HTMLElement; left: number; top: number; height: number } | null;
  /** Insert a docen slice payload (DOCEN_CLIP_MIME) into the ACTIVE story at
   *  the caret — the host's ribbon Paste routes here after reading the system
   *  clipboard. False when the payload did not parse. */
  insertSlicePayload(raw: string): boolean;
  /** The pinned payload of the most recent in-editor copy/cut (null after a
   *  copy that carried no slice). The async paste's fallback lane: Chromium
   *  never persists a copy event's custom types to the system clipboard, so
   *  the payload rides memory when the system text still matches it. */
  copiedSlice(): { payload: string; text: string } | null;
  /** Copy/cut the ACTIVE story's selection for the entry points that produce
   *  no copy event (ribbon / context-menu buttons). Pins the slice payload
   *  exactly like the keyboard path, so every paste entry recovers marks. */
  copySelection(cut: boolean): Promise<void>;
  /** The editor input currently routes into — the main editor, or the
   *  furniture story's when one is open. Ribbon commands must target the same
   *  editor the caret lives in, or they stamp the main document's stale
   *  selection. */
  activeEditor(): Editor;
  /** Move keyboard focus to the bridge's input surface (the editing focus —
   *  there is no DOM editor to focus). */
  focus(): void;
  /** Re-place the caret/selection/search overlays against the current
   *  geometry — needed when the zoom rescales the frames without a
   *  selection transaction. */
  replaceOverlays(): void;
  /** Hand the host's spell-check results to the squiggle overlay (the check
   *  itself runs in the host, debounced per transaction). */
  setSpellingIssues(issues: Array<{ from: number; to: number }>): void;
  destroy(): void;
}

/** One editing story — a viewless editor, its pixel↔position map, and its
 *  raf-merged render schedule. The main story lives for the bridge's
 *  lifetime; a furniture story (header/footer) joins it on entry and every
 *  input handler routes through the active one. */
interface Story {
  editor: Editor;
  map: CaretMap | null;
  /** The doc the map was zipped against — overlays refuse to draw when the
   *  editor has moved past it (a transaction applied, the raf-merged render
   *  not yet landed): a stale map's paragraphs no longer match the doc's
   *  positions, and painting through it puts the caret, selection and
   *  squiggles on the wrong text until the fresh map arrives. */
  mapDoc: PMNode | null;
  pageCount: number;
  lastCaretPos: number;
  /** The story's content callback — main: the full re-flow; furniture: the
   *  host's furniture-only re-projection. */
  onDoc: (json: JSONContent) => void;
  /** The single page frame every overlay mounts in (furniture stories never
   *  span pages); −1 = the main story follows each rect's own page. */
  anchorPage: number;
  /** The content origin per page (updatePages records it) — the story's
   *  caret map anchors at its anchor page's section origin. */
  pageOrigin: ((page: number) => { contentLeftPx: number; contentTopPx: number }) | null;
  raf: number;
  schedule(): void;
  /** Furniture-story only: what is being edited and its initial content. */
  kind?: StoryKind;
  slot?: StorySlot;
  initialJson?: string;
}

export function mountEditBridge(opts: EditBridgeOptions): EditBridge {
  const makeEditor = (content: unknown): Editor => {
    const editor = new Editor({
      element: null,
      extensions: [...(opts.extensions ?? docxExtensions), UndoRedo],
      content: content as never,
    });
    // element:null skips mount(), and with it plugin installation — the state
    // comes up schema-only (EditorState.create with no plugins). Register the
    // extension manager's full, priority-sorted plugin list by hand: undo
    // history, keymaps, outline, search … all read their plugin's state.
    for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
    return editor;
  };

  const makeStory = (
    content: unknown,
    onDoc: (json: JSONContent) => void,
    anchorPage: number,
  ): Story => {
    const s: Story = {
      editor: makeEditor(content),
      map: null,
      mapDoc: null,
      pageCount: 0,
      lastCaretPos: -1,
      onDoc,
      anchorPage,
      pageOrigin: null,
      raf: 0,
      // One relayout per frame regardless of keystroke bursts — content
      // changes only; selection-only transactions (drag-select, caret moves)
      // skip the render, their placement rides the synchronous selectionUpdate.
      schedule(): void {
        if (s.raf) return;
        s.raf = requestAnimationFrame(() => {
          s.raf = 0;
          s.onDoc(s.editor.getJSON());
        });
      },
    };
    return s;
  };

  const main = makeStory(opts.content, opts.onDoc, -1);
  let story: Story | null = null;
  const active = (): Story => story ?? main;
  // Word's Repeat (F4): the last plain-text typing session, kept on the
  // editor's storage so the host's repeat command shares this single source.
  // Consecutive insertText calls merge into one session ("typing a word");
  // any other transaction (caret move, delete, format, undo) ends it — the
  // recorded text stays repeatable, the next typing starts fresh. Multi-line
  // insertions (an insertContent paste shape) clear the record outright.
  // The marker rides the attachTransactions listener below, so it must exist
  // before any editor transaction can fire.
  let repeatInsert: string | null = null;
  // The main editor stays the public `editor` surface (document commands,
  // schema reads) — furniture stories are reachable only through the story
  // lifecycle below.
  const editor = main.editor;

  /** Content changes ride the story's own render (the map feed re-places).
   *  Everything else that still moves pixels — selection metas — re-places
   *  synchronously; idempotent with the selectionUpdate path. Also owns the
   *  Repeat typing-session bookkeeping (see repeatInsert above). */
  const attachTransactions = (s: Story): void => {
    s.editor.on("transaction", ({ transaction }) => {
      const storage = s.editor.storage as { repeat?: string; repeatLive?: boolean };
      if (repeatInsert !== null) {
        storage.repeat = storage.repeatLive ? (storage.repeat ?? "") + repeatInsert : repeatInsert;
        storage.repeatLive = true;
        repeatInsert = null;
      } else if (storage.repeatLive) {
        storage.repeatLive = false;
      }
      if (transaction.docChanged) s.schedule();
      else if (active() === s) placeCaret();
    });
  };
  attachTransactions(main);

  // The invisible input surface. Kept at 1px and positioned on click so the
  // IME candidate window anchors near the interaction point (true caret-point
  // anchoring lands with the caret milestone).
  const ta = document.createElement("textarea");
  // Marks the bridge input for the host's chrome-key gate: Ctrl+= / Ctrl+0
  // must still zoom with the caret in the document (this textarea IS the
  // document input, not a form field).
  ta.dataset.docenBridgeInput = "true";
  Object.assign(ta.style, {
    position: "absolute",
    width: "1px",
    height: "1px",
    opacity: "0",
    border: "none",
    padding: "0",
    margin: "0",
    resize: "none",
    outline: "none",
    overflow: "hidden",
    background: "transparent",
    color: "transparent",
    caretColor: "transparent",
    zIndex: "10",
  } satisfies Partial<CSSStyleDeclaration>);
  ta.setAttribute("aria-label", "Document text input");
  ta.setAttribute("autocapitalize", "off");
  ta.setAttribute("autocorrect", "off");
  ta.spellcheck = false;

  // The caret overlay — a thin div mounted inside the current page's frame
  // (page-relative positioning for free), blinked via the Web Animations API.
  const caret = document.createElement("div");
  Object.assign(caret.style, {
    position: "absolute",
    width: "1px",
    background: "#000",
    pointerEvents: "none",
    zIndex: "5",
    display: "none",
  } satisfies Partial<CSSStyleDeclaration>);
  opts.host.append(caret);
  let blink: Animation | null = null;

  // The pixel↔position geometry lives on each Story, rebuilt from every
  // render's feed. Between a transaction and its re-render it is one frame
  // stale — caretRect tolerates out-of-range positions by hiding until the
  // fresh map lands.

  /** The DOM page frame a rect's page maps to — furniture stories pin every
   *  overlay to their anchor page (their map has one pseudo page). */
  const framePage = (s: Story, page: number): number =>
    s.anchorPage >= 0 ? s.anchorPage + page : page;

  const restartBlink = (): void => {
    blink?.cancel();
    blink = caret.animate([{ opacity: 1 }, { opacity: 0 }], {
      duration: 1000,
      iterations: Infinity,
      direction: "alternate",
      easing: "steps(1,end)",
    });
  };

  /** The map is drawable only against the doc it was zipped from — a
   *  transaction the raf-merged render has not answered yet leaves it one
   *  generation behind, and painting through it puts the caret, selection
   *  and squiggles on the wrong text for a frame. Overlays hide until the
   *  fresh map lands (a selection-only change never bumps the doc, so
   *  drag-select and caret moves keep drawing synchronously). */
  const mapFresh = (s: Story): s is Story & { map: CaretMap } =>
    s.map?.valid === true && s.mapDoc === s.editor.state.doc;

  /** One pooled overlay strip: final CSS values, compared against the pool
   *  entry so an unchanged placement writes nothing (the per-keystroke
   *  placeCaret cascade must cost zero DOM mutations). */
  interface OverlayRect {
    page: number;
    x: number;
    y: number;
    width: number;
    height: number;
    background: string;
  }

  interface PoolEntry {
    el: HTMLDivElement;
    frame: HTMLElement;
    left: number;
    top: number;
    width: number;
    height: number;
    background: string;
  }

  /** Place rects into a layer's element pool, reusing divs in order: entries
   *  past the new count are removed, geometry/background changes rewrite the
   *  style in place, everything else is left untouched. */
  const pooledPlace = (
    pool: PoolEntry[],
    rects: OverlayRect[],
    base: Partial<CSSStyleDeclaration>,
  ): void => {
    const scale = opts.scale?.() ?? 1;
    const placed = rects
      .map((r) => ({ r, frame: opts.pageHost?.(r.page) ?? null }))
      .filter((p): p is { r: OverlayRect; frame: HTMLElement } => p.frame !== null);
    for (let i = placed.length; i < pool.length; i++) pool[i]!.el.remove();
    pool.length = placed.length;
    for (const [i, p] of placed.entries()) {
      const left = p.r.x * scale;
      const top = p.r.y * scale;
      const width = p.r.width * scale;
      const height = p.r.height * scale;
      let entry = pool[i];
      if (!entry) {
        // Sentinel-recorded values force the new div through the write path
        // below — the pool tracks what was last written, not what is wanted.
        entry = {
          el: document.createElement("div"),
          frame: p.frame,
          left: NaN,
          top: NaN,
          width: NaN,
          height: NaN,
          background: "",
        };
        pool[i] = entry;
        p.frame.append(entry.el);
      } else if (entry.frame !== p.frame) {
        p.frame.append(entry.el);
        entry.frame = p.frame;
      }
      if (
        entry.left !== left ||
        entry.top !== top ||
        entry.width !== width ||
        entry.height !== height ||
        entry.background !== p.r.background
      ) {
        Object.assign(entry.el.style, base, {
          position: "absolute",
          pointerEvents: "none",
          background: p.r.background,
          left: `${left}px`,
          top: `${top}px`,
          width: `${width}px`,
          height: `${height}px`,
        });
        Object.assign(entry, { left, top, width, height, background: p.r.background });
      }
    }
  };

  /** The selection highlight — one translucent div per crossed line, kept in
   *  a pool so an unchanged selection costs nothing (the placement runs on
   *  every caret move, not just selection changes).
   *  A cell selection highlights each cell's whole grid box instead (Word's
   *  cell highlight covers the slot, insets included). A node selection (a
   *  picked drawing) highlights nothing: the drawing's selection frame owns
   *  the look, and a range sweep across the atom paints a line box wider
   *  than the drawing — Word shows no text highlight under it. */
  const selectionPool: PoolEntry[] = [];
  const placeSelection = (): void => {
    const s = active();
    const rects: OverlayRect[] = [];
    if (mapFresh(s)) {
      const sel = s.editor.state.selection;
      const spans =
        sel instanceof CellSelection
          ? s.map.cellSelectionRects(sel)
          : sel instanceof NodeSelection || sel.from === sel.to
            ? []
            : s.map.selectionRects(sel.from, sel.to);
      for (const r of spans)
        rects.push({
          page: framePage(s, r.page),
          x: r.xPx,
          y: r.yPx,
          width: r.widthPx,
          height: r.heightPx,
          background: "rgba(0,120,215,.25)",
        });
    }
    pooledPlace(selectionPool, rects, { zIndex: "4" });
  };

  /** Search-match highlights — the selection layer's pooled pattern:
   *  prosemirror-search owns the matches (PM decorations), and each match
   *  range becomes one translucent div per crossed line. The active match —
   *  findNext/replaceNext select it, so it is the match overlapping the
   *  selection — gets the deeper tint. zIndex 3 keeps every match under the
   *  selection (4) and caret (5); an empty query matches nothing, so the
   *  layer is simply empty. */
  const searchPool: PoolEntry[] = [];
  const placeSearch = (): void => {
    const s = active();
    const rects: OverlayRect[] = [];
    if (mapFresh(s)) {
      const sel = s.editor.state.selection;
      for (const deco of getMatchHighlights(s.editor.state).find()) {
        const { from, to } = deco as { from: number; to: number };
        const activeMatch = from <= sel.to && sel.from <= to;
        for (const r of s.map.selectionRects(from, to))
          rects.push({
            page: framePage(s, r.page),
            x: r.xPx,
            y: r.yPx,
            width: r.widthPx,
            height: r.heightPx,
            background: activeMatch ? "rgba(255,141,35,.7)" : "rgba(255,213,79,.45)",
          });
      }
    }
    pooledPlace(searchPool, rects, { zIndex: "3" });
  };

  /** Spelling squiggles — the search layer's pooled pattern again: the host
   *  runs the dictionary check (debounced per transaction) and hands the
   *  issue ranges here; each becomes one thin div hugging the line's
   *  baseline with a red wave drawn by a repeating SVG. zIndex 2 keeps
   *  squiggles under search matches, selection, and caret. A squiggle an
   *  in-front float covers keeps only its visible strips — Word hides the
   *  wave under a front-of-text picture, while the selection and caret stay
   *  whole. */
  const spellingPool: PoolEntry[] = [];
  const SQUIGGLE =
    "url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='6' height='3'%3E" +
    "%3Cpath d='M0 2.5 L1.5 0.5 L3 2.5 L4.5 0.5 L6 2.5' fill='none' stroke='%23e81123'/%3E%3C/svg%3E\")";
  let spellingIssues: Array<{ from: number; to: number }> = [];

  /** `r` minus every hole it meets — axis-aligned leftovers only (a 3px-tall
   *  squiggle cut by a float box keeps its left/right strips; a fully
   *  covered one vanishes). */
  const rectMinus = (
    r: { x: number; y: number; width: number; height: number },
    holes: readonly { x: number; y: number; width: number; height: number }[],
  ): Array<{ x: number; y: number; width: number; height: number }> => {
    let parts = [r];
    for (const h of holes) {
      const next: typeof parts = [];
      for (const p of parts) {
        const hx0 = Math.max(p.x, h.x);
        const hx1 = Math.min(p.x + p.width, h.x + h.width);
        const hy0 = Math.max(p.y, h.y);
        const hy1 = Math.min(p.y + p.height, h.y + h.height);
        if (hx0 >= hx1 || hy0 >= hy1) {
          next.push(p);
          continue;
        }
        if (hy0 > p.y) next.push({ x: p.x, y: p.y, width: p.width, height: hy0 - p.y });
        if (hy1 < p.y + p.height) {
          next.push({ x: p.x, y: hy1, width: p.width, height: p.y + p.height - hy1 });
        }
        if (hx0 > p.x) next.push({ x: p.x, y: hy0, width: hx0 - p.x, height: hy1 - hy0 });
        if (hx1 < p.x + p.width) {
          next.push({ x: hx1, y: hy0, width: p.x + p.width - hx1, height: hy1 - hy0 });
        }
      }
      parts = next;
    }
    return parts;
  };

  const placeSpelling = (): void => {
    const s = active();
    const rects: OverlayRect[] = [];
    if (mapFresh(s)) {
      for (const issue of spellingIssues) {
        for (const r of s.map.selectionRects(issue.from, issue.to)) {
          const wave = { x: r.xPx, y: r.yPx + r.heightPx - 3, width: r.widthPx, height: 3 };
          const holes = opts.frontFloats?.(r.page) ?? [];
          for (const g of holes.length ? rectMinus(wave, holes) : [wave])
            rects.push({
              page: framePage(s, r.page),
              x: g.x,
              y: g.y,
              width: g.width,
              height: g.height,
              background: `${SQUIGGLE} repeat-x`,
            });
        }
      }
    }
    pooledPlace(spellingPool, rects, { zIndex: "2" });
  };

  const placeCaret = (): void => {
    placeSelection();
    placeSearch();
    placeSpelling();
    // A selection that stopped being the drawing's NodeSelection (arrow keys,
    // a command, undo) drops the selection box — the box mirrors the PM state.
    if (draw.selected && !(main.editor.state.selection instanceof NodeSelection)) {
      draw.clear();
    }
    const s = active();
    if (!mapFresh(s)) {
      caret.style.display = "none";
      return;
    }
    const { from, to } = s.editor.state.selection;
    if (from !== to) {
      // A selection replaces the caret (Word hides it too).
      caret.style.display = "none";
      return;
    }
    const rect = s.map.caretRect(from);
    const frame = rect ? (opts.pageHost?.(framePage(s, rect.page)) ?? null) : null;
    if (!rect || !frame) {
      caret.style.display = "none";
      return;
    }
    if (frame !== caret.parentElement) frame.append(caret);
    caret.style.display = "block";
    const scale = opts.scale?.() ?? 1;
    caret.style.left = `${rect.xPx * scale}px`;
    caret.style.top = `${rect.yPx * scale}px`;
    caret.style.height = `${rect.heightPx * scale}px`;
    // Keep the textarea anchored at the caret so the IME candidate window
    // opens at the typing point.
    const hostRect = opts.inputHost.getBoundingClientRect();
    const frameRect = frame.getBoundingClientRect();
    ta.style.left = `${frameRect.left - hostRect.left + rect.xPx * scale}px`;
    ta.style.top = `${frameRect.top - hostRect.top + rect.yPx * scale}px`;
    if (from !== s.lastCaretPos) {
      s.lastCaretPos = from;
      restartBlink();
    }
  };
  main.editor.on("selectionUpdate", placeCaret);

  /** A viewport point → the hit page, its page-local coordinates, and its
   *  rect (in semantic page px — the caret map knows nothing of the zoom).
   *  Pure frame geometry: story routing decides what a hit means. */
  const hitPage = (
    clientX: number,
    clientY: number,
  ): { page: number; lx: number; ly: number; w: number; h: number } | null => {
    const hostRect = opts.host.getBoundingClientRect();
    const x = clientX - hostRect.left;
    const y = clientY - hostRect.top;
    const scale = opts.scale?.() ?? 1;
    for (let p = 0; p < main.pageCount; p++) {
      const frame = opts.pageHost?.(p);
      if (!frame) continue;
      const r = frame.getBoundingClientRect();
      const lx = x - (r.left - hostRect.left);
      const ly = y - (r.top - hostRect.top);
      if (lx >= 0 && ly >= 0 && lx < r.width && ly < r.height) {
        return { page: p, lx: lx / scale, ly: ly / scale, w: r.width / scale, h: r.height / scale };
      }
    }
    return null;
  };

  const setSel = (pos: number, anchor?: number): void => {
    active().editor.commands.command(({ state, dispatch }) => {
      // The cast bridges the dual PM d.ts identity (same runtime
      // instance — see the module's command casts). create takes
      // (anchor, head) — the anchor leads.
      dispatch?.(
        state.tr.setSelection(TextSelection.create(state.doc, anchor ?? pos, pos) as never),
      );
      return true;
    });
  };

  // Word's multi-click selection: the second click takes the word under the
  // caret, the third the whole paragraph. Word boundaries approximate as
  // same-class runs — Latin/digit words, CJK ideograph runs, whitespace runs;
  // punctuation stands alone (Word picks the single mark).
  const charClass = (ch: string): number =>
    /\s/.test(ch) ? 0 : /[㐀-鿿豈-﫿]/.test(ch) ? 1 : /[0-9A-Za-z]/.test(ch) ? 2 : 3;

  const setSelClick = (pos: number, clicks: number, extendFrom?: number): void => {
    const { doc } = active().editor.state;
    if (clicks >= 2) {
      const $pos = doc.resolve(pos);
      const po = $pos.parentOffset;
      const base = pos - po;
      // Word's Shift+multi-click: the word/paragraph pick extends the prior
      // anchor — the selection grows to cover whichever side the pick is on.
      const extend = (from: number, to: number): void => {
        if (extendFrom != null) {
          if (extendFrom <= from) setSel(to, extendFrom);
          else setSel(extendFrom, from);
        } else {
          setSel(to, from);
        }
      };
      if (clicks >= 3) {
        extend(base, base + $pos.parent.content.size);
        return;
      }
      // Flat text with one placeholder per inline leaf keeps the string index
      // aligned with the parent's content offsets (text nodes contribute
      // their length, atom nodes their nodeSize of 1).
      const flat = $pos.parent.textBetween(0, $pos.parent.content.size, undefined, "￼");
      if (flat.length === $pos.parent.content.size) {
        const anchor = flat[po] ?? flat[po - 1];
        if (anchor != null) {
          const cls = charClass(anchor);
          let from = po;
          let to = po;
          if (cls === 3) {
            to = Math.min(po + 1, flat.length);
            from = to - 1;
          } else {
            while (from > 0 && charClass(flat[from - 1]!) === cls) from--;
            while (to < flat.length && charClass(flat[to]!) === cls) to++;
          }
          extend(base + from, base + to);
          return;
        }
      }
    }
    if (extendFrom != null) {
      if (extendFrom <= pos) setSel(pos, extendFrom);
      else setSel(extendFrom, pos);
    } else {
      setSel(pos);
    }
  };

  /** A click's selection: Word's Shift+Click extends from the caret's anchor
   *  instead of dropping a fresh caret (the pick — word, paragraph, or bare
   *  point — grows whichever side it falls on), and a continuing drag grows
   *  that same selection from the anchor. A plain click drops a caret and
   *  arms the drag from it. */
  const clickSelection = (pos: number, event: MouseEvent, clicks: number): void => {
    const anchor = event.shiftKey ? active().editor.state.selection.anchor : undefined;
    if (clicks >= 2) {
      setSelClick(pos, clicks, anchor);
      // Word keeps drag-extending after a word/paragraph pick — the drag
      // anchor is the pick's far edge, not the click point.
      dragAnchor = active().editor.state.selection.anchor;
      dragStart = { x: event.clientX, y: event.clientY };
      dragMoved = false;
      return;
    }
    setSel(pos, anchor);
    dragAnchor = anchor ?? pos;
    dragStart = { x: event.clientX, y: event.clientY };
    dragMoved = false;
  };

  /** The drawing selection state machine — the selection frame and crop
   *  overlays plus their write-backs (resize/move/rotate/crop land through
   *  the engine's drawing commands). Injected with this bridge's host
   *  surface; the bridge keeps only the hit chains that route into it. */
  const draw = new DrawingGestures({
    editor: () => main.editor,
    drawingSelection: (hit, enter) => opts.drawingSelection?.(hit, enter) ?? null,
    drawingBoxOf: (para, index, kind, childPath) =>
      opts.drawingBoxOf?.(para, index, kind, childPath) ?? null,
    chartPartBoxes: (para, index, kind) => opts.chartPartBoxes?.(para, index, kind) ?? [],
    // The drop re-anchor's target: the paragraph under the drop point, as its
    // content-end insertion position plus laid box origin (the re-anchor's
    // offset seed). Clamped: a drop into the top/bottom margin or a band the
    // wrap zone just emptied has no line under it within the unclamped
    // distance cap — the nearest line's paragraph hosts the re-anchor.
    paragraphAt: (page, x, y) => {
      if (story || !main.map?.valid) return null;
      const pos = main.map.posAtPoint(page, x, y, true);
      if (pos == null) return null;
      const doc = main.editor.state.doc;
      const $pos = doc.resolve(pos);
      let depth = $pos.depth;
      while (depth > 0 && $pos.node(depth).type.name !== "paragraph") depth--;
      if (depth === 0) return null;
      const paraPos = $pos.before(depth);
      const para = doc.nodeAt(paraPos);
      if (!para) return null;
      const rect = main.map.caretRect(paraPos + 1);
      if (!rect) return null;
      return { contentPos: paraPos + para.nodeSize - 1, xPx: rect.xPx, yPx: rect.yPx };
    },
    // The drop page for a drag that may have crossed pages: only the host
    // knows the pages' screen geometry. Null off the pages — the drop
    // commits nothing rather than guessing a frame.
    pageAtPoint: (clientX, clientY) => {
      const hit = hitPage(clientX, clientY);
      return hit ? { page: hit.page, x: hit.lx, y: hit.ly, w: hit.w, h: hit.h } : null;
    },
    // A drop whose paragraph sits in another table cell than the anchor's
    // (out of the cell, across cells, or into one) must re-home: the cell
    // clamp pins the box, so offsets can't express the move. Both sides
    // resolve as the innermost cell around the position — a drawing anchored
    // in the body vs a drop inside a table crosses just the same (Word).
    crossesCell: (hit, page, x, y) => {
      if (story || !main.map?.valid) return false;
      const cellPos = (pos: number): number | null => {
        const $pos = main.editor.state.doc.resolve(pos);
        for (let d = $pos.depth; d > 0; d--) {
          const type = $pos.node(d).type.name;
          if (type === "tableCell" || type === "tableHeader") return $pos.before(d);
        }
        return null;
      };
      const anchor = opts.drawingSelection?.(hit) ?? null;
      if (anchor == null) return false;
      const drop = main.map.posAtPoint(page, x, y, true);
      if (drop == null) return false;
      return cellPos(anchor) !== cellPos(drop);
    },
    pageHost: (page) => opts.pageHost?.(page) ?? null,
    scale: () => opts.scale?.() ?? 1,
    // The value-drag commit: one data point on the chart the NodeSelection
    // holds (the gesture only arms while the chart is framed).
    applyChartValue: (series, point, value) => {
      const sel = main.editor.state.selection;
      if (!(sel instanceof NodeSelection) || sel.node.type.name !== "chart") return;
      main.editor.commands["chart-value-apply"](JSON.stringify({ series, point, value }));
    },
  });
  draw.mount(opts.host);

  /** A viewport point → the active story's doc position (furniture stories
   *  map through their single pseudo page). Clamping drags resolve the
   *  nearest line regardless of distance — a drag overshooting past the
   *  last line's band must keep extending to the line's end (Word), not
   *  stall the selection mid-line. */
  const posAtClient = (clientX: number, clientY: number, clamp = false): number | null => {
    const s = active();
    const hit = hitPage(clientX, clientY);
    if (!hit || !s.map?.valid) return null;
    return s.map.posAtPoint(story ? 0 : hit.page, hit.lx, hit.ly, clamp);
  };

  // Word's table grips: the black bar arrows hover just outside the table
  // (down arrows over each column, right arrows beside each row) plus the
  // select-all square at the table's top-left corner. One overlay div per
  // hover state; clicking resolves the geometry back to a PM cell through
  // the caret map's cell boxes and runs the select command.
  type GripKind = "col" | "row" | "table" | null;
  const grip: { kind: GripKind; zone: TableZone | null; index: number } = {
    kind: null,
    zone: null,
    index: -1,
  };
  const GRIP_WINDOW = 12;
  const gripEl = document.createElement("div");
  Object.assign(gripEl.style, {
    position: "absolute",
    display: "none",
    pointerEvents: "none",
    zIndex: "7",
    cursor: "pointer",
  } satisfies Partial<CSSStyleDeclaration>);
  // Two faces: the strip arrow (columns/rows) and Word's boxed-cross grid
  // square (the select-all grip) — placeGrip shows exactly one per mode.
  gripEl.innerHTML =
    '<svg width="12" height="12" viewBox="0 0 12 12"><path d="M1 6h7M5 2l4 4-4 4z" fill="#454545"/></svg>' +
    '<svg width="12" height="12" viewBox="0 0 12 12"><rect x="0.5" y="0.5" width="11" height="11" fill="#ffffff" stroke="#7f7f7f"/><path d="M6 1v10M1 6h10" stroke="#7f7f7f"/></svg>';
  opts.host.append(gripEl);

  const placeGrip = (): void => {
    const frame = grip.zone ? (opts.pageHost?.(framePage(active(), grip.zone.page)) ?? null) : null;
    if (!grip.zone || !frame) {
      gripEl.style.display = "none";
      return;
    }
    if (frame !== gripEl.parentElement) frame.append(gripEl);
    const scale = opts.scale?.() ?? 1;
    const z = grip.zone;
    let left = z.xPx;
    let top = z.yPx;
    let width = 12;
    let height = 12;
    let rotate = "";
    if (grip.kind === "col") {
      const colLeft = z.colEdges[grip.index]!;
      width = (z.colEdges[grip.index + 1]! - colLeft) * scale;
      left = (z.xPx + colLeft) * scale;
      top = (z.yPx - GRIP_WINDOW - 2) * scale;
      rotate = "rotate(90deg)";
    } else if (grip.kind === "row") {
      const rowTop = z.rowEdges[grip.index]!;
      height = (z.rowEdges[grip.index + 1]! - rowTop) * scale;
      left = (z.xPx - GRIP_WINDOW - 2) * scale;
      top = (z.yPx + rowTop) * scale;
    } else {
      left = (z.xPx - GRIP_WINDOW - 1) * scale;
      top = (z.yPx - GRIP_WINDOW - 1) * scale;
    }
    gripEl.style.left = `${left}px`;
    gripEl.style.top = `${top}px`;
    gripEl.style.width = `${width}px`;
    gripEl.style.height = `${height}px`;
    // The column arrow is the row arrow rotated to face down, centered in
    // its strip; the corner square (the select grip and its always-on hover
    // preview alike) swaps the arrow for the boxed-cross grid face — the
    // strip modes must stay square-less or the grid would smear into a bar
    // across the whole strip.
    const [arrow, grid] = Array.from(gripEl.children) as SVGElement[];
    arrow?.setAttribute(
      "style",
      rotate ? `display:block;margin:auto;transform:${rotate}` : "display:none",
    );
    grid?.setAttribute("style", rotate ? "display:none" : "display:block");
    gripEl.style.background = "transparent";
    gripEl.style.display = "block";
  };

  /** The hover pass — mousemoves outside a drag resolve against the active
   *  page's table zones (innermost zone wins on nested tables). */
  const hoverTableGrip = (event: MouseEvent): void => {
    grip.kind = null;
    grip.zone = null;
    grip.index = -1;
    const s = active();
    if (s.map?.valid && !story) {
      const hit = hitPage(event.clientX, event.clientY);
      // The bars live just outside the zone — widen the zone test to the
      // grip window (the strips reach ~14px past the table's edges).
      const zone = hit
        ? s.map.tableZoneAt(story ? 0 : hit.page, hit.lx, hit.ly, GRIP_WINDOW + 2)
        : null;
      if (hit && zone) {
        const lx = hit.lx - zone.xPx;
        const ly = hit.ly - zone.yPx;
        // The corner square wins the top-left overlap with the row strip.
        if (lx >= -GRIP_WINDOW - 1 && lx <= 1 && ly >= -GRIP_WINDOW - 1 && ly <= 1) {
          grip.kind = "table";
          grip.zone = zone;
        } else if (ly >= -GRIP_WINDOW - 2 && ly <= 4 && lx > 0 && lx < zone.widthPx) {
          const idx = zone.colEdges.findIndex(
            (e, i) => i < zone.colEdges.length - 1 && lx >= e && lx < zone.colEdges[i + 1]!,
          );
          if (idx >= 0) {
            grip.kind = "col";
            grip.zone = zone;
            grip.index = idx;
          }
        } else if (lx >= -GRIP_WINDOW - 2 && lx <= 4 && ly > 0 && ly < zone.heightPx) {
          const idx = zone.rowEdges.findIndex(
            (e, i) => i < zone.rowEdges.length - 1 && ly >= e && ly < zone.rowEdges[i + 1]!,
          );
          if (idx >= 0) {
            grip.kind = "row";
            grip.zone = zone;
            grip.index = idx;
          }
        } else if (lx > 0 && lx < zone.widthPx && ly > 0 && ly < zone.heightPx) {
          // Word: hovering anywhere in the table shows the corner square —
          // but only the corner window itself clicks it; here a click edits.
          grip.zone = zone;
        }
      }
    }
    placeGrip();
  };

  /** A grip click: park the caret in the strip's first cell (the commands
   *  resolve the target from the caret's ancestry) and run Word's select. */
  const applyGrip = (): void => {
    const s = active();
    const zone = grip.zone;
    const kind = grip.kind;
    grip.kind = null;
    grip.zone = null;
    placeGrip();
    if (!kind || !zone || !s.map?.valid) return;
    // The first cell box inside the strip — its PM cell pos (+2) is the
    // inner position the select commands anchor from. Col/row boxes OVERLAP
    // the strip rather than filling it: a strip swallowed by a merged cell
    // must still grip (the select commands widen to the covering span).
    let anchorPos = -1;
    for (const [pos, boxes] of s.map.cellBoxes) {
      // A cell repeated across pieces keeps one box per placement — the
      // first (the original placement) is the geometry the strips test.
      const box = boxes[0]!;
      if (box.page !== zone.page) continue;
      const inner =
        kind === "col"
          ? box.xPx + box.widthPx > zone.xPx + zone.colEdges[grip.index]! + 0.5 &&
            box.xPx < zone.xPx + zone.colEdges[grip.index + 1]! &&
            box.yPx >= zone.yPx &&
            box.yPx + box.heightPx <= zone.yPx + zone.heightPx + 0.5
          : kind === "row"
            ? box.yPx + box.heightPx > zone.yPx + zone.rowEdges[grip.index]! + 0.5 &&
              box.yPx < zone.yPx + zone.rowEdges[grip.index + 1]! &&
              box.xPx >= zone.xPx &&
              box.xPx + box.widthPx <= zone.xPx + zone.widthPx + 0.5
            : box.xPx >= zone.xPx && box.yPx >= zone.yPx;
      if (inner && (anchorPos < 0 || pos < anchorPos)) anchorPos = pos + 2;
    }
    if (anchorPos < 0) return;
    setSel(anchorPos);
    if (kind === "col") s.editor.commands["select-table-column"]();
    else if (kind === "row") s.editor.commands["select-table-row"]();
    else s.editor.commands["select-table"]();
  };

  // Mouse selection: mousedown anchors, moves extend, mouseup settles. The
  // 3px threshold keeps a plain click from flashing a degenerate selection.
  let dragAnchor: number | null = null;
  let dragMoved = false;
  let dragStart: { x: number; y: number } | null = null;

  // Border-painter sweep (Table Design → Draw Border): a press on a table
  // edge collects every edge crossed before release, deduped by cell+side —
  // re-crossing the same line in one drag paints it once.
  let borderSweep: Map<string, { pos: number; side: "top" | "bottom" | "left" | "right" }> | null =
    null;

  // Drag auto-scroll (Word): a drag resting in the scroll container's
  // top/bottom edge hot zone keeps the document scrolling so the selection
  // reaches content outside the viewport — without it the last line below
  // the fold can never be dragged into the selection. Each scrolled frame
  // re-resolves the head at the resting pointer (clamped), so the selection
  // grows while the document moves under it.
  const SCROLL_EDGE_PX = 28;
  const SCROLL_MAX_SPEED_PX = 16;
  const dragPoint = { x: 0, y: 0 };
  let scrollRaf = 0;
  const scrollContainer = (): HTMLElement | null => {
    let el: HTMLElement | null = opts.host.parentElement;
    while (el) {
      if (el.scrollHeight > el.clientHeight + 1) return el;
      el = el.parentElement;
    }
    return null;
  };
  const dragAutoScroll = (): void => {
    scrollRaf = 0;
    if (dragAnchor == null) return;
    const container = scrollContainer();
    if (!container) return;
    const r = container.getBoundingClientRect();
    let dy = 0;
    const fromTop = dragPoint.y - r.top;
    const fromBottom = r.bottom - dragPoint.y;
    // Speed ramps 1→16 px/frame from the hot zone's outer edge to the edge.
    if (fromTop < SCROLL_EDGE_PX)
      dy = -Math.ceil((1 - fromTop / SCROLL_EDGE_PX) * SCROLL_MAX_SPEED_PX);
    else if (fromBottom < SCROLL_EDGE_PX)
      dy = Math.ceil((1 - fromBottom / SCROLL_EDGE_PX) * SCROLL_MAX_SPEED_PX);
    if (dy !== 0) {
      container.scrollTop += dy;
      const head = posAtClient(dragPoint.x, dragPoint.y, true);
      if (head != null) setDragSelection(dragAnchor, head);
      scrollRaf = requestAnimationFrame(dragAutoScroll);
    }
  };
  const startDragAutoScroll = (): void => {
    if (!scrollRaf) scrollRaf = requestAnimationFrame(dragAutoScroll);
  };
  const stopDragAutoScroll = (): void => {
    if (scrollRaf) cancelAnimationFrame(scrollRaf);
    scrollRaf = 0;
  };

  /** The drag's moving edge. Word's cross-cell drag: when anchor and head
   *  sit in different cells of one table the selection rectangle becomes a
   *  cell selection (every crossed cell selected whole); within one cell it
   *  stays a text selection. */
  const setDragSelection = (anchorPos: number, headPos: number): void => {
    const { doc } = active().editor.state;
    const anchorCell = cellAt(doc.resolve(anchorPos));
    const headCell = cellAt(doc.resolve(headPos));
    if (
      anchorCell &&
      headCell &&
      anchorCell.pos !== headCell.pos &&
      inSameTable(anchorCell, headCell)
    ) {
      active().editor.commands.command(({ state, dispatch }) => {
        dispatch?.(
          state.tr.setSelection(
            CellSelection.create(state.doc, anchorCell.pos, headCell.pos) as never,
          ),
        );
        return true;
      });
      return;
    }
    setSel(headPos, anchorPos);
  };

  // Word's link surface: hovering a link shows its URL with the Ctrl+Click
  // hint, and Ctrl+Click follows it — a plain click keeps dropping a caret.
  const linkAt = (pos: number): LinkHit | null => {
    const mark = active()
      .editor.state.doc.resolve(pos)
      .marks()
      .find((m) => m.type.name === "link");
    if (!mark) return null;
    const a = mark.attrs as { href?: unknown; target?: unknown };
    return typeof a.href === "string" && a.href
      ? { href: a.href, target: typeof a.target === "string" ? a.target : null }
      : null;
  };
  const linkHover = installLinkHover({
    host: opts.inputHost,
    posAtClient: (x, y) => posAtClient(x, y),
    linkAt,
  });
  opts.host.addEventListener("mouseleave", linkHover.hide);
  const chartHover = installChartHover({ host: opts.inputHost });
  opts.host.addEventListener("mouseleave", chartHover.hide);

  // The pointer cursor's single owner — Word's cursors per surface: a
  // floating drawing shows the four-headed move arrow, an inline picture the
  // plain arrow (never the text I-beam — an object is not text), a link the
  // hand, the text surface the I-beam (the canvas's stylesheet rule; writing
  // "" falls back to it). One writer computing the final value per move —
  // two writers with their own memory leave stale blanks behind each other
  // (the link tooltip's hide cleared the move cursor every move once, and
  // the drawing side's "unchanged, skip" memory never rewrote it). A
  // furniture story deactivates the body's objects, so their cursors go
  // with it (the story's own links keep the hand).
  const applyCursor = (event: MouseEvent): void => {
    // The armed painter owns the cursor: crosshair for the pen, the dense
    // cell cross for the eraser (Word's pencil/eraser, CSS-native), the
    // brush I-beam for the format painter.
    const paint = story ? null : opts.borderPaint?.();
    if (paint?.active) {
      opts.host.style.cursor = paint.eraser ? "cell" : "crosshair";
      return;
    }
    if (!story && opts.formatPaint?.()) {
      opts.host.style.cursor = FORMAT_PAINTER_CURSOR;
      return;
    }
    const hit = story ? null : hitPage(event.clientX, event.clientY);
    const drawHit = hit && opts.drawingAt ? opts.drawingAt(hit.page, hit.lx, hit.ly) : null;
    let want = "";
    if (drawHit) {
      // A chart sub-element reads as selectable content, not a movable
      // object — the plain arrow until the drag gestures reach it.
      want = drawHit.chartPart ? "default" : drawHit.kind === "drawing" ? "move" : "default";
    } else {
      const pos = posAtClient(event.clientX, event.clientY);
      const link = pos != null ? linkAt(pos) : null;
      if (link?.href) want = "pointer";
    }
    opts.host.style.cursor = want;
  };

  // The hovered chart element's ScreenTip data — the series/point read off
  // the chart node the hit pairs with (null off-chart or unpairable).
  const chartTipAt = (event: MouseEvent): ChartTip | null => {
    const hit = story ? null : hitPage(event.clientX, event.clientY);
    const drawHit = hit && opts.drawingAt ? opts.drawingAt(hit.page, hit.lx, hit.ly) : null;
    const part = drawHit?.chartPart;
    if (!drawHit || !part || part.title || part.series == null) return null;
    const nodePos = draw.nodePosOf(drawHit);
    const node = nodePos != null ? main.editor.state.doc.nodeAt(nodePos) : null;
    const chart = node?.attrs?.chart as Record<string, unknown> | undefined;
    if (!chart) return null;
    const series = (
      chart.series as { name?: string; values?: number[]; xValues?: number[] }[] | undefined
    )?.[part.series];
    const at = part.point ?? -1;
    return {
      name: series?.name ?? `${t("chart.series", opts.inputHost)} ${part.series + 1}`,
      ...(part.point != null
        ? {
            category: (chart.categories as string[] | undefined)?.[at],
            value: series?.values?.[at] ?? series?.xValues?.[at],
          }
        : {}),
    };
  };

  const onMouseMove = (event: MouseEvent): void => {
    if (borderSweep) {
      const hit = story ? null : hitPage(event.clientX, event.clientY);
      const edges = hit ? main.map?.tableEdgeAt(hit.page, hit.lx, hit.ly) : null;
      if (edges) for (const s of edges.sides) borderSweep.set(`${s.pos}:${s.side}`, s);
      return;
    }
    if (dragAnchor == null) {
      hoverTableGrip(event);
      linkHover.onMove(event);
      chartHover.onMove(chartTipAt(event), event.clientX, event.clientY);
      applyCursor(event);
      return;
    }
    if (
      !dragMoved &&
      dragStart &&
      Math.hypot(event.clientX - dragStart.x, event.clientY - dragStart.y) < 3
    ) {
      return;
    }
    dragMoved = true;
    dragPoint.x = event.clientX;
    dragPoint.y = event.clientY;
    startDragAutoScroll();
    // Clamp: a drag overshooting past a line/page edge keeps extending to the
    // nearest text (Word) — unclamped hits would stall the head mid-line.
    const head = posAtClient(event.clientX, event.clientY, true);
    if (head != null) setDragSelection(dragAnchor, head);
  };
  const onMouseUp = (): void => {
    if (borderSweep) {
      const sides = [...borderSweep.values()];
      borderSweep = null;
      if (sides.length) opts.applyBorderPaint?.(sides);
      return;
    }
    dragAnchor = null;
    dragMoved = false;
    dragStart = null;
    stopDragAutoScroll();
  };
  document.addEventListener("mousemove", onMouseMove);
  document.addEventListener("mouseup", onMouseUp);

  /** Enter a furniture story: a second viewless editor over the slot's
   *  JSON, a caret map over its laid stack (wrapped as one pseudo page
   *  anchored at the band), the caret dropped at the story's end. */
  const enterStory = (
    kind: StoryKind,
    page: number,
    geoIn: NonNullable<ReturnType<NonNullable<EditBridgeStory["geometry"]>>>,
  ): boolean => {
    const storyCfg = opts.story;
    // A read-only document has no stories (viewless editing is command-driven
    // — setEditable(false) alone cannot stop a transaction).
    if (!storyCfg || story || composing || !geoIn.band || !main.editor.isEditable) return false;
    const source = storyCfg.read(kind, geoIn.slot, page);
    // An empty story starts with one empty paragraph (Word's blank header).
    const initial = source.length > 0 ? source : [{ type: "paragraph" }];
    // Register the story first — the host's onDoc guards on its own story
    // state, and an empty slot needs one onDoc pass below to lay the strut.
    storyCfg.entered(kind, geoIn.slot, page);
    // An empty slot has no laid stack, so the caret map would have nowhere to
    // land — project the initial paragraph (the host's onDoc lays the strut
    // into the furniture) and re-read the geometry. updateStoryMap is a no-op
    // at this point (the story isn't registered yet); the refreshed geometry
    // below is what carries the stack into the new map.
    let geo = geoIn;
    if (!geo.stack) {
      storyCfg.onDoc(kind, geo.slot, initial);
      const refreshed = storyCfg.geometry(kind, page);
      if (!refreshed?.band || !refreshed.stack) return false;
      geo = refreshed;
    }
    const s = makeStory(
      { type: "doc", content: initial },
      (json) => storyCfg.onDoc(kind, geo.slot, (json.content ?? []) as JSONContent[]),
      page,
    );
    s.kind = kind;
    s.slot = geo.slot;
    // Dirty baseline AFTER schema normalization: the editor's getJSON is the
    // schema-filled shape (paragraph attrs expanded), not the lean `initial`
    // — comparing against that would flag every untouched story dirty.
    s.initialJson = JSON.stringify(s.editor.getJSON().content);
    if (geo.stack && geo.band) {
      // Local consts: `geo` is captured by the makeStory closure above, so its
      // narrowing doesn't survive to these property accesses.
      const band = geo.band;
      const origin = main.pageOrigin?.(s.anchorPage);
      s.map = new CaretMap([{ items: geo.stack }] as never, s.editor.state.doc, () => ({
        contentLeftPx: origin?.contentLeftPx ?? 0,
        contentTopPx: band.paintY,
      }));
    }
    attachTransactions(s);
    story = s;
    // The caret enters at the story's end (Word drops you after the text) —
    // the last textblock's end, not the doc's outer boundary (no caret there).
    setSel(TextSelection.atEnd(s.editor.state.doc).from);
    return true;
  };

  /** Tear the furniture story down and hand its final JSON to the host.
   *  The main story's overlays re-place against the geometry that is
   *  already there (the host's exit handler re-renders if dirty). */
  const leaveStory = (): {
    kind: StoryKind;
    slot: StorySlot;
    json: JSONContent[];
    dirty: boolean;
  } | null => {
    const s = story;
    if (!s) return null;
    story = null;
    if (s.raf) cancelAnimationFrame(s.raf);
    const json = (s.editor.getJSON().content ?? []) as JSONContent[];
    const dirty = JSON.stringify(json) !== s.initialJson;
    s.editor.destroy();
    opts.story?.exit({ kind: s.kind!, slot: s.slot!, json, dirty });
    placeCaret();
    return { kind: s.kind!, slot: s.slot!, json, dirty };
  };

  // Word's entry click: a DOUBLE click on a furniture band opens its story
  // (single clicks there are inert); while a story is active, clicks inside
  // its band position the story caret and any other click closes it.
  let lastClick: { t: number; x: number; y: number; count: number; key: string } | null = null;
  // Identity tags for the opaque hit paragraphs — template-stringifying them
  // directly collapses every paragraph onto "[object Object]", blurring the
  // double-click window across drawings.
  const paraTags = new WeakMap<object, number>();
  let paraTagSeq = 0;
  const tagOf = (para: unknown): number => {
    let tag = paraTags.get(para as object);
    if (tag == null) paraTags.set(para as object, (tag = ++paraTagSeq));
    return tag;
  };
  // The double-click window counts clicks on the SAME target — a click that
  // moved between the body and a drawing (or between drawings) restarts, so
  // an edit-session click followed by a shape click never reads as a double.
  const clickCount = (event: MouseEvent, key: string): number => {
    const again =
      lastClick != null &&
      lastClick.key === key &&
      event.timeStamp - lastClick.t < 500 &&
      Math.hypot(event.clientX - lastClick.x, event.clientY - lastClick.y) < 4;
    const count = again ? lastClick!.count + 1 : 1;
    lastClick = { t: event.timeStamp, x: event.clientX, y: event.clientY, count, key };
    return count;
  };
  // The armed Set Transparent Color callback (null = disarmed) — see
  // setTransparentPick in the public surface.
  let transparentPick: ((hit: DrawingHit, nx: number, ny: number) => void) | null = null;

  const takeFocus = (event: MouseEvent): void => {
    // Overlay widgets (the floating comment compose) own their focus — a
    // click inside one must not be dragged back to the input textarea.
    const path = event.composedPath() as HTMLElement[];
    if (path.some((el) => el instanceof HTMLElement && el.hasAttribute?.("data-docen-overlay")))
      return;
    // preventDefault keeps the click from blurring on mousedown; the caret
    // placement below is the real focus move.
    event.preventDefault();
    // A right-click only opens the context menu — its mousedown must not
    // disturb the selection (Word keeps a selection right-clicked inside it;
    // clicking elsewhere moves the caret from the menu handler, not here).
    if (event.button === 2) return;
    if (composing) return;
    linkHover.hide();
    // Park the textarea at the click point BEFORE focusing it (anchors the
    // IME window at the click; it no longer sits in the scroll container, so
    // focus() cannot yank the surface anymore, but parking stays harmless).
    const hostRect = opts.inputHost.getBoundingClientRect();
    ta.style.left = `${event.clientX - hostRect.left}px`;
    ta.style.top = `${event.clientY - hostRect.top}px`;
    const storyCfg = opts.story;
    const hit = hitPage(event.clientX, event.clientY);
    const drawHit = hit && opts.drawingAt ? opts.drawingAt(hit.page, hit.lx, hit.ly) : null;
    // Set Transparent Color: the armed pick samples the clicked drawing's
    // pixel instead of running the select chains (Word's eyedropper press);
    // a press off any drawing disarms and falls through as a plain click.
    if (transparentPick && drawHit && hit) {
      const onPick = transparentPick;
      transparentPick = null;
      // Un-rotate the point into the box's own space (the hit test's
      // formula), then normalize against the painted box.
      let px = hit.lx;
      let py = hit.ly;
      if (drawHit.rotation) {
        const rad = (-drawHit.rotation * Math.PI) / 180;
        const cx = drawHit.x + drawHit.width / 2;
        const cy = drawHit.y + drawHit.height / 2;
        const dx = px - cx;
        const dy = py - cy;
        px = cx + dx * Math.cos(rad) - dy * Math.sin(rad);
        py = cy + dx * Math.sin(rad) + dy * Math.cos(rad);
      }
      const nx = Math.min(Math.max((px - drawHit.x) / drawHit.width, 0), 1);
      const ny = Math.min(Math.max((py - drawHit.y) / drawHit.height, 0), 1);
      onPick(drawHit, nx, ny);
      ta.focus();
      ta.value = "";
      return;
    }
    const clicks = clickCount(
      event,
      drawHit ? `${tagOf(drawHit.para)}/${drawHit.index}/${drawHit.kind}` : "text",
    );
    const dbl = clicks >= 2;
    if (storyCfg && hit) {
      const header = storyCfg.geometry("header", hit.page);
      const footer = storyCfg.geometry("footer", hit.page);
      const inHeader =
        header?.band != null && hit.ly >= header.band.top && hit.ly < header.band.bottom;
      const inFooter =
        footer?.band != null && hit.ly >= footer.band.top && hit.ly < footer.band.bottom;
      if (story) {
        const own = story.kind === "header" ? inHeader : inFooter;
        // The anchor page's band edits in place; any other click (another
        // page, the body, the other story's band) closes the story — Word's
        // "double-click the body" exit, single-clicked.
        if (own && hit.page === story.anchorPage) {
          const pos = posAtClient(event.clientX, event.clientY);
          if (pos != null) clickSelection(pos, event, clicks);
        } else {
          leaveStory();
          const pos = posAtClient(event.clientX, event.clientY);
          if (pos != null) clickSelection(pos, event, clicks);
        }
        ta.focus();
        ta.value = "";
        return;
      }
      // A band click on an OBJECT routes by the object's story instead (a
      // body-anchored float dragged into the band is still a body object —
      // Word grabs it; the band owns only its empty regions): fall through
      // to the drawing chains below.
      if ((inHeader || inFooter) && !dbl && !drawHit) {
        // A single click on a band does nothing (Word), but must not blur
        // into a body position under it.
        ta.focus();
        ta.value = "";
        return;
      }
      if (inHeader && header && !drawHit) {
        enterStory("header", hit.page, header);
        ta.focus();
        ta.value = "";
        return;
      }
      if (inFooter && footer && !drawHit) {
        enterStory("footer", hit.page, footer);
        ta.focus();
        ta.value = "";
        return;
      }
    } else if (story) {
      leaveStory();
    }
    // The border painter's press (Table Design → Draw Border): a hit on a
    // table edge starts a sweep — moves collect further crossed edges, the
    // release commits them as one command. A press off any edge falls
    // through (the painter paints nothing outside a table).
    if (!story && opts.borderPaint?.().active && hit) {
      const edges = main.map?.tableEdgeAt(hit.page, hit.lx, hit.ly) ?? null;
      if (edges) {
        borderSweep = new Map();
        for (const s of edges.sides) borderSweep.set(`${s.pos}:${s.side}`, s);
        ta.focus();
        ta.value = "";
        return;
      }
    }
    // A grip click selects Word-style (column strip / row strip / corner
    // square) instead of dropping a caret. The hover state is refreshed for
    // a click that lands without a prior move reaching the strip.
    if (!story) {
      hoverTableGrip(event);
      if (grip.kind) {
        applyGrip();
        ta.focus();
        ta.value = "";
        return;
      }
    }
    // Body editing (the main story). A click landing on a drawing grabs it
    // (Word's picture selection) instead of dropping a caret behind the art;
    // any other click drops a standing drawing selection first.
    if (drawHit) {
      // Word: Shift+Click toggles a floating drawing into the multi-selection
      // (the group/distribute/align set) — a hit it declines (inline art, an
      // unpairable box) falls through to the plain click paths below.
      if (event.shiftKey && draw.toggleMulti(drawHit)) {
        ta.focus();
        ta.value = "";
        return;
      }
      // Word: a double click on a text-carrying shape enters its text body;
      // while a shape is being edited a click inside it moves the caret, and
      // a double click on a plain shape still just selects. "Editing" means
      // the caret sits strictly INSIDE the shape's span — a NodeSelection on
      // the shape (a fresh insert, Escape's exit) or a selection resting on
      // the span edge is selection state, so the next click selects (Word's
      // single-click-selects), never silently re-enters. The double click is
      // judged BEFORE the move press: it is an instant verdict (the second
      // tap inside 500ms), while a move only commits on a real drag.
      const pmSel = main.editor.state.selection;
      const span = main.map?.shapeAtPos(pmSel.from) ?? null;
      const editing =
        span && pmSel instanceof TextSelection && pmSel.from > span.from ? span : null;
      if (dbl || editing) {
        const pos =
          main.map?.posAtShapePoint(drawHit.page, drawHit.x, drawHit.y, editing ?? undefined) ??
          null;
        if (pos != null) {
          draw.clear();
          draw.place();
          setSel(pos);
          ta.focus();
          ta.value = "";
          return;
        }
        // A double click whose point the map cannot place (the shape's stack
        // failed to re-register after a relayout) still enters the text body:
        // fall back to the first caret position inside the hit shape rather
        // than dropping a body caret under the art.
        if (dbl) {
          const nodePos = draw.nodePosOf(drawHit);
          const node = nodePos != null ? main.editor.state.doc.nodeAt(nodePos) : null;
          if (nodePos != null && node?.type.name === "wpsShape" && node.textContent) {
            const inner = TextSelection.near(main.editor.state.doc.resolve(nodePos + 1), 1).from;
            draw.clear();
            draw.place();
            setSel(inner);
            ta.focus();
            ta.value = "";
            return;
          }
        }
      }
      // Word: a press on the already-selected offset-anchored floating
      // drawing starts a move drag (the frame trails the pointer, release
      // writes the new offsets); the drawing itself stays put until the
      // drag commits. Any other press (another drawing, an align-anchored
      // float, or the same one after a re-layout cleared the frame) just
      // selects.
      const sel = draw.selected;
      const same =
        sel &&
        drawHit.para === sel.para &&
        drawHit.index === sel.index &&
        drawHit.kind === sel.kind &&
        sameChildPath(drawHit.childPath, sel.childPath);
      const movable = same && draw.frameActive && draw.movableFloating();
      // Word's two-stage chart selection: with the chart framed, a click on
      // a plot element selects it (the series, then the point) — but a drag
      // from a plot element moves the whole floating chart like any drawing,
      // so the press arms the shared move gesture and the sub-selection only
      // lands on a clean click. The frame's dead space keeps the plain move.
      if (drawHit.chartPart && draw.chartPartOn(drawHit)) {
        // A value-draggable element (a bar, a line point) drags to change its
        // value (Excel); a clean click still sub-selects. Everything else
        // moves the chart like any drawing, sub-selecting on a plain click.
        if (
          draw.beginValueDrag(drawHit, event.clientX, event.clientY, () =>
            draw.selectChartPart(drawHit),
          )
        ) {
          // the gesture owns the press
        } else if (movable)
          draw.beginMove(event.clientX, event.clientY, () => draw.selectChartPart(drawHit));
        else draw.selectChartPart(drawHit);
        ta.focus();
        ta.value = "";
        return;
      }
      if (movable) {
        draw.beginMove(event.clientX, event.clientY);
        ta.focus();
        ta.value = "";
        return;
      }
      // A drawing the PM side cannot pair (furniture-anchored art whose host
      // paragraph lives outside the main doc, a stale box after a re-layout)
      // must not swallow the click — fall through to the text placement
      // (Word: body clicks pass through header-anchored shapes). A double
      // click enters a group (the member hit targets the member); a single
      // click keeps selecting the group as a whole.
      if (draw.select(drawHit, dbl)) {
        ta.focus();
        ta.value = "";
        return;
      }
    }
    draw.clear();
    draw.place();
    const pos = posAtClient(event.clientX, event.clientY);
    if (pos != null) {
      // Word's Ctrl+Click follows the link instead of dropping a caret; a
      // plain click keeps its editing meaning, and in viewing mode (read-only)
      // a plain click follows too. A `#name` href is a bookmark anchor —
      // followLink declines it, so the host resolves the jump.
      const link = linkAt(pos);
      const follow = event.ctrlKey || event.metaKey || !active().editor.isEditable;
      if (
        follow &&
        link &&
        (link.href.startsWith("#")
          ? (opts.onInternalAnchor?.(link.href.slice(1)), true)
          : followLink(link))
      ) {
        ta.focus();
        ta.value = "";
        return;
      }
      clickSelection(pos, event, clicks);
    }
    ta.focus();
    ta.value = "";
  };
  opts.host.addEventListener("mousedown", takeFocus);

  let composing = false;

  const insertText = (text: string): void => {
    if (text.includes("\n")) {
      const lines = text.split(/\r?\n/);
      const paragraphs = lines.map((line) => ({
        type: "paragraph",
        content: line ? [{ type: "text", text: line }] : [],
      }));
      active().editor.commands.insertContent(paragraphs);
      const storage = active().editor.storage as { repeat?: string; repeatLive?: boolean };
      storage.repeat = undefined;
      storage.repeatLive = false;
      return;
    }
    repeatInsert = text;
    active().editor.commands.command(({ state, dispatch }) => {
      const { from, to } = state.selection;
      dispatch?.(state.tr.insertText(text, from, to));
      return true;
    });
    // The transaction listener consumes the marker; this only covers a
    // refused command (no dispatch, no event).
    repeatInsert = null;
  };

  const backspace = (word = false): void => {
    active().editor.commands.command(({ state, dispatch }) => {
      const { selection } = state;
      if (!selection.empty) {
        dispatch?.(state.tr.deleteSelection());
        return true;
      }
      const $from = selection.$from;
      if ($from.parent.isTextblock) {
        // A non-text inline atom (picture, shape, break) contributes nothing
        // to textContent, so slicing textContent by the content offset
        // misaligns once the paragraph holds one — the cut lands off the real
        // characters and every later press no-ops (the stranded-caret jam).
        // Such an atom at the cut edge deletes outright (Word: one Backspace
        // removes the image); text nodes are leaves too (isAtom), so they are
        // excluded and zero-width atoms fall through to the text cut.
        // Otherwise textBetween maps the content offset to text units.
        const before = $from.nodeBefore;
        if (
          $from.parentOffset > 0 &&
          before &&
          !before.isText &&
          before.isAtom &&
          before.nodeSize > 0
        ) {
          dispatch?.(state.tr.delete($from.pos - before.nodeSize, $from.pos));
          return true;
        }
        const textBefore = $from.parent.textBetween(0, $from.parentOffset);
        if (textBefore) {
          const cut = word
            ? wordUnitsBackward(textBefore, textBefore.length)
            : lastGraphemeUnits(textBefore);
          if (cut > 0) {
            dispatch?.(state.tr.delete($from.pos - cut, $from.pos));
            return true;
          }
        }
      }
      // Word list backspace: at offset 0 of a list item, outdent if level > 0,
      // or clear list formatting if level === 0 (do not merge into previous block).
      if ($from.parentOffset === 0 && $from.parent.type.name === "paragraph") {
        const attrs = $from.parent.attrs as Record<string, unknown>;
        const bullet = attrs.bullet as { level?: number } | null | undefined;
        const numbering = attrs.numbering as
          | { reference?: string; level?: number }
          | null
          | undefined;
        if (bullet || numbering) {
          const level = bullet?.level ?? numbering?.level ?? 0;
          if (level > 0) {
            const patch = listLevelStepPatch(attrs, -1);
            if (patch) {
              dispatch?.(state.tr.setNodeMarkup($from.before(), undefined, { ...attrs, ...patch }));
              return true;
            }
          } else {
            dispatch?.(
              state.tr.setNodeMarkup($from.before(), undefined, {
                ...attrs,
                bullet: null,
                numbering: null,
              }),
            );
            return true;
          }
        }
      }
      // Same runtime PM instance (single .pnpm dir); the cast bridges the
      // dual d.ts identity between this package's @tiptap/pm and the engine's.
      // The selectNode tail is the stock backspace chain's: at a block start
      // facing a table it selects the table (Word) instead of dead-ending.
      return joinBackward(state as never, dispatch) || selectNodeBackward(state as never, dispatch);
    });
  };

  const deleteForward = (word = false): void => {
    active().editor.commands.command(({ state, dispatch }) => {
      const { selection } = state;
      if (!selection.empty) {
        dispatch?.(state.tr.deleteSelection());
        return true;
      }
      const $from = selection.$from;
      if ($from.parent.isTextblock) {
        // The forward mirror of backspace's atom handling: a non-text atom at
        // the cut edge deletes outright; the text cut reads units past the
        // caret through textBetween (atom-blind textContent misaligns the
        // same way).
        const after = $from.nodeAfter;
        if (after && !after.isText && after.isAtom && after.nodeSize > 0) {
          dispatch?.(state.tr.delete($from.pos, $from.pos + after.nodeSize));
          return true;
        }
        const textAfter = $from.parent.textBetween($from.parentOffset, $from.parent.content.size);
        if (textAfter) {
          const cut = word ? wordUnitsForward(textAfter, 0) : firstGraphemeUnits(textAfter);
          if (cut > 0) {
            dispatch?.(state.tr.delete($from.pos, $from.pos + cut));
            return true;
          }
        }
      }
      return joinForward(state as never, dispatch) || selectNodeForward(state as never, dispatch);
    });
  };

  /** Delete from the caret to a boundary target (delete-to-line-edge family:
   *  the target is the same edge Home/End resolve to). */
  const deleteTo = (toEnd: boolean): void => {
    const state = active().editor.state;
    const target = edgeTarget(state, state.selection.head, toEnd);
    if (target == null || target === state.selection.head) return;
    active().editor.commands.command(({ state: s, dispatch }) => {
      const head = s.selection.head;
      dispatch?.(s.tr.delete(Math.min(head, target), Math.max(head, target)));
      return true;
    });
  };

  const onBeforeInput = (event: InputEvent): void => {
    // During composition the browser owns the textarea's text (IME preview,
    // candidate navigation). Preventing insertCompositionText breaks that
    // management — the final text is taken from ta.value on compositionend.
    if (composing) return;
    // Viewing mode refuses text entry (the bridge textarea is invisible but
    // focused — without this gate typing would still mutate the doc).
    if (!active().editor.isEditable) {
      event.preventDefault();
      return;
    }
    event.preventDefault();
    switch (event.inputType) {
      case "insertText": {
        if (!event.data) break;
        // Autocorrect rides the typed leg only — the last character before
        // the caret decides the smart quote's side, and a boundary character
        // corrects the word behind it. IME/paste/spell legs skip this. The
        // Markdown input mode preempts it on the same leg: a block sequence
        // or a closed inline delimiter converts instead, and a leading
        // hyphen run stays literal (the horizontal-rule sequence must
        // survive autocorrect's em-dash rewrite). Off-mode autocorrect
        // behaves exactly as before.
        const s = active();
        const sel = s.editor.state.selection;
        const $from = sel.$from;
        const textBefore = $from.parent.textBetween(0, $from.parentOffset);
        const md = !sel.empty ? false : (opts.markdown?.() ?? false);
        const block = md ? blockRuleOf(event.data, textBefore) : null;
        const inline = !block && md ? inlineRuleOf(event.data, textBefore) : null;
        if (block) {
          s.editor.commands.command(({ state, dispatch }) => {
            const { from } = state.selection;
            const attrs = $from.parent.attrs as Record<string, unknown>;
            const target: Record<string, unknown> = {
              ...attrs,
              heading: null,
              style: null,
              bullet: null,
              numbering: null,
              thematicBreak: null,
            };
            if (block.kind === "heading") target.heading = HEADING_COMPILE_MAP[block.level];
            else if (block.kind === "quote") target.style = "IntenseQuote";
            else if (block.kind === "bullet") target.bullet = { level: 0 };
            else {
              target.numbering = {
                reference: nextOrderedReference(
                  collectListReferences(state.doc),
                  (state.doc.attrs as { numbering?: unknown }).numbering,
                ),
                level: 0,
                start: block.start,
              };
            }
            const tr = state.tr;
            tr.insertText("", from - textBefore.length, from);
            tr.setNodeMarkup($from.before(), undefined, target);
            dispatch?.(tr.scrollIntoView());
            return true;
          });
        } else if (inline) {
          s.editor.commands.command(({ state, dispatch }) => {
            const { from, to } = state.selection;
            const base = from - textBefore.length;
            const openStart = base + inline.openStart;
            const tr = state.tr;
            // The typed character is never inserted: the closing delimiter
            // loses only its earlier half (already in textBefore).
            tr.insertText("", from - (inline.openLen - 1), to);
            tr.insertText("", openStart, openStart + inline.openLen);
            tr.addMark(
              openStart,
              openStart + inline.innerLen,
              state.schema.marks[inline.mark]!.create(),
            );
            dispatch?.(tr.scrollIntoView());
            return true;
          });
        } else if (md && event.data === "-" && isHyphenRun(textBefore)) {
          insertText(event.data);
        } else {
          const fix = autocorrectOf(event.data, textBefore);
          if (fix) {
            s.editor.commands.command(({ state, dispatch }) => {
              const { from, to } = state.selection;
              dispatch?.(state.tr.insertText(fix.text, from - fix.back, to));
              return true;
            });
          } else {
            insertText(event.data);
          }
        }
        break;
      }
      // Chrome reports textarea Enter as insertLineBreak; keep both mapped to
      // a paragraph split (Shift+Enter variants included for now). PM's
      // splitBlock drops attrs — a list paragraph would lose its
      // bullet/numbering on every Enter — so the split re-applies the
      // paragraph's attrs (minus the section-close markers, which belong to
      // the paragraph closing the section). An EMPTY list paragraph exits the
      // list instead (Word: Enter on an empty item ends the list), and an
      // empty Code paragraph exits the code style the same way. With the
      // Markdown input mode on, a paragraph whose whole text is an opening
      // code fence converts to a Code paragraph instead of splitting, and a
      // horizontal-rule line converts to a thematic break followed by a fresh
      // paragraph for the caret.
      case "insertParagraph":
      case "insertLineBreak":
        active().editor.commands.command(({ state, dispatch }) => {
          const { $from, empty } = state.selection;
          const parent = $from.parent;
          if (parent.type.name !== "paragraph") {
            return splitBlock(state as never, dispatch);
          }
          const attrs = parent.attrs as Record<string, unknown>;
          const isList = attrs.bullet != null || attrs.numbering != null;
          if (isList && empty && parent.content.size === 0) {
            dispatch?.(
              state.tr.setNodeMarkup($from.before(), undefined, {
                ...attrs,
                bullet: null,
                numbering: null,
              }),
            );
            return true;
          }
          if (attrs.style === "Code" && empty && parent.content.size === 0) {
            dispatch?.(
              state.tr.setNodeMarkup($from.before(), undefined, {
                ...attrs,
                style: null,
                codeLanguage: null,
              }),
            );
            return true;
          }
          if (opts.markdown?.() ?? false) {
            const fence = enterRuleOf(parent.textContent);
            if (fence?.kind === "code") {
              const tr = state.tr;
              // The fence line itself never becomes code content.
              tr.delete($from.start(), $from.end());
              tr.setNodeMarkup($from.before(), undefined, {
                ...attrs,
                style: "Code",
                codeLanguage: fence.language,
              });
              dispatch?.(tr.scrollIntoView());
              return true;
            }
            if (fence?.kind === "hr" && dispatch) {
              const tr = state.tr;
              tr.delete($from.start(), $from.end());
              tr.setNodeMarkup($from.before(), undefined, { ...attrs, thematicBreak: true });
              const carried = { ...attrs };
              delete carried.sectionProperties;
              delete carried.sectionHeaders;
              delete carried.sectionFooters;
              delete carried.thematicBreak;
              tr.split(tr.mapping.map($from.pos), 1, [{ type: parent.type, attrs: carried }]);
              dispatch(tr.scrollIntoView());
              return true;
            }
          }
          if (dispatch) {
            const tr = state.tr;
            if (!empty) tr.deleteSelection();
            const carried = { ...attrs };
            delete carried.sectionProperties;
            delete carried.sectionHeaders;
            delete carried.sectionFooters;
            tr.split(tr.mapping.map($from.pos), 1, [{ type: parent.type, attrs: carried }]);
            dispatch(tr.scrollIntoView());
          }
          return true;
        });
        break;
      case "deleteContentBackward":
        backspace();
        break;
      case "deleteWordBackward":
        backspace(true);
        break;
      case "deleteContentForward":
        deleteForward();
        break;
      case "deleteWordForward":
        deleteForward(true);
        break;
      // Cmd/Ctrl+Backspace-adjacent line deletes (macOS reports these as soft
      // line deletes; Windows IMEs occasionally emit the hard variants). The
      // target is the wrapped line's edge — the same edge Home/End resolve to.
      case "deleteSoftLineBackward":
      case "deleteHardLineBackward":
        deleteTo(false);
        break;
      case "deleteSoftLineForward":
      case "deleteHardLineForward":
        deleteTo(true);
        break;
      // Spell-check corrections / autofill: the replacement text rides in
      // `data` or the dataTransfer (data is null on Chrome's context-menu
      // correction). Replacing a non-empty selection; at an empty caret the
      // target word is browser-internal, so fall back to plain insertion.
      case "insertReplacementText": {
        const text = event.data ?? event.dataTransfer?.getData("text/plain");
        if (text) insertText(text);
        break;
      }
      default:
        break;
    }
    ta.value = "";
  };

  /** One step horizontally from a position — a grapheme or word inside the block
   *  (atoms step a whole node), the nearest text position past its edge. */
  const hStep = (state: Editor["state"], pos: number, dir: -1 | 1, word = false): number | null => {
    const $from = state.doc.resolve(pos);
    const offset = $from.parentOffset;
    const size = $from.parent.content.size;
    const text = $from.parent.textBetween(0, size);
    if (dir < 0 && offset > 0) {
      const cut = word
        ? wordUnitsBackward(text, offset)
        : lastGraphemeUnits(text.slice(0, offset)) || 1;
      return pos - cut;
    }
    if (dir > 0 && offset < size) {
      const cut = word
        ? wordUnitsForward(text, offset)
        : firstGraphemeUnits(text.slice(offset)) || 1;
      return pos + cut;
    }
    const edge = dir < 0 ? $from.before() : $from.after();
    if (edge < 0 || edge > state.doc.content.size) return null;
    const $near = TextSelection.near(state.doc.resolve(edge), dir);
    return $near.from === pos ? null : $near.from;
  };

  /** A line's boundary position off the caret map (the only place the wrap
   *  geometry lives); unmapped falls back to the block's boundaries. */
  const edgeTarget = (state: Editor["state"], pos: number, toEnd: boolean): number => {
    const $from = state.doc.resolve(pos);
    const edges = active().map?.valid ? active().map!.lineEdges(pos) : null;
    if (edges) return toEnd ? edges.end : edges.home;
    return toEnd ? $from.end() : $from.start();
  };

  /** One line up/down at the goal column (null at the paragraph's edge). */
  const vStep = (pos: number, dir: -1 | 1): number | null => {
    const map = active().map;
    return map?.valid ? map.posVertical(pos, dir) : null;
  };

  /** Move the caret to a target — or extend the selection to it (the anchor
   *  holds, the head moves). */
  const apply = (target: number | null, extend: boolean): void => {
    if (target == null) return;
    if (extend) {
      setSel(target, active().editor.state.selection.anchor);
    } else {
      setSel(target);
    }
  };

  /** Finds the doc position of the next or previous table cell (reading order). */
  const adjacentCellPos = (
    state: Editor["state"],
    dir: -1 | 1,
  ): { pos: number; isLastInTable: boolean } | null => {
    const $from = state.selection.$from;
    const $cell = cellAt($from);
    if (!$cell) return null;
    const table = $cell.node($cell.depth - 1);
    const tableStart = $cell.before($cell.depth - 1);
    const cells: number[] = [];
    let rowPos = tableStart + 1;
    for (let r = 0; r < table.childCount; r++) {
      const row = table.child(r);
      let cellPos = rowPos + 1;
      for (let c = 0; c < row.childCount; c++) {
        cells.push(cellPos);
        cellPos += row.child(c).nodeSize;
      }
      rowPos += row.nodeSize;
    }
    const idx = cells.indexOf($cell.pos);
    if (idx < 0) return null;
    const targetIdx = idx + dir;
    if (targetIdx >= 0 && targetIdx < cells.length) {
      return { pos: cells[targetIdx]! + 1, isLastInTable: false };
    }
    return { pos: $cell.pos, isLastInTable: dir > 0 && targetIdx >= cells.length };
  };

  /** Scrolls the caret's page so the caret sits a third of the way down the
   *  workspace viewport — only when it is out of view (Home/End/PageUp/PageDown,
   *  find-next). A caret already visible keeps its position. */
  const scrollIntoView = (pos: number): void => {
    const rect = main.map?.valid ? main.map.caretRect(pos) : null;
    if (!rect) return;
    const pageEl = opts.pageHost?.(rect.page);
    if (!pageEl) return;
    const scale = opts.scale?.() ?? 1;
    const scrollParent =
      (pageEl.closest(
        "[data-docen-scroll-container], .workspace, .editor-viewport, [style*='overflow']",
      ) as HTMLElement | null) ?? pageEl.parentElement;
    if (scrollParent && scrollParent.scrollHeight > scrollParent.clientHeight) {
      const pageRect = pageEl.getBoundingClientRect();
      const parentRect = scrollParent.getBoundingClientRect();
      const caretY = pageRect.top + rect.yPx * scale;
      if (caretY < parentRect.top || caretY + rect.heightPx * scale > parentRect.bottom) {
        const targetScrollTop =
          scrollParent.scrollTop + (caretY - parentRect.top) - parentRect.height / 3;
        scrollParent.scrollTo({ top: Math.max(0, targetScrollTop), behavior: "auto" });
        return;
      }
    }
    pageEl.scrollIntoView({ block: "nearest", behavior: "auto" });
  };

  // Word: Delete/Backspace on a chart sub-selection edits it — remove the
  // sub-selected series, clear a sub-selected title; a sub-selected data
  // point declines (the number model has no gaps to collapse). False on a
  // bare frame, where the keys keep their node/text meaning.
  const chartSubDelete = (): boolean => {
    const series = draw.chartSeriesSelected();
    if (series != null) {
      // The command declines on the last series (the model needs one) — the
      // sub-selection only drops when the removal actually happened.
      if (active().editor.commands["chart-series-delete"](String(series))) draw.clearChartPart();
      return true;
    }
    if (draw.chartTitleSelected()) {
      active().editor.commands["chart-data-apply"]('{"title":""}');
      draw.clearChartPart();
      return true;
    }
    return draw.chartPartSelected;
  };

  const onKeyDown = (event: KeyboardEvent): void => {
    // The IME owns the keyboard during composition (candidate navigation,
    // commit keys) — our caret moves would cancel it mid-word.
    if (composing) return;
    // Viewing mode: caret moves and selection stay live (the READONLY_LIVE
    // ribbon set's keyboard counterpart), but nothing may mutate the doc.
    const editable = active().editor.isEditable;
    // A cell selection's delete: Word removes the table node itself when the
    // selection covers EVERY cell of it, and empties the selected cells
    // otherwise (the grid survives). The default join path would tear cell
    // content across the range, so this must run before it.
    if (editable && (event.key === "Backspace" || event.key === "Delete")) {
      const sel = active().editor.state.selection;
      if (sel instanceof NodeSelection) {
        event.preventDefault();
        // A chart sub-selection parks the NodeSelection on the chart node —
        // the keys edit the sub-selection (above); only a bare frame deletes
        // the node itself.
        if (sel.node.type.name === "chart" && chartSubDelete()) return;
        active().editor.commands.command(({ state, dispatch }) => {
          if (dispatch) dispatch(state.tr.deleteSelection());
          return true;
        });
        if (draw.selected != null) {
          draw.clear();
          draw.place();
        }
        return;
      }
      if (sel instanceof CellSelection) {
        event.preventDefault();
        active().editor.commands.command(({ state, dispatch }) => {
          if (!dispatch) return true;
          // The host table through the anchor cell (resolved at its start —
          // the walk lands in the row layer, so depth-1 is the table).
          const $a = state.doc.resolve(sel.anchorCell);
          const table = $a.node($a.depth - 1);
          let total = 0;
          table.forEach((row) => {
            total += row.childCount;
          });
          let picked = 0;
          sel.forEachCell(() => {
            picked += 1;
          });
          if (total > 0 && picked >= total) {
            dispatch(
              state.tr.delete($a.before($a.depth - 1), $a.before($a.depth - 1) + table.nodeSize),
            );
          } else {
            sel.replace(state.tr as never);
          }
          return true;
        });
        return;
      }
    }
    // Shift+Enter inserts a soft line break (w:br inside the paragraph) —
    // captured here because the textarea reports both Enter flavors to
    // beforeinput as the same insertLineBreak.
    if (editable && event.key === "Enter" && event.shiftKey && !event.ctrlKey && !event.metaKey) {
      event.preventDefault();
      active().editor.commands.command(({ state, dispatch }) => {
        if (dispatch) {
          const tr = state.tr;
          const br = state.schema.nodes.hardBreak?.create();
          if (!br) return false;
          if (!state.selection.empty) tr.deleteSelection();
          tr.insert(state.selection.from, br);
          dispatch(tr.scrollIntoView());
        }
        return true;
      });
      return;
    }
    // F4 (Word Repeat) — retype the last text insertion at the caret.
    if (event.key === "F4" && !event.ctrlKey && !event.metaKey && !event.altKey) {
      event.preventDefault();
      if (editable) {
        const repeat = (active().editor.storage as { repeat?: string }).repeat;
        if (repeat) insertText(repeat);
      }
      return;
    }
    if (event.ctrlKey || event.metaKey) {
      const key = event.key;
      const lower = key.toLowerCase();
      // Modifier combos dispatch commands (registered in KEYBOARD_SHORTCUTS).
      // Undo / Redo
      if (lower === "z") {
        event.preventDefault();
        if (!editable) return;
        if (event.shiftKey) active().editor.commands.redo();
        else active().editor.commands.undo();
        return;
      }
      if (lower === "y") {
        event.preventDefault();
        if (!editable) return;
        active().editor.commands.redo();
        return;
      }
      // Select all — without this the browser default selects the 1px
      // textarea's (empty) contents and the press is lost.
      if (lower === "a") {
        event.preventDefault();
        active().editor.commands.command(({ state, dispatch }) =>
          selectAll(state as never, dispatch),
        );
        return;
      }
      // Mod-K: Insert / edit hyperlink (Word standard).
      if (lower === "k") {
        event.preventDefault();
        opts.host.dispatchEvent(
          new CustomEvent("command", {
            bubbles: true,
            composed: true,
            detail: { event: "link" },
          }),
        );
        return;
      }
      // Mod-D: Font dialog (Word standard).
      if (lower === "d" && !event.shiftKey) {
        event.preventDefault();
        opts.host.dispatchEvent(
          new CustomEvent("command", {
            bubbles: true,
            composed: true,
            detail: { event: "font-dialog" },
          }),
        );
        return;
      }
      // Viewless editors have no EditorView, so nothing dispatches Tiptap's
      // per-extension keyboard shortcuts — match the shared table here (the
      // DocenKeymap extension serves the same table on a DOM route). Named
      // keys keep their spelling ("Mod-Enter"); single characters uppercase
      // ("Mod-B") — a blanket toUpperCase turned Enter into "ENTER" and
      // silently dead-matched the table.
      const combo = `Mod${event.shiftKey ? "-Shift" : ""}-${key.length === 1 ? lower.toUpperCase() : key}`;
      const command = KEYBOARD_SHORTCUTS[combo];
      if (command) {
        if (!editable) return;
        event.preventDefault();
        const [name, arg] = command.split(":");
        (
          active().editor.commands as unknown as Record<
            string,
            ((arg?: string) => boolean) | undefined
          >
        )[name]?.(arg);
      }
      return;
    }
    const extend = event.shiftKey;
    const head = () => active().editor.state.selection.head;
    switch (event.key) {
      case "ArrowLeft":
        event.preventDefault();
        if (draw.selected != null) {
          draw.clear();
          draw.place();
        }
        apply(hStep(active().editor.state, head(), -1, event.ctrlKey || event.metaKey), extend);
        break;
      case "ArrowRight":
        event.preventDefault();
        if (draw.selected != null) {
          draw.clear();
          draw.place();
        }
        apply(hStep(active().editor.state, head(), 1, event.ctrlKey || event.metaKey), extend);
        break;
      case "ArrowUp":
        event.preventDefault();
        apply(vStep(head(), -1), extend);
        break;
      case "ArrowDown":
        event.preventDefault();
        apply(vStep(head(), 1), extend);
        break;
      case "Home": {
        event.preventDefault();
        const target =
          event.ctrlKey || event.metaKey ? 0 : edgeTarget(active().editor.state, head(), false);
        apply(target, extend);
        scrollIntoView(target);
        break;
      }
      case "End": {
        event.preventDefault();
        const target =
          event.ctrlKey || event.metaKey
            ? active().editor.state.doc.content.size
            : edgeTarget(active().editor.state, head(), true);
        apply(target, extend);
        scrollIntoView(target);
        break;
      }
      case "PageUp": {
        event.preventDefault();
        const map = active().map;
        if (map?.valid) {
          const cur = head();
          const curPage = map.caretRect(cur)?.page ?? 0;
          const targetPage = Math.max(0, curPage - 1);
          const targetPos = map.firstPosOfPage(targetPage) ?? 0;
          apply(targetPos, extend);
          scrollIntoView(targetPos);
        }
        break;
      }
      case "PageDown": {
        event.preventDefault();
        const map = active().map;
        if (map?.valid) {
          const cur = head();
          const curPage = map.caretRect(cur)?.page ?? 0;
          const targetPos =
            map.firstPosOfPage(curPage + 1) ?? active().editor.state.doc.content.size;
          apply(targetPos, extend);
          scrollIntoView(targetPos);
        }
        break;
      }
      case "Delete":
        if (editable) {
          event.preventDefault();
          deleteForward(event.ctrlKey || event.metaKey);
        }
        break;
      // Leaving a furniture story (Word: Esc = Close Header and Footer).
      case "Escape":
        // The armed Set Transparent Color pick is the shallowest mode — Esc
        // disarms it before any other Escape meaning (Word).
        if (transparentPick) {
          transparentPick = null;
          event.preventDefault();
          return;
        }
        if (editable && main.map?.shapeAtPos(main.editor.state.selection.from)) {
          // Leaving a text box's edit mode selects the whole shape (Word) —
          // the span's `from` IS the wpsShape node's position.
          event.preventDefault();
          const span = main.map.shapeAtPos(main.editor.state.selection.from)!;
          main.editor.commands.command(({ state, dispatch }) => {
            dispatch?.(state.tr.setSelection(NodeSelection.create(state.doc, span.from) as never));
            return true;
          });
          return;
        }
        // Word's chart stepdown: point → series → the bare chart frame.
        if (draw.escapeChartPart()) {
          event.preventDefault();
          return;
        }
        if (draw.selected != null) {
          event.preventDefault();
          draw.clear();
          draw.place();
          // Word: Escape at the bare frame deselects entirely — the drawing
          // must not keep holding an invisible NodeSelection.
          const pmSel = active().editor.state.selection;
          if (pmSel instanceof NodeSelection) {
            setSel(TextSelection.near(active().editor.state.doc.resolve(pmSel.from)).from);
          }
          return;
        }
        if (story) {
          event.preventDefault();
          leaveStory();
        }
        break;
      // Tab / Shift+Tab: table cells navigate; list paragraphs adjust nesting; normal text inserts tab.
      case "Tab": {
        event.preventDefault();
        const { from, to, $from } = active().editor.state.selection;
        // 1. Table navigation: Tab moves to next cell; in the last cell, it inserts a new row below.
        const $cell = cellAt($from);
        if ($cell) {
          const adj = adjacentCellPos(active().editor.state, event.shiftKey ? -1 : 1);
          if (adj) {
            if (adj.isLastInTable) {
              if (editable) {
                active().editor.commands["insert-row-below"]?.();
              }
            } else {
              setSel(TextSelection.near(active().editor.state.doc.resolve(adj.pos)).from);
            }
          }
          break;
        }
        // 2. List paragraphs: Tab / Shift+Tab adjusts nesting level (Word).
        const patches: { pos: number; patch: Record<string, unknown> }[] = [];
        active().editor.state.doc.nodesBetween(from, to, (node, pos) => {
          if (node.type.name !== "paragraph") return true;
          const patch = listLevelStepPatch(
            node.attrs as Record<string, unknown>,
            event.shiftKey ? -1 : 1,
          );
          if (patch) patches.push({ pos, patch });
          return true;
        });
        if (patches.length > 0) {
          active().editor.commands.command(({ state, dispatch }) => {
            const tr = state.tr;
            for (const { pos, patch } of patches) {
              const live = tr.doc.nodeAt(pos);
              if (!live || live.type.name !== "paragraph") continue;
              tr.setNodeMarkup(pos, undefined, {
                ...(live.attrs as Record<string, unknown>),
                ...patch,
              });
            }
            dispatch?.(tr.scrollIntoView());
            return true;
          });
          break;
        }
        // 3. Normal paragraphs: insert tab character or node
        if (editable && !event.shiftKey) {
          if (active().editor.state.schema.nodes.tab) {
            active().editor.commands.command(({ state, dispatch }) => {
              dispatch?.(state.tr.replaceSelectionWith(state.schema.nodes.tab.create()));
              return true;
            });
          } else {
            insertText("\t");
          }
        }
        break;
      }
      default:
        break;
    }
  };

  const onCompositionStart = (): void => {
    composing = true;
  };
  const onCompositionEnd = (): void => {
    composing = false;
    const data = ta.value;
    ta.value = "";
    if (data) insertText(data);
  };
  // A cancelled composition (IME dismissed, focus stolen mid-composition —
  // paths where some browsers never fire compositionend) still must clear the
  // flag, or every input handler above stays gated off permanently.
  const onCompositionCancel = (): void => {
    composing = false;
    ta.value = "";
  };

  /** Insert pasted JSON at the caret, dropping stray empty text nodes the
   *  clipboard HTML can leave behind. Returns true when something landed. */
  const insertPastedJSON = (html: string): boolean => {
    const body = new DOMParser().parseFromString(html, "text/html").body;
    const json = parseHTMLBody(body, active().editor.state.schema);
    const content = (json.content ?? []).filter((n) => n.type !== "text" || n.text);
    if (!content.length) return false;
    active().editor.commands.insertContent(content);
    return true;
  };

  /** Insert a docen slice payload (the DOCEN_CLIP_MIME lane) at the caret —
   *  marks, node attrs, and open depths all survive. Returns true when the
   *  payload parsed and landed. */
  const insertSlicePayload = (raw: string): boolean => {
    try {
      const parsed = JSON.parse(raw) as { openStart?: number; openEnd?: number; content?: unknown };
      if (!Array.isArray(parsed.content)) return false;
      const { state, view } = active().editor;
      const slice = new Slice(
        Fragment.fromJSON(state.schema, parsed.content),
        parsed.openStart ?? 0,
        parsed.openEnd ?? 0,
      );
      view.dispatch(state.tr.replaceSelection(slice));
      return true;
    } catch {
      return false;
    }
  };

  const onPaste = (event: ClipboardEvent): void => {
    event.preventDefault();
    // The docen lane first (a copy from a docen editor round-trips losslessly);
    // then styled HTML through the schema's parse rules so external rich text
    // maps to its DOCX equivalents; plain text is the last resort. The custom
    // lane reads through BOTH spellings: the `web `-prefixed key is what the
    // system clipboard carries (Chromium's custom-format spec — and what the
    // async clipboard.read() below matches), the bare key is what a same-page
    // DataTransfer passthrough may still hand back.
    const data = event.clipboardData;
    const docen = data?.getData(`web ${DOCEN_CLIP_MIME}`) || data?.getData(DOCEN_CLIP_MIME);
    if (docen && insertSlicePayload(docen)) {
      const plain = data?.getData("text/plain") ?? "";
      opts.onRichPaste?.({ kind: "slice", raw: docen, text: plain });
      return;
    }
    const html = data?.getData("text/html");
    if (html && insertPastedJSON(html)) {
      opts.onRichPaste?.({ kind: "html", raw: html, text: data?.getData("text/plain") ?? "" });
      return;
    }
    const text = data?.getData("text/plain");
    if (text) {
      // A ribbon/context-menu copy wrote its custom format through the async
      // API, which a paste event never sees — recover the marks via the pin
      // when the plain text still matches it (stale after a copy elsewhere,
      // which is exactly when the fallback must not fire).
      if (lastCopied?.text === text && insertSlicePayload(lastCopied.payload)) return;
      insertText(text);
    }
  };

  const selectionText = (): string | null => {
    const { from, to } = active().editor.state.selection;
    return from === to ? null : active().editor.state.doc.textBetween(from, to, "\n");
  };

  /** The most recent copy/cut's slice payload. Chromium never persists a copy
   *  EVENT's custom types to the system clipboard (the `web ` spelling only
   *  works through the async write API), so a same-page paste via the async
   *  read() — the ribbon and context-menu Paste — can't see the lane. The
   *  payload is pinned here instead; the host's paste falls back to it when
   *  the system clipboard's plain text matches (stale after a copy elsewhere,
   *  which is exactly when the fallback must not fire). */
  let lastCopied: { payload: string; text: string } | null = null;

  /** Pin the current selection's copy pieces and register them for the
   *  fallback lane. Neither clipboard channel carries the custom format to a
   *  paste event — a copy EVENT's types never persist to the system clipboard,
   *  and the async write's `web ` format never reaches a paste event — so the
   *  pinned payload plus the matching plain text is what every paste path
   *  recovers marks through. */
  const pinCopied = (): { text: string; payload: string | null } | null => {
    const text = selectionText();
    if (text == null) return null;
    const payload = selectionSlicePayload(active().editor.state);
    lastCopied = payload ? { payload, text } : null;
    opts.onClipboardCollect?.({ text, payload });
    return { text, payload };
  };

  const onCopy = (event: ClipboardEvent): void => {
    const copied = pinCopied();
    if (!copied) return;
    event.preventDefault();
    event.clipboardData?.setData("text/plain", copied.text);
    // The `web ` prefix is Chromium's spelling for custom clipboard formats;
    // through a copy event it only survives same-page DataTransfer passthrough
    // (the async read() needs the pinned payload above).
    if (copied.payload) event.clipboardData?.setData(`web ${DOCEN_CLIP_MIME}`, copied.payload);
  };

  const onCut = (event: ClipboardEvent): void => {
    const copied = pinCopied();
    if (!copied) return;
    event.preventDefault();
    event.clipboardData?.setData("text/plain", copied.text);
    if (copied.payload) event.clipboardData?.setData(`web ${DOCEN_CLIP_MIME}`, copied.payload);
    active().editor.commands.command(({ state, dispatch }) => {
      dispatch?.(state.tr.deleteSelection());
      return true;
    });
  };

  /** Copy/cut for the entry points that produce no copy event (the ribbon and
   *  context-menu buttons — the selection is canvas-rendered). Writes the
   *  system clipboard through the async API (the custom format survives there
   *  for clipboard.read()-based pastes) and pins the payload so a keyboard
   *  paste — which cannot see either custom-format channel — still recovers
   *  the marks. */
  const copySelection = async (cut: boolean): Promise<void> => {
    const copied = pinCopied();
    if (!copied) return;
    try {
      await navigator.clipboard.write([
        new ClipboardItem({
          "text/plain": new Blob([copied.text], { type: "text/plain" }),
          ...(copied.payload
            ? { [`web ${DOCEN_CLIP_MIME}`]: new Blob([copied.payload], { type: DOCEN_CLIP_MIME }) }
            : {}),
        }),
      ]);
    } catch {
      try {
        await navigator.clipboard.writeText(copied.text);
      } catch {
        // Clipboard write may be denied (permissions/policy) — still cut.
      }
    }
    if (cut) {
      active().editor.commands.command(({ state, dispatch }) => {
        dispatch?.(state.tr.deleteSelection());
        return true;
      });
    }
  };

  ta.addEventListener("beforeinput", onBeforeInput);
  ta.addEventListener("keydown", onKeyDown);
  ta.addEventListener("compositionstart", onCompositionStart);
  ta.addEventListener("compositionend", onCompositionEnd);
  ta.addEventListener("compositioncancel", onCompositionCancel);
  ta.addEventListener("paste", onPaste);
  ta.addEventListener("copy", onCopy);
  ta.addEventListener("cut", onCut);
  opts.inputHost.append(ta);

  return {
    editor,
    updatePages(pages, pageOrigin): void {
      main.pageOrigin = pageOrigin;
      main.mapDoc = main.editor.state.doc;
      main.map = new CaretMap(
        pages,
        main.editor.state.doc,
        (page) => pageOrigin(page) ?? { contentLeftPx: 0, contentTopPx: 0 },
      );
      const stacks = opts.shapeTextStacks?.() ?? [];
      if (stacks.length)
        main.map.registerShapeStacks(stacks, (host) => opts.shapeResolve?.(host) ?? null);
      main.pageCount = pages.length;
      draw.place();
      placeCaret();
    },
    updateStoryMap(stack, band): void {
      const s = story;
      if (!s) return;
      const origin = main.pageOrigin?.(s.anchorPage);
      s.mapDoc = s.editor.state.doc;
      s.map = stack
        ? new CaretMap([{ items: stack }] as never, s.editor.state.doc, () => ({
            contentLeftPx: origin?.contentLeftPx ?? 0,
            contentTopPx: band.paintY,
          }))
        : null;
      placeCaret();
    },
    storyKind(): StoryKind | null {
      return story?.kind ?? null;
    },
    activeEditor(): Editor {
      return active().editor;
    },
    copiedSlice(): { payload: string; text: string } | null {
      return lastCopied;
    },
    copySelection(cut: boolean): Promise<void> {
      return copySelection(cut);
    },
    enterStory(kind, page, seed): boolean {
      const storyCfg = opts.story;
      if (!storyCfg) return false;
      const p =
        page ??
        (main.map?.valid ? (main.map.caretRect(main.editor.state.selection.from)?.page ?? 0) : 0);
      const geo = storyCfg.geometry(kind, p);
      if (!geo || !enterStory(kind, p, geo)) return false;
      if (seed) {
        // The caret already sits at the story's end (enterStory dropped it
        // there) — the seed lands right after it, the render riding the
        // story's own raf-merged onDoc.
        const s = story!;
        s.editor.commands.insertContentAt(TextSelection.atEnd(s.editor.state.doc).from, seed);
      }
      ta.focus();
      ta.value = "";
      return true;
    },
    exitStory() {
      return leaveStory();
    },
    scrollIntoView(pos): void {
      scrollIntoView(pos);
    },
    /** The page index a doc position renders on (null when unmappable). */
    pageOf(pos): number | null {
      return main.map?.valid ? (main.map.caretRect(pos)?.page ?? null) : null;
    },
    /** The first doc position rendered on a page (null when unmappable). */
    firstPosOfPage(page: number): number | null {
      return main.map?.valid ? main.map.firstPosOfPage(page) : null;
    },
    posOfPara(para): number | null {
      return main.map?.valid
        ? main.map.posOfPara(para as import("@docen/layout").LaidOutParagraph)
        : null;
    },
    posAtClient(clientX, clientY): number | null {
      return posAtClient(clientX, clientY);
    },
    selectDrawingAtClient(clientX, clientY) {
      const hit = hitPage(clientX, clientY);
      if (!hit || !opts.drawingAt) return false;
      const drawHit = opts.drawingAt(story ? 0 : hit.page, hit.lx, hit.ly);
      return drawHit ? draw.select(drawHit) : false;
    },
    enterCropMode(): boolean {
      return draw.enterCropMode();
    },
    setTransparentPick(onPick) {
      transparentPick = onPick;
    },
    drawingMulti() {
      return draw.multiPayload();
    },
    commentAnchorRect(from, to) {
      // Comments anchor main-doc text — a furniture story's geometry cannot
      // host one.
      if (story || !main.map?.valid) return null;
      const rects = main.map.selectionRects(from, to);
      const last = rects[rects.length - 1];
      if (!last) return null;
      const frame = opts.pageHost?.(framePage(main, last.page));
      if (!frame) return null;
      const scale = opts.scale?.() ?? 1;
      return {
        frame,
        left: (last.xPx + last.widthPx) * scale,
        top: last.yPx * scale,
        height: last.heightPx * scale,
      };
    },
    pasteAnchorRect(pos) {
      if (story || !main.map?.valid) return null;
      const rect = main.map.caretRect(pos);
      if (!rect) return null;
      const frame = opts.pageHost?.(framePage(main, rect.page));
      if (!frame) return null;
      const scale = opts.scale?.() ?? 1;
      return {
        frame,
        left: rect.xPx * scale,
        top: rect.yPx * scale,
        height: rect.heightPx * scale,
      };
    },
    /** Insert a docen slice payload (DOCEN_CLIP_MIME) into the ACTIVE story at
     *  the caret — the host's ribbon Paste routes here after reading the
     *  system clipboard. False when the payload did not parse. */
    insertSlicePayload(raw: string): boolean {
      return insertSlicePayload(raw);
    },
    focus(): void {
      ta.focus();
    },
    replaceOverlays(): void {
      // A re-render under an open crop layer orphans its geometry — the
      // drawing layer drops the mode and re-places its overlays (the drag
      // commits through Enter/click, never mid-transaction).
      draw.replaceOverlays();
      placeCaret();
    },
    /** Hand the host's fresh spell-check results to the overlay (the check
     *  itself runs in the host, debounced per transaction). */
    setSpellingIssues(issues: Array<{ from: number; to: number }>): void {
      spellingIssues = issues;
      placeSpelling();
    },
    destroy(): void {
      if (main.raf) cancelAnimationFrame(main.raf);
      blink?.cancel();
      main.editor.off("selectionUpdate", placeCaret);
      story?.editor.destroy();
      story = null;
      main.editor.destroy();
      stopDragAutoScroll();
      opts.host.removeEventListener("mousedown", takeFocus);
      document.removeEventListener("mousemove", onMouseMove);
      document.removeEventListener("mouseup", onMouseUp);
      draw.destroy();
      pooledPlace(selectionPool, [], {});
      pooledPlace(searchPool, [], {});
      pooledPlace(spellingPool, [], {});
      ta.remove();
      caret.remove();
    },
  };
}
