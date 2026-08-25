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

import { docxExtensions, type JSONContent } from "@docen/docx";
import { Editor } from "@docen/docx/core";
import type { FlowPage } from "@docen/layout";
import { UndoRedo } from "@tiptap/extensions";
import { joinBackward, joinForward, splitBlock } from "@tiptap/pm/commands";
import { TextSelection } from "@tiptap/pm/state";

import { CaretMap } from "./caret-map";

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

export interface EditBridgeOptions {
  /** A positioned host covering the canvas surface — the bridge's textarea
   *  overlays it and captures clicks to take focus. */
  host: HTMLElement;
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
}

export interface EditBridge {
  editor: Editor;
  /** Feed each render's flow result — rebuilds the pixel↔position map and
   *  re-places the caret against the fresh geometry. */
  updatePages(
    pages: readonly FlowPage[],
    flow: { contentLeftPx: number; contentTopPx: number },
  ): void;
  /** Scroll the page holding a doc position into view (null when unmappable). */
  scrollIntoView(pos: number): void;
  /** The page index a doc position renders on (null when unmappable). */
  pageOf(pos: number): number | null;
  /** The first doc position rendered on a page (null when unmappable). */
  firstPosOfPage(page: number): number | null;
  /** Move keyboard focus to the bridge's input surface (the editing focus —
   *  there is no DOM editor to focus). */
  focus(): void;
  destroy(): void;
}

export function mountEditBridge(opts: EditBridgeOptions): EditBridge {
  const editor = new Editor({
    element: null,
    extensions: [...(opts.extensions ?? docxExtensions), UndoRedo],
    content: opts.content as never,
  });
  // element:null skips mount(), and with it plugin installation — the state
  // comes up schema-only (EditorState.create with no plugins). Register the
  // extension manager's full, priority-sorted plugin list by hand: undo
  // history, keymaps, outline, search … all read their plugin's state.
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);

  // One relayout per frame regardless of keystroke bursts.
  let raf = 0;
  const schedule = (): void => {
    if (raf) return;
    raf = requestAnimationFrame(() => {
      raf = 0;
      opts.onDoc(editor.getJSON());
    });
  };
  editor.on("transaction", schedule);

  // The invisible input surface. Kept at 1px and positioned on click so the
  // IME candidate window anchors near the interaction point (true caret-point
  // anchoring lands with the caret milestone).
  const ta = document.createElement("textarea");
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
    width: "2px",
    background: "#000",
    pointerEvents: "none",
    zIndex: "5",
    display: "none",
  } satisfies Partial<CSSStyleDeclaration>);
  opts.host.append(caret);
  let blink: Animation | null = null;

  // The pixel↔position geometry, rebuilt from every render's flow result.
  // Between a transaction and its re-render it is one frame stale — caretRect
  // tolerates out-of-range positions by hiding until the fresh map lands.
  let map: CaretMap | null = null;
  let pageCount = 0;
  let lastCaretPos = -1;

  const restartBlink = (): void => {
    blink?.cancel();
    blink = caret.animate([{ opacity: 1 }, { opacity: 0 }], {
      duration: 1000,
      iterations: Infinity,
      direction: "alternate",
      easing: "steps(1,end)",
    });
  };

  /** The selection highlight — one translucent div per crossed line, rebuilt
   *  on every placement (selections span few lines; rebuild beats diffing). */
  const selectionLayer: HTMLDivElement[] = [];
  const placeSelection = (): void => {
    for (const el of selectionLayer) el.remove();
    selectionLayer.length = 0;
    const { from, to } = editor.state.selection;
    if (from === to || !map?.valid) return;
    for (const r of map.selectionRects(from, to)) {
      const frame = opts.pageHost?.(r.page);
      if (!frame) continue;
      const el = document.createElement("div");
      Object.assign(el.style, {
        position: "absolute",
        background: "rgba(0,120,215,.25)",
        pointerEvents: "none",
        zIndex: "4",
        left: `${r.xPx}px`,
        top: `${r.yPx}px`,
        width: `${r.widthPx}px`,
        height: `${r.heightPx}px`,
      } satisfies Partial<CSSStyleDeclaration>);
      frame.append(el);
      selectionLayer.push(el);
    }
  };

  const placeCaret = (): void => {
    placeSelection();
    if (!map?.valid) {
      caret.style.display = "none";
      return;
    }
    const { from, to } = editor.state.selection;
    if (from !== to) {
      // A selection replaces the caret (Word hides it too).
      caret.style.display = "none";
      return;
    }
    const rect = map.caretRect(from);
    const frame = rect ? (opts.pageHost?.(rect.page) ?? null) : null;
    if (!rect || !frame) {
      caret.style.display = "none";
      return;
    }
    if (frame !== caret.parentElement) frame.append(caret);
    caret.style.display = "block";
    caret.style.left = `${rect.xPx}px`;
    caret.style.top = `${rect.yPx}px`;
    caret.style.height = `${rect.heightPx}px`;
    // Keep the textarea anchored at the caret so the IME candidate window
    // opens at the typing point.
    const hostRect = opts.host.getBoundingClientRect();
    const frameRect = frame.getBoundingClientRect();
    ta.style.left = `${frameRect.left - hostRect.left + rect.xPx}px`;
    ta.style.top = `${frameRect.top - hostRect.top + rect.yPx}px`;
    if (from !== lastCaretPos) {
      lastCaretPos = from;
      restartBlink();
    }
  };
  editor.on("selectionUpdate", placeCaret);

  /** A viewport point → the hit page and its page-local coordinates. */
  const hitPage = (
    clientX: number,
    clientY: number,
  ): { page: number; lx: number; ly: number } | null => {
    if (!map?.valid) return null;
    const hostRect = opts.host.getBoundingClientRect();
    const x = clientX - hostRect.left;
    const y = clientY - hostRect.top;
    for (let p = 0; p < pageCount; p++) {
      const frame = opts.pageHost?.(p);
      if (!frame) continue;
      const r = frame.getBoundingClientRect();
      const lx = x - (r.left - hostRect.left);
      const ly = y - (r.top - hostRect.top);
      if (lx >= 0 && ly >= 0 && lx < r.width && ly < r.height) return { page: p, lx, ly };
    }
    return null;
  };

  const setSel = (pos: number, anchor?: number): void => {
    editor.commands.command(({ state, dispatch }) => {
      // The cast bridges the dual PM d.ts identity (same runtime
      // instance — see the module's command casts). create takes
      // (anchor, head) — the anchor leads.
      dispatch?.(
        state.tr.setSelection(TextSelection.create(state.doc, anchor ?? pos, pos) as never),
      );
      return true;
    });
  };

  const posAtClient = (clientX: number, clientY: number): number | null => {
    const hit = hitPage(clientX, clientY);
    return hit ? map!.posAtPoint(hit.page, hit.lx, hit.ly) : null;
  };

  // Mouse selection: mousedown anchors, moves extend, mouseup settles. The
  // 3px threshold keeps a plain click from flashing a degenerate selection.
  let dragAnchor: number | null = null;
  let dragMoved = false;
  let dragStart: { x: number; y: number } | null = null;

  const onMouseMove = (event: MouseEvent): void => {
    if (dragAnchor == null) return;
    if (
      !dragMoved &&
      dragStart &&
      Math.hypot(event.clientX - dragStart.x, event.clientY - dragStart.y) < 3
    ) {
      return;
    }
    dragMoved = true;
    const head = posAtClient(event.clientX, event.clientY);
    if (head != null) setSel(head, dragAnchor);
  };
  const onMouseUp = (): void => {
    dragAnchor = null;
    dragMoved = false;
    dragStart = null;
  };
  opts.host.addEventListener("mousemove", onMouseMove);
  document.addEventListener("mouseup", onMouseUp);

  const takeFocus = (event: MouseEvent): void => {
    // preventDefault keeps the click from blurring on mousedown; the caret
    // placement below is the real focus move.
    event.preventDefault();
    const pos = posAtClient(event.clientX, event.clientY);
    if (pos != null) {
      setSel(pos);
      dragAnchor = pos;
      dragStart = { x: event.clientX, y: event.clientY };
      dragMoved = false;
    }
    ta.focus();
    ta.value = "";
  };
  opts.host.addEventListener("mousedown", takeFocus);

  let composing = false;

  const insertText = (text: string): void => {
    editor.commands.command(({ state, dispatch }) => {
      const { from, to } = state.selection;
      dispatch?.(state.tr.insertText(text, from, to));
      return true;
    });
  };

  const backspace = (): void => {
    editor.commands.command(({ state, dispatch }) => {
      const { selection } = state;
      if (!selection.empty) {
        dispatch?.(state.tr.deleteSelection());
        return true;
      }
      const $from = selection.$from;
      const text = $from.parent.textContent;
      if ($from.parentOffset > 0 && text) {
        const cut = lastGraphemeUnits(text.slice(0, $from.parentOffset));
        if (cut > 0) {
          dispatch?.(state.tr.delete($from.pos - cut, $from.pos));
          return true;
        }
      }
      // Same runtime PM instance (single .pnpm dir); the cast bridges the
      // dual d.ts identity between this package's @tiptap/pm and the engine's.
      return joinBackward(state as never, dispatch);
    });
  };

  const deleteForward = (): void => {
    editor.commands.command(({ state, dispatch }) => {
      const { selection } = state;
      if (!selection.empty) {
        dispatch?.(state.tr.deleteSelection());
        return true;
      }
      const $from = selection.$from;
      const text = $from.parent.textContent;
      const offset = $from.parentOffset;
      if (offset < text.length) {
        const cut = firstGraphemeUnits(text.slice(offset));
        if (cut > 0) {
          dispatch?.(state.tr.delete($from.pos, $from.pos + cut));
          return true;
        }
      }
      return joinForward(state as never, dispatch);
    });
  };

  const onBeforeInput = (event: InputEvent): void => {
    event.preventDefault();
    if (composing) return;
    switch (event.inputType) {
      case "insertText":
        if (event.data) insertText(event.data);
        break;
      // Chrome reports textarea Enter as insertLineBreak; keep both mapped to
      // a paragraph split (Shift+Enter variants included for now).
      case "insertParagraph":
      case "insertLineBreak":
        editor.commands.command(({ state, dispatch }) => splitBlock(state as never, dispatch));
        break;
      case "deleteContentBackward":
      case "deleteWordBackward":
        backspace();
        break;
      case "deleteContentForward":
      case "deleteWordForward":
        deleteForward();
        break;
      default:
        break;
    }
    ta.value = "";
  };

  /** One step horizontally from a position — a grapheme inside the block
   *  (atoms step a whole node), the nearest text position past its edge. */
  const hStep = (state: Editor["state"], pos: number, dir: -1 | 1): number | null => {
    const $from = state.doc.resolve(pos);
    const offset = $from.parentOffset;
    const size = $from.parent.content.size;
    if (dir < 0 && offset > 0) {
      // textBetween maps the content offset to text units, skipping atoms.
      const cut = lastGraphemeUnits($from.parent.textBetween(0, offset)) || 1;
      return pos - cut;
    }
    if (dir > 0 && offset < size) {
      const cut = firstGraphemeUnits($from.parent.textBetween(offset, size)) || 1;
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
    const edges = map?.valid ? map.lineEdges(pos) : null;
    if (edges) return toEnd ? edges.end : edges.home;
    return toEnd ? $from.end() : $from.start();
  };

  /** One line up/down at the goal column (null at the paragraph's edge). */
  const vStep = (pos: number, dir: -1 | 1): number | null => {
    return map?.valid ? map.posVertical(pos, dir) : null;
  };

  /** Move the caret to a target — or extend the selection to it (the anchor
   *  holds, the head moves). */
  const apply = (target: number | null, extend: boolean): void => {
    if (target == null) return;
    if (extend) {
      setSel(target, editor.state.selection.anchor);
    } else {
      setSel(target);
    }
  };

  const onKeyDown = (event: KeyboardEvent): void => {
    if (event.ctrlKey || event.metaKey) {
      const key = event.key.toLowerCase();
      if (key === "z") {
        event.preventDefault();
        if (event.shiftKey) editor.commands.redo();
        else editor.commands.undo();
      } else if (key === "y") {
        event.preventDefault();
        editor.commands.redo();
      }
      return;
    }
    const extend = event.shiftKey;
    const head = () => editor.state.selection.head;
    switch (event.key) {
      case "ArrowLeft":
        event.preventDefault();
        apply(hStep(editor.state, head(), -1), extend);
        break;
      case "ArrowRight":
        event.preventDefault();
        apply(hStep(editor.state, head(), 1), extend);
        break;
      case "ArrowUp":
        event.preventDefault();
        apply(vStep(head(), -1), extend);
        break;
      case "ArrowDown":
        event.preventDefault();
        apply(vStep(head(), 1), extend);
        break;
      case "Home":
        event.preventDefault();
        apply(edgeTarget(editor.state, head(), false), extend);
        break;
      case "End":
        event.preventDefault();
        apply(edgeTarget(editor.state, head(), true), extend);
        break;
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

  const onPaste = (event: ClipboardEvent): void => {
    event.preventDefault();
    const text = event.clipboardData?.getData("text/plain");
    if (text) insertText(text);
  };

  const selectionText = (): string | null => {
    const { from, to } = editor.state.selection;
    return from === to ? null : editor.state.doc.textBetween(from, to, "\n");
  };

  const onCopy = (event: ClipboardEvent): void => {
    const text = selectionText();
    if (text == null) return;
    event.preventDefault();
    event.clipboardData?.setData("text/plain", text);
  };

  const onCut = (event: ClipboardEvent): void => {
    const text = selectionText();
    if (text == null) return;
    event.preventDefault();
    event.clipboardData?.setData("text/plain", text);
    editor.commands.command(({ state, dispatch }) => {
      dispatch?.(state.tr.deleteSelection());
      return true;
    });
  };

  ta.addEventListener("beforeinput", onBeforeInput);
  ta.addEventListener("keydown", onKeyDown);
  ta.addEventListener("compositionstart", onCompositionStart);
  ta.addEventListener("compositionend", onCompositionEnd);
  ta.addEventListener("paste", onPaste);
  ta.addEventListener("copy", onCopy);
  ta.addEventListener("cut", onCut);
  opts.host.append(ta);

  return {
    editor,
    updatePages(pages, flow): void {
      map = new CaretMap(pages, editor.state.doc, flow);
      pageCount = pages.length;
      placeCaret();
    },
    scrollIntoView(pos): void {
      const rect = map?.valid ? map.caretRect(pos) : null;
      if (!rect) return;
      opts.pageHost?.(rect.page)?.scrollIntoView({ block: "start", behavior: "auto" });
    },
    /** The page index a doc position renders on (null when unmappable). */
    pageOf(pos): number | null {
      return map?.valid ? (map.caretRect(pos)?.page ?? null) : null;
    },
    /** The first doc position rendered on a page (null when unmappable). */
    firstPosOfPage(page: number): number | null {
      return map?.valid ? map.firstPosOfPage(page) : null;
    },
    focus(): void {
      ta.focus();
    },
    destroy(): void {
      if (raf) cancelAnimationFrame(raf);
      blink?.cancel();
      editor.off("selectionUpdate", placeCaret);
      editor.destroy();
      opts.host.removeEventListener("mousedown", takeFocus);
      opts.host.removeEventListener("mousemove", onMouseMove);
      document.removeEventListener("mouseup", onMouseUp);
      for (const el of selectionLayer) el.remove();
      ta.remove();
      caret.remove();
    },
  };
}
