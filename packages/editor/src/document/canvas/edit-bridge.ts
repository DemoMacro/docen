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
import { UndoRedo } from "@tiptap/extensions";
import { joinBackward, joinForward, splitBlock } from "@tiptap/pm/commands";
import { history } from "@tiptap/pm/history";

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
}

export interface EditBridge {
  editor: Editor;
  destroy(): void;
}

export function mountEditBridge(opts: EditBridgeOptions): EditBridge {
  const editor = new Editor({
    element: null,
    extensions: [...docxExtensions, UndoRedo],
    content: opts.content as never,
  });
  // element:null skips mount(), and with it plugin installation — the state
  // comes up schema-only (EditorState.create with no plugins), so the UndoRedo
  // extension's history plugin never lands. Register it by hand; its commands
  // (undo/redo) read this plugin's state.
  editor.registerPlugin(history());

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

  const takeFocus = (event: MouseEvent): void => {
    // preventDefault keeps the click from blurring on mousedown; the caret
    // placement from pixel coordinates is the caret milestone's job.
    event.preventDefault();
    ta.style.left = `${event.offsetX}px`;
    ta.style.top = `${event.offsetY}px`;
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

  ta.addEventListener("beforeinput", onBeforeInput);
  ta.addEventListener("keydown", onKeyDown);
  ta.addEventListener("compositionstart", onCompositionStart);
  ta.addEventListener("compositionend", onCompositionEnd);
  ta.addEventListener("paste", onPaste);
  opts.host.append(ta);

  return {
    editor,
    destroy(): void {
      if (raf) cancelAnimationFrame(raf);
      editor.destroy();
      opts.host.removeEventListener("mousedown", takeFocus);
      ta.remove();
    },
  };
}
