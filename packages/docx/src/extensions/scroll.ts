import type { EditorView } from "@tiptap/pm/view";

/**
 * Scroll helpers shared by the editor and the docx engine's link-click handler,
 * so every "jump then scroll" path (outline heading, search result, find, TOC
 * Ctrl+Click, post-reflow caret follow) scrolls the SAME way — to the top of the
 * viewport, Word-style. ProseMirror's default `tr.scrollIntoView()` parks the
 * caret at the bottom edge, which reads wrong for a page/heading/TOC jump.
 */

/** Nearest scrollable ancestor of the editor surface (e.g. the docen-canvas). */
export function scrollContainerOf(view: EditorView): HTMLElement | null {
  let el: HTMLElement | null = view.dom.parentElement;
  while (el) {
    if (el.clientHeight > 0 && el.scrollHeight > el.clientHeight) {
      const overflowY = getComputedStyle(el).overflowY;
      if (overflowY === "auto" || overflowY === "scroll") return el;
    }
    el = el.parentElement;
  }
  return null;
}

/** Scroll the caret to the TOP of the viewport when it has left the visible area
 *  (Word-style page follow). No-op while the caret stays in view, so normal
 *  typing doesn't fight the user's scroll. Replaces PM's default scrollIntoView. */
export function scrollCaretToTop(view: EditorView): void {
  const scroller = scrollContainerOf(view);
  if (!scroller) return;
  const margin = 64;
  const caretTop = view.coordsAtPos(view.state.selection.head).top;
  const rect = scroller.getBoundingClientRect();
  if (caretTop < rect.top + margin || caretTop > rect.bottom - margin) {
    scroller.scrollTop += caretTop - rect.top - margin;
  }
}
