import type { EditorView } from "@tiptap/pm/view";

/**
 * Scroll helpers shared by the editor and the docx engine's link-click handler,
 * so every "jump then scroll" path (outline heading, search result, find, TOC
 * Ctrl+Click, post-reflow caret follow) scrolls the SAME way — to the top of the
 * viewport, Office-style. ProseMirror's default `tr.scrollIntoView()` parks the
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

/** Resolve the `.docen-page` block holding `pos` (C-route pagination), or null
 *  when the position isn't inside a page wrapper (non-C-route). At doc start/end
 *  the caret maps to the editor root rather than a page, so fall back to the
 *  page whose vertical band contains the caret (last page at/after doc end). */
function pageElementAt(view: EditorView, pos: number): HTMLElement | null {
  const { node } = view.domAtPos(pos);
  const el = node.nodeType === Node.ELEMENT_NODE ? (node as HTMLElement) : node.parentElement;
  const direct = el?.closest<HTMLElement>(".docen-page");
  if (direct) return direct;
  const pages = view.dom.querySelectorAll<HTMLElement>(".docen-page");
  if (pages.length === 0) return null;
  const caretTop = view.coordsAtPos(pos).top;
  for (const p of pages) {
    if (caretTop <= p.getBoundingClientRect().bottom) return p;
  }
  return pages[pages.length - 1];
}

/** Scroll the caret to the TOP of the viewport when it has left the visible area
 *  (Office-style page follow). No-op while the caret stays in view, so normal
 *  typing doesn't fight the user's scroll. Replaces PM's default scrollIntoView.
 *
 * Follows the caret's PAGE, not the caret itself: a caret near a page's bottom
 * edge (e.g. after a select-all delete leaves the caret at the end of an
 * otherwise-empty page) would otherwise park that page's lower half in view,
 * which reads as "scrolled to the bottom." Anchoring on the page top keeps the
 * page's start in view. Falls back to the caret position outside C-route. */
export function scrollCaretToTop(view: EditorView): void {
  const scroller = scrollContainerOf(view);
  if (!scroller) return;
  const margin = 64;
  const head = view.state.selection.head;
  const caretTop = view.coordsAtPos(head).top;
  const rect = scroller.getBoundingClientRect();
  // Only scroll when the caret has left the viewport — normal in-view typing
  // stays put so it doesn't fight the user's scroll. Replaces PM's default
  // scrollIntoView, which parks the caret at the bottom edge.
  if (caretTop < rect.top + margin || caretTop > rect.bottom - margin) {
    // Anchor on the caret's PAGE top, not the caret itself: a caret near a
    // page's bottom edge (e.g. after a select-all delete leaves the caret at the
    // end of an otherwise-empty page) would otherwise park that page's lower
    // half in view, reading as "scrolled to the bottom." Falls back to the
    // caret position outside C-route.
    const pageEl = pageElementAt(view, head);
    const anchorTop = pageEl ? pageEl.getBoundingClientRect().top : caretTop;
    scroller.scrollTop += anchorTop - rect.top - margin;
  }
}
