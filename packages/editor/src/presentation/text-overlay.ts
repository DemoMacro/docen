const textMirrors = new WeakMap<HTMLTextAreaElement, HTMLDivElement>();

/** Measure a textarea's text stack without the control's row/intrinsic-height
 *  feedback. Reading a textarea's scrollHeight after it has had a large
 *  height can change between keystrokes; the mirror uses the same text styles
 *  and content width, so vertical anchoring stays stable while deleting. */
export function measureTextAreaContent(editor: HTMLTextAreaElement): number {
  let mirror = textMirrors.get(editor);
  if (!mirror) {
    mirror = document.createElement("div");
    mirror.setAttribute("aria-hidden", "true");
    Object.assign(mirror.style, {
      position: "fixed",
      left: "-999999px",
      top: "0",
      visibility: "hidden",
      pointerEvents: "none",
      boxSizing: "content-box",
      whiteSpace: "pre-wrap",
    });
    editor.before(mirror);
    textMirrors.set(editor, mirror);
  }

  const style = getComputedStyle(editor);
  const rect = editor.getBoundingClientRect();
  const horizontalPadding =
    Number.parseFloat(style.paddingLeft) + Number.parseFloat(style.paddingRight);
  Object.assign(mirror.style, {
    width: `${Math.max(0, rect.width - horizontalPadding)}px`,
    fontFamily: style.fontFamily,
    fontSize: style.fontSize,
    fontWeight: style.fontWeight,
    fontStyle: style.fontStyle,
    lineHeight: style.lineHeight,
    letterSpacing: style.letterSpacing,
    textAlign: style.textAlign,
    textTransform: style.textTransform,
    textIndent: style.textIndent,
    overflowWrap: style.overflowWrap,
    wordBreak: style.wordBreak,
    tabSize: style.tabSize,
    direction: style.direction,
  });
  mirror.textContent = editor.value;
  return mirror.getBoundingClientRect().height;
}

export function disposeTextAreaMirror(editor: HTMLTextAreaElement | null): void {
  const mirror = editor && textMirrors.get(editor);
  mirror?.remove();
  if (editor) textMirrors.delete(editor);
}
