import { Extension } from "@docen/docx/core";

/**
 * Centralized Tiptap keymap for docen EDITING shortcuts (MS Office-aligned).
 *
 * Each shortcut dispatches its mapped command directly on `editor.commands`
 * — every command name IS a native Tiptap command (see ./commands), so a
 * shortcut and its ribbon button share ONE definition with no bridge. Add
 * entries to {@link KEYBOARD_SHORTCUTS} to bind more.
 *
 * ── Scope — two layers, by necessity ──
 *
 * This extension binds only EDITING-layer shortcuts: keystrokes that mutate the
 * ProseMirror document via a command. Chrome-layer shortcuts operate on the UI
 * shell (not the document) and live as a host `keydown` listener in index.ts
 * (`#onZoomKey`), because they read host state the editor cannot reach and must
 * be ignored inside inputs/comboboxes:
 *   • Ctrl/Cmd + = / +   → zoom in        (canvas CSS zoom)
 *   • Ctrl/Cmd + - / _   → zoom out
 *   • Ctrl/Cmd + 0       → zoom reset 100%
 *   • Ctrl/Cmd + F       → open Find
 *   • Ctrl/Cmd + H       → open Find & Replace
 *
 * Per-extension defaults (bold=Mod-B, italic=Mod-I, HardBreak Mod/Shift-Enter,
 * the ListKeymap) stay with their owning extensions by Tiptap convention.
 */
const KEYBOARD_SHORTCUTS: Readonly<Record<string, string>> = {
  // Ctrl+Enter → page break, Ctrl+Shift+Enter → column break (Word). Shift+Enter
  // (soft line break) stays on @tiptap/extension-hard-break's default. High
  // priority: HardBreak also maps Mod-Enter (to a soft break), and these must win.
  "Mod-Enter": "page-break",
  "Mod-Shift-Enter": "column-break",
};

export const DocenKeymap = Extension.create({
  name: "docenKeymap",
  priority: 1000,
  addKeyboardShortcuts() {
    return Object.fromEntries(
      Object.entries(KEYBOARD_SHORTCUTS).map(([key, event]) => [
        key,
        () => {
          const cmd = (
            this.editor.commands as unknown as Record<string, (() => boolean) | undefined>
          )[event];
          return typeof cmd === "function" ? cmd() : false;
        },
      ]),
    );
  },
});
