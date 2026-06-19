import { Emoji as BaseEmoji } from "./tiptap";

/**
 * Custom Emoji extension — caches the resolved glyph for DOCX export.
 *
 * Base `@tiptap/extension-emoji` stores only `name` (a shortcode); the glyph
 * itself is looked up from a dataset at render time. DOCX compile runs outside
 * the editor (no dataset available), so we cache the resolved character in
 * `attrs.emoji`, filled from the element's text content on `parseHTML`.
 *
 * DOCX has no emoji structure — a glyph is just a text run. Compile therefore
 * emits `attrs.emoji` (falling back to the `:name:` shortcode) as plain text;
 * resolve degrades DOCX text back to a text node (known lossy round-trip, like
 * codeBlock's dropped `language`). HTML rendering is inherited from the base.
 */

export const Emoji = BaseEmoji.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      // Resolved emoji glyph (e.g. "😀"), read from the rendered span's text.
      // Null when the base renders an <img> fallback (no text content) or when
      // the node was created in-editor without an HTML round-trip.
      emoji: {
        default: null,
        rendered: false,
        parseHTML: (el: HTMLElement) => el.textContent || null,
      },
    };
  },
});
