/**
 * Editor demo entry — registers Fluent components + theme, then mounts the
 * `<docen-document>` component (canvas route: chrome + LeaferJS pages + the
 * viewless editing bridge).
 *
 * Layout is a full-height flex column (declared in index.html).
 */
// Any named import from @docen/editor evaluates the module, which defines
// the <docen-document> custom element (the @customElement decorator).
import { applyTheme, registerComponents } from "@docen/editor";

// The content attribute accepts Tiptap JSON only (HTML input is the clipboard
// paste path, not the attribute).
const demoContent = JSON.stringify({
  type: "doc",
  content: [
    {
      type: "paragraph",
      attrs: { heading: "Heading1" },
      content: [{ type: "text", text: "Canvas Document" }],
    },
    {
      type: "paragraph",
      content: [
        {
          type: "text",
          text: "The docen-document component renders on canvas: click to place the caret, type to edit, and every transaction re-flows the pages.",
        },
      ],
    },
    {
      type: "paragraph",
      content: [
        {
          type: "text",
          text: "Use the ribbon to format, the file menu to open a .docx, and Ctrl+F to search.",
        },
      ],
    },
  ],
});

// registerComponents is async (it dynamically imports + defines the web
// components). Chain via .then — not top-level await, so this file stays
// tsc-clean under the repo tsconfig. `void` marks the floating promise.
void registerComponents().then(() => {
  applyTheme("light");

  const el = document.createElement("docen-document");
  el.className = "demo-doc";
  el.setAttribute("filename", "Demo.docx");
  el.setAttribute("content", demoContent);
  document.body.append(el);
});
