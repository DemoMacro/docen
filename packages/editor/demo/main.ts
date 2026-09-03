/**
 * Editor demo entry — registers Fluent components + theme, then mounts the
 * `<docen-document>` component (canvas route: chrome + LeaferJS pages + the
 * viewless editing bridge).
 *
 * Layout is a full-height flex column (declared in index.html).
 */
// Any named import from @docen/editor evaluates the module, which defines
// the <docen-document> custom element (the @customElement decorator).
import { normalizeDocument, type JSONContent } from "@docen/docx";
import { applyTheme, registerComponents } from "@docen/editor";

// A text run with optional marks — the demo document's only repetition.
const t = (text: string, marks?: JSONContent["marks"]): JSONContent => ({
  type: "text",
  text,
  ...(marks ? { marks } : {}),
});

// The sample pictures ship from office-open's demo corpus, served statically
// out of this demo directory by the dev server. The canvas pipeline consumes
// data URLs (the same contract prepareDocument gives export), so the demo
// inlines them before mounting.
const dog = { src: "/demo/images/dog.png", width: 166, height: 150 };
const cat = { src: "/demo/images/cat.jpg", width: 208, height: 139 };

async function toDataUrl(url: string): Promise<string> {
  const blob = await (await fetch(url)).blob();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (): void => resolve(reader.result as string);
    reader.onerror = (): void => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

// The showcase document: two sections — section 1 (title, rich text, links,
// lists, a table, pictures) ends at the section break; section 2 flows in two
// columns with a separator. Built lazily so the picture srcs are
// already-inlined data URLs.
const demoDocument = (): JSONContent => ({
  type: "doc",
  content: [
    // ── Section 1 ──
    {
      type: "paragraph",
      attrs: { heading: "Title", alignment: "center" },
      content: [t("docen Feature Showcase")],
    },
    {
      type: "paragraph",
      attrs: { alignment: "center" },
      content: [
        t("Canvas typesetting · Word-faithful layout · Fully editable", [
          { type: "textStyle", attrs: { color: "595959", characterSpacing: 60 } },
        ]),
      ],
    },
    {
      type: "paragraph",
      attrs: { heading: "Heading1" },
      content: [t("Character Formatting")],
    },
    {
      type: "paragraph",
      content: [
        t("This paragraph exercises character-level formatting: "),
        t("bold", [{ type: "bold" }]),
        t(", "),
        t("italic", [{ type: "italic" }]),
        t(", "),
        t("underline", [{ type: "underline" }]),
        t(", "),
        t("strikethrough", [{ type: "strike" }]),
        t(", "),
        t("highlighted text", [{ type: "highlight", attrs: { color: "yellow" } }]),
        t(", "),
        t("colored text", [{ type: "textStyle", attrs: { color: "C00000" } }]),
        t(", plus superscript m"),
        t("2", [{ type: "superscript" }]),
        t(" and subscript H"),
        t("2", [{ type: "subscript" }]),
        t("O. Edit anything — the canvas re-flows live."),
      ],
    },
    {
      type: "paragraph",
      content: [
        t("Links and character styles: visit the "),
        t("project homepage", [
          { type: "link", attrs: { href: "https://github.com/DemoMacro/docen", target: "_blank" } },
          { type: "textStyle", attrs: { style: "Hyperlink" } },
        ]),
        t(" (the link blue comes from the Hyperlink character style, matching Word)."),
      ],
    },
    {
      type: "paragraph",
      attrs: { heading: "Heading2" },
      content: [t("Lists")],
    },
    {
      type: "paragraph",
      attrs: { bullet: { level: 0 } },
      content: [t("A first-level bullet item")],
    },
    {
      type: "paragraph",
      attrs: { bullet: { level: 0 } },
      content: [t("Pagination matches Word's layout engine")],
    },
    {
      type: "paragraph",
      attrs: { bullet: { level: 1 } },
      content: [t("A nested second-level item")],
    },
    {
      type: "paragraph",
      attrs: { heading: "Heading2" },
      content: [t("Tables")],
    },
    {
      type: "table",
      attrs: {
        columnWidths: [2400, 4200, 2400],
        // Word's TableGrid look: 0.5pt black grid (size in 1/8 pt).
        borders: {
          top: { style: "single", size: 4, color: "000000" },
          bottom: { style: "single", size: 4, color: "000000" },
          left: { style: "single", size: 4, color: "000000" },
          right: { style: "single", size: 4, color: "000000" },
          insideHorizontal: { style: "single", size: 4, color: "000000" },
          insideVertical: { style: "single", size: 4, color: "000000" },
        },
      },
      content: [
        {
          type: "tableRow",
          attrs: { tableHeader: true },
          content: ["Feature", "Description", "Status"].map((label) => ({
            type: "tableCell",
            attrs: { shading: { fill: "DEEAF6" }, verticalAlign: "center" },
            content: [
              {
                type: "paragraph",
                content: [t(label, [{ type: "bold" }])],
              },
            ],
          })),
        },
        {
          type: "tableRow",
          content: [
            {
              type: "tableCell",
              attrs: { verticalAlign: "center" },
              content: [{ type: "paragraph", content: [t("Mid-row split")] }],
            },
            {
              type: "tableCell",
              content: [
                {
                  type: "paragraph",
                  content: [
                    t(
                      "A row taller than the page splits at a line boundary; the header row repeats on every page (tblHeader).",
                    ),
                  ],
                },
              ],
            },
            {
              type: "tableCell",
              attrs: { verticalAlign: "center" },
              content: [{ type: "paragraph", content: [t("Shipped")] }],
            },
          ],
        },
        {
          type: "tableRow",
          content: [
            {
              type: "tableCell",
              attrs: { columnSpan: 2, verticalAlign: "center" },
              content: [{ type: "paragraph", content: [t("A spanned cell (columnSpan = 2)")] }],
            },
            {
              type: "tableCell",
              content: [{ type: "paragraph", content: [t("Shipped")] }],
            },
          ],
        },
      ],
    },
    {
      type: "paragraph",
      content: [],
    },
    {
      type: "paragraph",
      attrs: { alignment: "center" },
      content: [
        {
          type: "image",
          attrs: { src: dog.src, width: dog.width, height: dog.height, alt: "A dog" },
        },
        t("  "),
        {
          type: "image",
          attrs: { src: cat.src, width: cat.width, height: cat.height, alt: "A cat" },
        },
      ],
    },
    // Section break — an explicit next-page section break; the next section
    // starts on a fresh page.
    {
      type: "paragraph",
      attrs: { sectionProperties: {} },
      content: [],
    },

    // ── Section 2: two columns with a separator. The FINAL section's settings
    // ride the doc attrs (body-level sectPr) — a sectionProperties on the last
    // paragraph would close an extra empty section (a blank trailing page).
    // The body text is long enough to fill the left column and spill into the
    // right one, so both columns carry content like a real newsletter. ──
    {
      type: "paragraph",
      attrs: { heading: "Heading1" },
      content: [t("Multi-Column Layout")],
    },
    {
      type: "paragraph",
      content: [
        t(
          "This section flows in two balanced columns, the way Word lays out newsletters, brochures and minutes. Columns are part of the OOXML section properties (w:cols) and share the same section state machine as the page setup: the engine fills the left column first, then continues into the right one, and a new page starts only when BOTH columns are full — exactly the order Word's layout engine walks.",
        ),
      ],
    },
    {
      type: "paragraph",
      content: [
        t(
          "The separator line between the columns comes from the section's separate flag, and the gap between them is measured in twentieths of a point just like Word's spacing-to setting. Manual column widths and column breaks project from the same section model, so a document saved here opens in Word with the identical column geometry — and a document Word authored parses back into the same section state.",
        ),
      ],
    },
    {
      type: "paragraph",
      content: [
        t(
          "Editing behaves the same inside columns as anywhere else: put the caret in any column and type — wrapping, column flow and repagination happen live. Watch the boundary between the columns move as text grows: sentences migrate from the bottom of the left column to the top of the right one without any explicit instruction, because the layout engine — not the document — owns the flow.",
        ),
      ],
    },
    {
      type: "paragraph",
      content: [
        t(
          "Each section of a document carries its own page setup: size, orientation, margins, headers and footers, and these column settings. The section break above is a next-page break, so section two starts on a fresh page and keeps its two-column grid while section one stays single-column — the two models coexist in one file, one per section.",
        ),
      ],
    },
    {
      type: "paragraph",
      content: [
        t(
          "Under the canvas the same story holds: every keystroke is one transaction, and every transaction re-runs the pipeline — compile to the document model, project to layout geometry, paginate across columns and pages, then paint. That is the core loop of the canvas route, and columns are simply one more constraint the paginator respects while filling pages.",
        ),
      ],
    },
    {
      type: "paragraph",
      content: [
        t(
          "Try it: select a word here and make it bold, or drag the pictures above into the columns. Formatting, tables, images and links all flow through the same two-column text stream — nothing about a column changes how content behaves, only where the paginator is allowed to place it.",
        ),
      ],
    },
  ],
  attrs: { sectionProperties: { columns: { count: 2, space: 360, separate: true } } },
});

// The UI language follows the browser (the library itself has no navigator
// fallback — resolveLang reads <html lang> and defaults to "en"). Chrome stays
// localized and the Cell Size boxes report in the locale's unit system
// (厘米 for zh, inches otherwise); the DOCUMENT content stays English either
// way, like an English document open in localized Word.
document.documentElement.lang = navigator.language || "en";

// registerComponents is async (it dynamically imports + defines the web
// components). Chain via .then — not top-level await, so this file stays
// tsc-clean under the repo tsconfig. `void` marks the floating promise.
void registerComponents().then(async () => {
  applyTheme("light");

  [dog.src, cat.src] = await Promise.all([toDataUrl(dog.src), toDataUrl(cat.src)]);

  const el = document.createElement("docen-document");
  el.className = "demo-doc";
  el.setAttribute("filename", "Demo.docx");
  // Formatting marks (¶, →, ·) on by default — Word's Show/Hide ¶.
  el.setAttribute("show-marks", "");
  // A hand-built doc carries no style library — normalizeDocument fills the
  // document-level defaults (doc attrs win over them), same as setJSON does.
  el.setAttribute("content", JSON.stringify(normalizeDocument(demoDocument())));
  document.body.append(el);
});
