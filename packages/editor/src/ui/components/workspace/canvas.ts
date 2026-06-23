const template = document.createElement("template");
template.innerHTML = `
  <style>
    :host {
      display: block;
      flex: 1 1 auto;
      min-width: 0;
      min-height: 0;
      overflow: auto;
      background: var(--docen-color-canvas, #ffffff);
      padding: var(--docen-page-gap, 24px);
    }
    /* Hosts the editor wrapper (.docen-pages), which in turn hosts the
       Tiptap .ProseMirror. The ProseMirror renders one .docen-page NODE per
       page — each page carries its own paper geometry (fixed width/height,
       margin, crop marks) via the CSS variables set on :host below. This
       slotted wrapper is just the scroll surface's child — it is NOT the
       paper itself. (C-route pagination — see CLAUDE.md.) */
    ::slotted(.docen-pages) {
      display: block;
      margin-inline: auto;
      width: fit-content;
    }
    /* Print: drop the scroll chrome so only the pages print. The page nodes'
       own @media print rules drop their shadow/gap. */
    @media print {
      :host { overflow: visible; padding: 0; background: #fff; }
    }
  </style>
  <slot></slot>`;

/**
 * `<docen-canvas>` — the editor workspace surface: a scrolling container. The
 * editor packages slot their engine wrapper (`<div class="docen-pages">`)
 * into the default slot; that wrapper hosts the Tiptap `.ProseMirror`, which
 * renders one `.docen-page` node per page.
 *
 * Page geometry is driven by attributes (raw lengths, mm if unit-less) and
 * exposed as CSS variables consumed by the page-node style:
 *  - `page-width` / `page-height` → `--docen-page-width` / `--docen-page-min-height`.
 *  - `margin` → `--docen-page-margin` (content padding).
 *  - `orientation` — "portrait" (default) | "landscape"; swaps width/height in
 *    the variables so page nodes flip automatically.
 *
 * Paper-size and margin *presets* (A4, Letter, Normal, Narrow, …) live in the
 * host document, which resolves them to mm and sets these attributes.
 */
class DocenCanvas extends HTMLElement {
  static get observedAttributes(): string[] {
    return ["page-width", "page-height", "margin", "orientation", "zoom"];
  }

  attributeChangedCallback(): void {
    this.#applyPage();
  }

  connectedCallback(): void {
    if (!this.shadowRoot) {
      this.attachShadow({ mode: "open" }).append(template.content.cloneNode(true));
    }
    this.#applyPage();
  }

  /** Apply the page geometry from attributes to :host CSS variables. Values
   *  are raw lengths (mm when unit-less); the page-node style consumes them.
   *  `orientation` swaps width/height HERE (in the variables), so the page
   *  nodes flip with no extra CSS rule. */
  #applyPage(): void {
    if (!this.isConnected) return;
    const norm = (v: string | null, fallback: string): string =>
      v == null ? fallback : /^\d+(\.\d+)?$/.test(v.trim()) ? `${v}mm` : v;
    const w = norm(this.getAttribute("page-width"), "210mm");
    const h = norm(this.getAttribute("page-height"), "297mm");
    const landscape = this.getAttribute("orientation") === "landscape";
    this.style.setProperty("--docen-page-width", landscape ? h : w);
    this.style.setProperty("--docen-page-min-height", landscape ? w : h);
    // Normalize each side to a length (mm when unit-less) — same as width/height.
    // `margin` arrives as a 1–4 value shorthand; without units `padding` would
    // silently fall back to 0 (a raw "31.70 25.40 ..." is not a valid length).
    const margin = this.getAttribute("margin");
    if (margin != null) {
      const sides = margin
        .trim()
        .split(/\s+/)
        .map((v) => (/^\d+(\.\d+)?$/.test(v) ? `${v}mm` : v));
      this.style.setProperty("--docen-page-margin", sides.join(" "));
    }
    // Zoom (percent) — CSS `zoom` rescales the pages and reflows the scroll
    // surface (Chromium-native). "150" → 1.5; absent clears it (100%).
    const zoom = this.getAttribute("zoom");
    this.style.zoom = zoom ? String(Math.max(10, parseFloat(zoom)) / 100) : "";
  }
}

customElements.define("docen-canvas", DocenCanvas);

export default DocenCanvas;
