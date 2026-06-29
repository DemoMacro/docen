import { Extension } from "@docen/docx/core";
import type { Node as PmNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import type { EditorView } from "@tiptap/pm/view";
import { imageMeta } from "image-meta";

import { sectionContentDims } from "./page-plugin";

const key = new PluginKey("docenImageCap");

/** Decode a `data:image/...;base64,...` URL to bytes; null if not a data URL.
 *  Mirrors @docen/docx prepare.ts — the editor resolves docx via dist at runtime
 *  (no converter import), so the tiny decoder is duplicated here. */
function decodeDataUrl(src: string | undefined): Uint8Array | null {
  if (!src) return null;
  const match = src.match(/^data:image\/[\w.+-]+;base64,(.+)$/);
  if (!match) return null;
  const binary = atob(match[1]);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

// Cache decoded natural sizes by src: appendTransaction runs on every docChanged
// transaction, and a large imported DOCX can carry dozens of embedded data URLs —
// re-running atob + imageMeta per image per keystroke is O(images × bytes)/tx.
// The src is the sole decode input and never changes for a given image, so it
// keys the result (null is cached too: non-data URLs / unreadable bytes never
// become decodable).
const naturalSizeCache = new Map<string, { width: number; height: number } | null>();

/** Natural pixel dimensions of a data-URL image, read synchronously from the file
 *  header (no load/decode round-trip). Returns null for http(s):// URLs (can't be
 *  sync-decoded) or unreadable bytes — the CSS max-width fallback constrains
 *  those. */
function naturalSize(src: string | undefined): { width: number; height: number } | null {
  if (!src) return null;
  const cached = naturalSizeCache.get(src);
  if (cached !== undefined) return cached;
  const bytes = decodeDataUrl(src);
  let result: { width: number; height: number } | null = null;
  if (bytes) {
    try {
      const meta = imageMeta(bytes);
      if (typeof meta.width === "number" && typeof meta.height === "number") {
        result = { width: meta.width, height: meta.height };
      }
    } catch {
      // Unreadable or unsupported image — leave to the CSS fallback.
    }
  }
  naturalSizeCache.set(src, result);
  return result;
}

/** The section geometry governing the image at `pos`. OOXML sectPr rides a
 *  section's LAST paragraph (paragraph.attrs.sectionProperties) and governs it
 *  + everything before it; doc.attrs.sectionProperties is the final section.
 *  Prefers the enclosing page's stamped section (reflow copies that source onto
 *  each page); falls back to scanning the section-ending paragraph when the page
 *  isn't stamped yet — reflow rebuilds page.sp from this source, so it can be
 *  briefly null right after a setContent/paste, and this fallback reads the
 *  same source reflow will, instead of dropping to the DOM default. */
function sectionAt(doc: PmNode, pos: number): unknown {
  const $pos = doc.resolve(pos);
  for (let d = $pos.depth; d > 0; d--) {
    const node = $pos.node(d);
    if (node.type.name === "page") {
      const sp = (node.attrs as { sectionProperties?: unknown }).sectionProperties;
      if (sp) return sp;
    }
  }
  // Fallback: the first section-ending paragraph at/after the image's paragraph
  // (its sectPr ends the section the image is in).
  let paraPos = pos;
  for (let d = $pos.depth; d > 0; d--) {
    if ($pos.node(d).type.name === "paragraph") {
      paraPos = $pos.before(d);
      break;
    }
  }
  let section = (doc.attrs as { sectionProperties?: unknown }).sectionProperties ?? null;
  let done = false;
  doc.descendants((node, nodePos) => {
    if (done || nodePos < paraPos) return true;
    if (node.type.name === "paragraph") {
      const sp = (node.attrs as { sectionProperties?: unknown }).sectionProperties;
      if (sp != null) {
        section = sp;
        done = true;
        return false;
      }
    }
    return true;
  });
  return section;
}

/** Pick the first image file on a paste/drop data transfer, else null. */
function pickImageFile(dt: DataTransfer | null): File | null {
  if (!dt) return null;
  return Array.from(dt.files ?? []).find((f) => f.type.startsWith("image/")) ?? null;
}

/** Read an image file as a data URL and insert an image node at `pos` (or the
 *  caret when `pos` is null). data URL — not the browser's default blob: URL —
 *  because a blob: URL can't be sync-decoded (no width cap below) and can't be
 *  exported (blob: URLs don't base64-encode into DOCX). */
function readAndInsert(view: EditorView, file: File, pos: number | null): void {
  const reader = new FileReader();
  reader.onload = () => {
    const src = reader.result;
    if (typeof src !== "string") return;
    const { state, dispatch } = view;
    const node = state.schema.nodes.image.create({ src });
    dispatch(pos != null ? state.tr.insert(pos, node) : state.tr.replaceSelectionWith(node));
  };
  reader.readAsDataURL(file);
}

/** Usable content width of the first rendered page (px), read from the DOM.
 *  Fallback for documents whose pages carry no sectionProperties (e.g. a blank
 *  editor before a page setup is applied): the page box still renders at a CSS
 *  default paper size + margin, so the content box is measurable without OOXML
 *  geometry. Page geometry is stable across edits, so reading it from the
 *  pre-update DOM inside appendTransaction is fine. */
function domContentWidth(view: EditorView | null): number | null {
  if (!view) return null;
  const page = view.dom.querySelector(".docen-page");
  if (!(page instanceof HTMLElement)) return null;
  const cs = getComputedStyle(page);
  const w = page.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
  return w > 0 ? w : null;
}

/** Fetch an http(s) image and read it as a data URL (so it can be sync-decoded
 *  for the cap below and embedded into DOCX on export). Returns null on
 *  network/CORS failure — the image keeps its http URL (CSS-only fallback). */
async function fetchToDataUrl(url: string): Promise<string | null> {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    const blob = await res.blob();
    return await new Promise<string | null>((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(typeof reader.result === "string" ? reader.result : null);
      reader.onerror = () => resolve(null);
      reader.readAsDataURL(blob);
    });
  } catch {
    return null;
  }
}

/** Stamp `patch` onto every image node whose `src` matches, in one transaction.
 *  No-op when none match or the view is gone. */
function markupMatchingSrc(view: EditorView, src: string, patch: Record<string, unknown>): void {
  if (view.isDestroyed) return;
  const { state, dispatch } = view;
  let tr = state.tr;
  let hit = false;
  state.doc.descendants((node, pos) => {
    if (node.type.name === "image" && node.attrs.src === src) {
      tr = tr.setNodeMarkup(pos, null, { ...node.attrs, ...patch });
      hit = true;
    }
    return true;
  });
  if (hit) dispatch(tr);
}

/** Scale (width, height) DOWN to `contentW`, never upscale — mirroring MS
 *  Office's image cap: a wider image takes the content width and a height
 *  scaled by the same factor. The single cap arithmetic shared by the
 *  appendTransaction pass and embedHttpImage's CORS fallback. */
function capSize(
  width: number,
  height: number | null,
  contentW: number,
): { width: number; height: number | null } {
  if (width <= contentW) return { width, height };
  const scale = contentW / width;
  return { width: contentW, height: height != null ? Math.round(height * scale) : null };
}

/** Natural pixel dimensions of the in-DOM <img> with `src`, awaited on load.
 *  Used when fetch is blocked (CORS): an <img> loads cross-origin freely — only
 *  fetch/canvas are CORS-gated — so the natural size is still readable even
 *  though the bytes can't be fetched for embedding. */
function domNaturalSize(
  view: EditorView,
  src: string,
): Promise<{ width: number; height: number } | null> {
  return new Promise((resolve) => {
    const img = Array.from(view.dom.querySelectorAll<HTMLImageElement>("img")).find(
      (i) => i.getAttribute("src") === src,
    );
    if (!img) return resolve(null);
    const read = () =>
      img.naturalWidth > 0 && img.naturalHeight > 0
        ? resolve({ width: img.naturalWidth, height: img.naturalHeight })
        : resolve(null);
    if (img.complete && img.naturalWidth > 0) return read();
    img.addEventListener("load", read, { once: true });
    img.addEventListener("error", () => resolve(null), { once: true });
  });
}

/** Inline an http(s) image. On a successful fetch, swap in the data URL so
 *  appendTransaction can cap it (via imageMeta) and renderDocx can embed it. On
 *  fetch failure (CORS/network) the bytes can't be embedded — but the <img>
 *  still loads cross-origin, so its natural dimensions are stamped onto the node
 *  and appendTransaction caps the width to the content area (the src stays http
 *  and is dropped on export). Word inlines pasted web images the same way. */
async function embedHttpImage(view: EditorView, src: string, fetching: Set<string>): Promise<void> {
  try {
    const dataUrl = await fetchToDataUrl(src);
    if (dataUrl) {
      markupMatchingSrc(view, src, { src: dataUrl });
      return;
    }
    if (view.isDestroyed) return;
    const natural = await domNaturalSize(view, src);
    if (natural) {
      // A CORS-blocked web image (manually pasted) can't be sync-decoded, so cap
      // it here from the DOM-read natural size — appendTransaction skips any
      // image that already carries a width, so this is its only cap pass.
      const contentW =
        sectionContentDims(
          (view.state.doc.attrs as { sectionProperties?: unknown }).sectionProperties,
        )?.width ?? domContentWidth(view);
      markupMatchingSrc(
        view,
        src,
        contentW && contentW > 0
          ? capSize(natural.width, natural.height, contentW)
          : { width: natural.width, height: natural.height },
      );
    }
  } finally {
    fetching.delete(src);
  }
}

/**
 * Word-style image width capping. MS Office scales an inline image DOWN to the
 * section content width (page width − margins) when it is wider, and otherwise
 * leaves it at its real size (never upscales) — and the cap is a real dimension
 * change (exported DOCX carries the capped size, not just a visual constraint).
 *
 * Scope: ONLY manually inserted images (ribbon Insert, paste, drop) — these
 * enter with `src` and no width, so appendTransaction caps them. Images loaded
 * from a DOCX (`openDOCX`/`parseDOCX`) carry their source `wp:extent` as width
 * and are left untouched (capping them would distort the source document's
 * sizing); the page's CSS max-width keeps any oversized import on the page
 * visually without changing its stored dimensions.
 *
 * Implemented as an `appendTransaction` plugin (idempotent: once capped,
 * `width === contentW`, so the next pass skips it). Section geometry gives the
 * exact content width; when absent (blank doc, before a page setup) the rendered
 * page box is measured instead. An http image is inlined first (fetch → data
 * URL, or DOM natural size on a CORS failure); the data-URL path is capped
 * here, the CORS-failure path is capped in embedHttpImage from the DOM size.
 */
export const ImageCap = Extension.create({
  name: "docenImageCap",
  addProseMirrorPlugins() {
    // Captured by the appendTransaction closure so it can read the live page
    // DOM for the content-width fallback (section geometry may be absent).
    let editorView: EditorView | null = null;
    // http(s) image URLs currently being fetched → data URL, deduped so a
    // re-flow doesn't kick off a second fetch for the same URL.
    const fetching = new Set<string>();
    return [
      new Plugin({
        key,
        view(v) {
          editorView = v;
          return {
            update(view, prevState) {
              if (view.state.doc === prevState.doc) return;
              // Inline any http(s) image: fetch → data URL so it can be capped
              // and exported (a remote URL alone can't be sync-decoded or
              // embedded into DOCX). Word inlines pasted web images too.
              view.state.doc.descendants((node) => {
                if (node.type.name !== "image") return true;
                const src = node.attrs.src;
                if (typeof src !== "string" || !/^https?:/.test(src)) return true;
                if (fetching.has(src)) return true;
                fetching.add(src);
                void embedHttpImage(view, src, fetching);
                return true;
              });
            },
          };
        },
        props: {
          // Intercept pasted/dropped image FILES: read them as data URLs and
          // insert, instead of letting the browser insert a blob: URL (which
          // skips the cap below and can't be exported). Returns true to claim
          // the event so the default blob: insert is suppressed.
          handlePaste: (view, event: ClipboardEvent) => {
            const file = pickImageFile(event.clipboardData);
            if (!file) return false;
            readAndInsert(view, file, null);
            return true;
          },
          handleDrop: (view, event: DragEvent) => {
            const file = pickImageFile(event.dataTransfer);
            if (!file) return false;
            const pos = view.posAtCoords({ left: event.clientX, top: event.clientY })?.pos ?? null;
            readAndInsert(view, file, pos);
            return true;
          },
        },
        appendTransaction: (trs, _oldState, newState) => {
          // Only react to document-changing transactions (skip selection-only).
          if (!trs.some((tr) => tr.docChanged)) return null;
          const doc = newState.doc;
          const tr = newState.tr;
          let changed = false;

          // The DOM content-width fallback (querySelector + getComputedStyle) is
          // the same for every image when section geometry is absent, so read it
          // once before the walk instead of per image.
          const fallbackW = domContentWidth(editorView);
          doc.descendants((node, pos) => {
            if (node.type.name !== "image") return true;
            const attrs = node.attrs as {
              src?: string;
              width?: number | null;
              height?: number | null;
            };
            // Only cap images with NO explicit width — i.e. manually
            // pasted/dropped/inserted images (readAndInsert and #onImageChange
            // create them with src only). Imported images carry their source
            // extent as width (parseDocx: wp:extent → px); re-capping those
            // would distort the source document's image sizing, so they stay at
            // their real size (the page's CSS max-width keeps them on the page
            // visually). An already-capped manual image keeps width == contentW,
            // so it's skipped too (idempotent). A CORS-blocked web image is
            // capped in embedHttpImage, where its width is stamped.
            if (attrs.width != null) return true;
            const natural = naturalSize(attrs.src);
            const displayW = natural?.width;
            if (displayW == null || displayW <= 0) return true;

            // Section content width from OOXML geometry; fall back to the
            // rendered page's content box when geometry is absent (blank doc).
            const contentW = sectionContentDims(sectionAt(doc, pos))?.width ?? fallbackW;
            if (!contentW || contentW <= 0) return true;
            if (displayW <= contentW) return true; // within bounds — never upscale

            const capped = capSize(displayW, attrs.height ?? natural?.height ?? null, contentW);
            tr.setNodeMarkup(pos, null, { ...attrs, width: capped.width, height: capped.height });
            changed = true;
            return true;
          });

          if (!changed) return null;
          tr.setMeta(key, true);
          return tr;
        },
      }),
    ];
  },
});
