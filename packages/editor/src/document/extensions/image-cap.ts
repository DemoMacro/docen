import { Extension } from "@docen/docx/core";
import { encodeBase64 } from "@office-open/core";
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

// http(s) srcs whose embed fetch already failed (CORS/network). Cached so a
// docChanged transaction doesn't re-fetch on every keystroke — the src stays
// http after a failure, so the walk would otherwise re-queue it each edit.
const failedImageSrcs = new Set<string>();

/** Clear the natural-size + failed-src caches. Call on document load so caches
 *  from a prior document neither grow unbounded nor suppress a legitimate
 *  re-fetch of a src that failed under a transient (network/CORS) state. */
export function clearImageCapCache(): void {
  naturalSizeCache.clear();
  failedImageSrcs.clear();
}

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
 *  for the cap below and embedded into DOCX on export). MIME is inferred from
 *  the image header bytes (imageMeta), not the URL extension — CDN/auth/token
 *  URLs (`img.png?token=`, `/avatar`) would otherwise mislabel a JPEG as PNG
 *  and corrupt the embedded blip. Returns `{ dataUrl, timedOut }`: dataUrl null
 *  on network/CORS failure or timeout; timedOut distinguishes the two so the
 *  caller stamps a `timeout` vs `error` placeholder. A custom handler
 *  (proxy/auth) has no abort hook, so its timeout is a Promise.race. */
async function fetchToDataUrl(
  url: string,
  handler: ((url: string) => Promise<Uint8Array>) | undefined,
  timeoutMs: number,
): Promise<{ dataUrl: string | null; timedOut: boolean }> {
  let bytes: Uint8Array | null;
  if (handler) {
    let timedOut = false;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const timeout = new Promise<null>((resolve) => {
      timer = setTimeout(() => {
        timedOut = true;
        resolve(null);
      }, timeoutMs);
    });
    try {
      bytes = await Promise.race([handler(url).catch(() => null), timeout]);
    } finally {
      if (timer) clearTimeout(timer);
    }
    if (!bytes) return { dataUrl: null, timedOut };
  } else {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const res = await fetch(url, { signal: controller.signal });
      if (!res.ok) return { dataUrl: null, timedOut: false };
      bytes = new Uint8Array(await res.arrayBuffer());
    } catch (error) {
      const timedOut = error instanceof DOMException && error.name === "AbortError";
      return { dataUrl: null, timedOut };
    } finally {
      clearTimeout(timer);
    }
  }
  let mime = "image/png";
  try {
    const type = imageMeta(bytes).type;
    if (type) mime = type === "jpg" ? "image/jpeg" : `image/${type}`;
  } catch {
    // Unreadable/unknown header — leave the PNG fallback.
  }
  return { dataUrl: `data:${mime};base64,${encodeBase64(bytes)}`, timedOut: false };
}

/** Stamp `patch` onto every image node whose `src` matches, in one transaction.
 *  `patch` may be a function of the node's current attrs, so per-node decisions
 *  (e.g. skip sizing when the node already carries explicit dimensions, or skip
 *  a node whose load state shouldn't change) are made against the live attrs at
 *  each match. A function may return null/`{}` to skip that node (no markup, no
 *  reflow) — used by embedHttpImage's failure branch to leave a sized image
 *  untouched. No-op when none match or the view is gone. */
function markupMatchingSrc(
  view: EditorView,
  src: string,
  patch:
    | Record<string, unknown>
    | ((attrs: Record<string, unknown>) => Record<string, unknown> | null),
): void {
  if (view.isDestroyed) return;
  const { state, dispatch } = view;
  let tr = state.tr;
  let hit = false;
  state.doc.descendants((node, pos) => {
    if (node.type.name === "image" && node.attrs.src === src) {
      const p = typeof patch === "function" ? patch(node.attrs as Record<string, unknown>) : patch;
      if (p && Object.keys(p).length > 0) {
        tr = tr.setNodeMarkup(pos, null, { ...node.attrs, ...p });
        hit = true;
      }
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

/** Natural pixel dimensions of the image at `src`. Used when fetch is blocked
 *  (CORS): an <img> loads cross-origin freely — only fetch/canvas are
 *  CORS-gated — so the natural size is still readable even though the bytes
 *  can't be fetched for embedding.
 *
 *  Probes via a standalone `new Image()`, NOT the rendered <img>: the rendered
 *  image carries `loading="lazy"` (image.ts renderHTML), which defers
 *  off-screen loads, so its load event may never fire and an await on it would
 *  hang — deadlocking the embed throttle (inFlight never decrements). A
 *  JS-created Image is outside the render tree, so lazy never applies and
 *  load/error always fire regardless of viewport proximity. A `timeoutMs` race
 *  caps a slow/hung probe so it can't hold an inFlight slot forever (a
 *  slow-loading web image would otherwise park the promise indefinitely). */
function domNaturalSize(
  src: string,
  timeoutMs: number,
): Promise<{ width: number; height: number } | null> {
  return new Promise((resolve) => {
    const probe = new Image();
    let done = false;
    const finish = (result: { width: number; height: number } | null): void => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      resolve(result);
    };
    const timer = setTimeout(() => {
      probe.onload = null;
      probe.onerror = null;
      probe.src = "";
      finish(null);
    }, timeoutMs);
    probe.onload = () =>
      finish(
        probe.naturalWidth > 0 && probe.naturalHeight > 0
          ? { width: probe.naturalWidth, height: probe.naturalHeight }
          : null,
      );
    probe.onerror = () => finish(null);
    probe.src = src;
  });
}

/** Inline an http(s) image. On a successful fetch, swap in the data URL so
 *  appendTransaction can cap it (via imageMeta) and renderDocx can embed it. On
 *  fetch failure/timeout (CORS/network) the bytes can't be embedded — but the
 *  <img> still loads cross-origin, so its natural dimensions may still be
 *  readable for the cap. Word inlines pasted web images the same way.
 *
 *  Only UNSIZED http images (width == null, stamped loadState=loading by
 *  appendTransaction) go through the placeholder state machine: loading →
 *  (loaded | error | timeout), each stamp settling a size so measure converges
 *  in one reflow instead of once per fetch. A SIZED http image is left alone on
 *  both success (its width is already correct — swap src only) and failure
 *  (layout settled, no placeholder) — zero placeholder interference, per the
 *  "judge by size, not by protocol" rule. */
async function embedHttpImage(
  view: EditorView,
  src: string,
  fetching: Set<string>,
  opts: {
    handler: ((url: string) => Promise<Uint8Array>) | undefined;
    fetchTimeoutMs: number;
    probeTimeoutMs: number;
  },
): Promise<void> {
  try {
    const { dataUrl, timedOut } = await fetchToDataUrl(src, opts.handler, opts.fetchTimeoutMs);
    if (dataUrl) {
      const natural = naturalSize(dataUrl);
      const contentW =
        sectionContentDims(
          (view.state.doc.attrs as { sectionProperties?: unknown }).sectionProperties,
        )?.width ?? domContentWidth(view);
      markupMatchingSrc(view, src, (attrs) => {
        const a = attrs as { loadState?: string | null };
        // A sized non-placeholder node keeps its width — swap src only (measure
        // and export already have the size). A loading placeholder (or any
        // unsized node) refines to the data URL's real dimensions and clears
        // the placeholder state.
        if (attrs.width != null && a.loadState !== "loading") return { src: dataUrl };
        if (!natural) return { src: dataUrl, loadState: null };
        if (contentW && natural.width > contentW) {
          return {
            src: dataUrl,
            loadState: null,
            ...capSize(natural.width, natural.height, contentW),
          };
        }
        return { src: dataUrl, loadState: null, width: natural.width, height: natural.height };
      });
      return;
    }
    // fetch failed or timed out: the src stays http. Remember it so the next
    // docChanged walk doesn't re-fetch on every keystroke (retry storm).
    failedImageSrcs.add(src);
    if (view.isDestroyed) return;
    const nextState = timedOut ? "timeout" : "error";
    // Only a loading placeholder needs an error/timeout stamp — a sized image's
    // layout is already settled, so leave it (the <img> shows its own broken
    // icon). The placeholder keeps a settled size (natural if readable, else
    // the 4:3 width appendTransaction already stamped) so measure doesn't
    // reflow.
    const natural = await domNaturalSize(src, opts.probeTimeoutMs);
    markupMatchingSrc(view, src, (attrs) => {
      const a = attrs as { loadState?: string | null };
      if (a.loadState !== "loading") return null;
      if (!natural) return { loadState: nextState };
      const contentW =
        sectionContentDims(
          (view.state.doc.attrs as { sectionProperties?: unknown }).sectionProperties,
        )?.width ?? domContentWidth(view);
      return contentW && contentW > 0
        ? { loadState: nextState, ...capSize(natural.width, natural.height, contentW) }
        : { loadState: nextState, width: natural.width, height: natural.height };
    });
  } finally {
    fetching.delete(src);
  }
}

/**
 * Office-style image width capping. MS Office scales an inline image DOWN to the
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
  addOptions() {
    return {
      // Concurrent http-image fetches. 6 suits HTTP/2 CDNs (multi-plexed, no
      // per-host ceiling) while bounding main-thread base64 transcoding; 4 was
      // conservative for HTTP/1.1's 6-connections-per-host limit. Override via
      // ImageCap.configure({ maxConcurrent }).
      maxConcurrent: 6,
      // Custom fetch handler (proxy/auth/cache), same shape as prepareImages'
      // ImageFetchHandler — lets a CORS-blocked source be rerouted. undefined →
      // global fetch with an AbortController timeout. Exposing this to consumers
      // (<docen-document> attr → addin) is a follow-up; the option is wired now.
      fetchHandler: undefined as ((url: string) => Promise<Uint8Array>) | undefined,
      fetchTimeoutMs: 30_000,
      probeTimeoutMs: 10_000,
    };
  },
  addProseMirrorPlugins() {
    const { maxConcurrent, fetchHandler, fetchTimeoutMs, probeTimeoutMs } = this.options;
    const embedOpts = { handler: fetchHandler, fetchTimeoutMs, probeTimeoutMs };
    // Captured by the appendTransaction closure so it can read the live page
    // DOM for the content-width fallback (section geometry may be absent).
    let editorView: EditorView | null = null;
    // http(s) image URLs currently being fetched → data URL, deduped so a
    // re-flow doesn't kick off a second fetch for the same URL.
    const fetching = new Set<string>();
    // Concurrency throttle: a doc with many large http images would otherwise
    // fire one fetch per image at once — a network storm plus piled-up base64
    // transcoding on the main thread. Pending URLs queue and drain at most
    // maxConcurrent at a time.
    const queue: string[] = [];
    let inFlight = 0;
    const drain = (): void => {
      const view = editorView;
      if (!view || view.isDestroyed) return;
      while (inFlight < maxConcurrent && queue.length > 0) {
        const src = queue.shift() as string;
        inFlight++;
        void embedHttpImage(view, src, fetching, embedOpts).finally(() => {
          inFlight--;
          drain();
        });
      }
    };
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
              // Queued + drained by drain() to bound concurrent fetches.
              view.state.doc.descendants((node) => {
                if (node.type.name !== "image") return true;
                const src = node.attrs.src;
                if (typeof src !== "string" || !/^https?:/.test(src)) return true;
                if (fetching.has(src) || failedImageSrcs.has(src)) return true;
                fetching.add(src);
                queue.push(src);
                return true;
              });
              drain();
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

          // The DOM content-width fallback (querySelector + getComputedStyle,
          // which can force layout) is the same for every image when section
          // geometry is absent. Read it lazily — only once the walk actually
          // hits an uncapped image — so pure-text edits (no images) never
          // touch the DOM per keystroke.
          let fallbackW: number | null | undefined;
          const getFallbackW = (): number | null => {
            if (fallbackW !== undefined) return fallbackW;
            return (fallbackW = domContentWidth(editorView));
          };
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
            // Unsized images (manual paste/insert — readAndInsert sets src only):
            // data-URL images are sync-read via imageMeta and capped; http images
            // can't be sync-read (naturalSize returns null) — refined async by embedHttpImage.
            const natural = naturalSize(attrs.src);
            const displayW = natural?.width;
            // Section content width from OOXML geometry; fall back to the
            // rendered page's content box when geometry is absent (blank doc).
            const contentW = sectionContentDims(sectionAt(doc, pos))?.width ?? getFallbackW();
            if (displayW == null || displayW <= 0) {
              // Unreadable size. A corrupt data URL (imageMeta can't read it)
              // has no async refiner → fall back to contentW × 0.75 (4:3,
              // matching the editor CSS placeholder + measure.placeholder) so
              // edit/measure/export agree. An http image gets the SAME 4:3 box
              // PLUS loadState=loading: a settled placeholder size lets measure
              // converge without reflowing once per fetch, and embedHttpImage
              // clears/flips loadState and refines to the real size when the
              // fetch resolves (success/error/timeout). embedHttpImage's
              // "already has width → swap src only" branch recognizes
              // loadState=loading and overrides the size, so stamping a
              // placeholder width here does NOT strand the image at 4:3.
              if (!contentW || contentW <= 0) return true;
              const isHttp = typeof attrs.src === "string" && /^https?:\/\//.test(attrs.src);
              tr.setNodeMarkup(pos, null, {
                ...attrs,
                width: contentW,
                height: Math.round(contentW * 0.75),
                ...(isHttp ? { loadState: "loading" } : {}),
              });
              changed = true;
              return true;
            }
            if (!contentW || contentW <= 0) return true;
            if (displayW <= contentW) {
              // Within bounds — keep natural size (never upscale), but stamp it so
              // the editor <img> doesn't fall back to width:100% (which would
              // upscale) and export matches display. Mirrors embedHttpImage's
              // small-image branch (edit == render == export).
              tr.setNodeMarkup(pos, null, {
                ...attrs,
                width: displayW,
                height: attrs.height ?? natural?.height ?? null,
              });
              changed = true;
              return true;
            }

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
