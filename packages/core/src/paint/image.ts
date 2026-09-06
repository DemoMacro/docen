import type { LayoutDrawingMember } from "@docen/layout";
import {
  Image as LeaferImage,
  ImageManager,
  Rect,
  Resource,
  type IGroup,
  type ILeaferImage,
} from "leafer-ui";

import type { PaintContext } from "./context";

/** Leafer evicts a decoded image larger than its 4MP cache threshold the
 *  moment a paint's use count drops back to zero (ImageManager.recycle →
 *  Resource.remove), so page recycling under scroll re-loaded every big
 *  banner/photo from its data URL and painted it a second late — the pop-in.
 *  One pinned use per media url keeps the decoded entry resident for the
 *  stage's lifetime: every later paint hits the ready entry and renders
 *  synchronously. The per-position `LeaferImage` elements stay thin shells —
 *  Leafer's paint resolves the bitmap through this shared entry. */
const pinnedImages = new Map<string, ILeaferImage>();

export function pinImage(url: string): ILeaferImage {
  let image = pinnedImages.get(url);
  // A tree clear can briefly drive the shared entry's use count to zero, and
  // Leafer's recycle judges eviction on a setTimeout — after the repaint has
  // already re-added the leaf. The Resource-table entry then disappears while
  // this stale handle still reports ready, and every later paint resolves the
  // url into a fresh async decode nothing ever frames. Re-anchor instead.
  if (image && !Resource.get(url)) {
    pinnedImages.delete(url);
    image = undefined;
  }
  if (!image) {
    image = ImageManager.get({ url }, "image");
    pinnedImages.set(url, image);
    image.load();
  }
  return image;
}

/** Composites (cropped views, masked-run blends) as cached data URLs: the
 *  decode → canvas work → toDataURL chain costs seconds per repainted member,
 *  and repaints happen per transaction — caching the derived url by content
 *  fingerprint turns every repaint after the first into a plain picture. */
const derivedImages = new Map<string, string>();
const DERIVED_LIMIT = 64;

function cacheDerived(fingerprint: string, url: string): void {
  if (derivedImages.size >= DERIVED_LIMIT) {
    derivedImages.delete(derivedImages.keys().next().value!);
  }
  derivedImages.set(fingerprint, url);
}

/** Release the pins at stage teardown — Leafer's own recycle then evicts the
 *  large entries it no longer shares. Derived composites die with the
 *  document they were produced from. */
export function releasePinnedImages(): void {
  for (const image of pinnedImages.values()) ImageManager.recycle(image);
  pinnedImages.clear();
  derivedImages.clear();
}

function plainImageLeaf(
  src: string,
  x: number,
  y: number,
  width: number,
  height: number,
  flipH?: boolean,
  flipV?: boolean,
): LeaferImage {
  return new LeaferImage({
    url: src,
    x: flipH ? x + width : x,
    y: flipV ? y + height : y,
    width,
    height,
    // Mirrors flip around the element's (x,y) origin: shifting the origin
    // to the far edge first makes the reflection cover the original box.
    ...(flipH ? { scaleX: -1 } : {}),
    ...(flipV ? { scaleY: -1 } : {}),
  });
}

/** The one way a bitmap joins the tree. A ready resource joins synchronously;
 *  on the first decode a placeholder keeps the paint-order slot open, and the
 *  real leaf replaces it through the same swap — nothing else inserts images.
 *  An existing slot (a composite produced asynchronously, see
 *  {@link addDerivedImage}) is replaced in place instead of joined at the end
 *  so the member's paint-order z-position survives the decode wait. */
function placeImage(
  tree: IGroup,
  image: ILeaferImage,
  src: string,
  x: number,
  y: number,
  width: number,
  height: number,
  ctx: PaintContext,
  flipH?: boolean,
  flipV?: boolean,
  slot?: Rect,
): void {
  if (image.ready) {
    const leaf = plainImageLeaf(src, x, y, width, height, flipH, flipV);
    if (slot) swapSlot(tree, slot, leaf, ctx);
    else tree.add(leaf);
    return;
  }
  const held = slot ?? new Rect({ x, y, width, height });
  if (!slot) tree.add(held);
  image.load(() => {
    // A repaint since the decode started cleared the tree (slot included) —
    // that repaint's own request now owns the paint-order slot.
    if (!held.parent) return;
    swapSlot(tree, held, plainImageLeaf(src, x, y, width, height, flipH, flipV), ctx);
  });
}

/** The decode resolves mid-frame: swapping immediately races the frame's
 *  already-consumed layout plan, and the forced render can miss the element
 *  entirely. Slide the swap to the next frame — the placeholder keeps the
 *  box open until then. */
function swapSlot(tree: IGroup, slot: Rect, leaf: LeaferImage, ctx: PaintContext): void {
  requestAnimationFrame(() => {
    if (!slot.parent) return;
    tree.addAfter(leaf, slot);
    tree.remove(slot);
    ctx.rerender();
  });
}

/** A decoded-plain picture at a box. Shared by drawing members and inline
 *  picture atoms. */
export function addDecodedImage(
  tree: IGroup,
  src: string,
  x: number,
  y: number,
  width: number,
  height: number,
  ctx: PaintContext,
  flipH?: boolean,
  flipV?: boolean,
): void {
  placeImage(tree, pinImage(src), src, x, y, width, height, ctx, flipH, flipV);
}

export function addPlainImage(
  tree: IGroup,
  m: Extract<LayoutDrawingMember, { kind: "picture" }>,
  mx: number,
  my: number,
  ctx: PaintContext,
): void {
  addDecodedImage(tree, m.src!, mx, my, m.width, m.height, ctx, m.flipH, m.flipV);
}

/** A derived composite (a cropped view or a blended run) joined at a box:
 *  the cache serves the composite as a plain picture; a first production
 *  keeps the paint-order slot open while `produce` works off-thread-bound
 *  canvases and delivers the composite url (undefined = the sources failed —
 *  the placeholder goes). `produce` runs synchronously here and delivers
 *  asynchronously; it receives the slot so its fallback decisions can check
 *  whether this request still owns the paint-order position. */
function addDerivedImage(
  tree: IGroup,
  fingerprint: string,
  produce: (deliver: (url: string | undefined) => void, slot: Rect) => void,
  x: number,
  y: number,
  width: number,
  height: number,
  ctx: PaintContext,
  flipH?: boolean,
  flipV?: boolean,
): void {
  const cached = derivedImages.get(fingerprint);
  if (cached) {
    placeImage(tree, pinImage(cached), cached, x, y, width, height, ctx, flipH, flipV);
    return;
  }
  const slot = new Rect({ x, y, width, height });
  tree.add(slot);
  const deliver = (url: string | undefined): void => {
    // A repaint since the work started cleared the tree (slot included) —
    // that repaint's own request now owns the paint-order slot.
    if (!slot.parent) return;
    if (url === undefined) {
      tree.remove(slot);
      return;
    }
    cacheDerived(fingerprint, url);
    placeImage(tree, pinImage(url), url, x, y, width, height, ctx, flipH, flipV, slot);
  };
  produce(deliver, slot);
}

/** A run of masked GDI blt members (SRCPAINT/SRCAND halves, optionally over a
 *  plain backdrop picture): flattened in record order through canvas
 *  `screen`/`multiply` compositing — the ternary raster-op semantics — and
 *  joined as one image. Decode failures drop that member; if nothing survives
 *  the run falls back to individual painting. */
export function addBlendedPictureRun(
  tree: IGroup,
  run: Extract<LayoutDrawingMember, { kind: "picture" }>[],
  boxX: number,
  boxY: number,
  ctx: PaintContext,
): void {
  const x0 = Math.min(...run.map((p) => p.x));
  const y0 = Math.min(...run.map((p) => p.y));
  const width = Math.ceil(Math.max(...run.map((p) => p.x + p.width)) - x0);
  const height = Math.ceil(Math.max(...run.map((p) => p.y + p.height)) - y0);
  if (width < 1 || height < 1 || width > 8192 || height > 8192) return;
  const fingerprint = `blend|${run
    .map((p) => `${p.src}|${p.blend}|${p.x},${p.y},${p.width},${p.height}`)
    .join(";")}`;
  addDerivedImage(
    tree,
    fingerprint,
    (deliver, slot) => {
      const loads = run.map(
        (p) =>
          new Promise<HTMLImageElement | null>((resolve) => {
            const el = new Image();
            el.onload = () => resolve(el);
            el.onerror = () => resolve(null);
            el.src = p.src!;
          }),
      );
      void Promise.all(loads).then((decoded) => {
        // A repaint since the decode started cleared the tree (slot included)
        // — that repaint's own run now owns the paint-order slot.
        if (!slot.parent) return;
        if (!decoded.some(Boolean)) {
          deliver(undefined);
          // Masked halves never paint alone (their opaque mask background
          // would slab the page) — only plain backdrops fall back. The
          // deliver above removed the run's slot, so these join at the tail
          // of whatever the current tree holds.
          for (const p of run) if (!p.blend) addPlainImage(tree, p, boxX + p.x, boxY + p.y, ctx);
          return;
        }
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const c2d = canvas.getContext("2d")!;
        // GDI replays these blts against its live destination surface — the
        // page behind the metafile — so the composite stays transparent
        // wherever the records keep the destination: a screen half only
        // marks the shape mask (never painted), and a multiply half lands
        // just its colored content inside that mask. Page color and lower
        // members show through.
        let maskData: Uint8ClampedArray | undefined;
        for (let k = 0; k < run.length; k++) {
          const img = decoded[k];
          if (!img) continue;
          const dx = run[k].x - x0;
          const dy = run[k].y - y0;
          const dw = run[k].width;
          const dh = run[k].height;
          if (run[k].blend === "screen") {
            maskData = shapeMaskAt(img, dx, dy, dw, dh, width, height);
            continue;
          }
          if (run[k].blend === "multiply") {
            const content = maskedContent(img, dx, dy, dw, dh, maskData, width);
            maskData = undefined;
            if (!content) continue;
            c2d.drawImage(content, dx, dy, dw, dh);
            continue;
          }
          maskData = undefined;
          c2d.drawImage(img, dx, dy, dw, dh);
        }
        deliver(canvas.toDataURL("image/png"));
      });
    },
    boxX + x0,
    boxY + y0,
    width,
    height,
    ctx,
  );
}

/** A screen half's brightness as the shape mask, sampled on the run's union
 *  box: each pixel's alpha takes its max channel — the 1bpp white shape
 *  lights up, the black backdrop drops out. This is the raster-op's "where
 *  the shape is" term, derived from the record's own bytes. */
function shapeMaskAt(
  img: HTMLImageElement,
  dx: number,
  dy: number,
  dw: number,
  dh: number,
  width: number,
  height: number,
): Uint8ClampedArray {
  const c = document.createElement("canvas");
  c.width = width;
  c.height = height;
  const g = c.getContext("2d", { willReadFrequently: true })!;
  g.drawImage(img, dx, dy, dw, dh);
  const d = g.getImageData(0, 0, width, height);
  const a = d.data;
  for (let i = 0; i < a.length; i += 4) a[i + 3] = Math.max(a[i], a[i + 1], a[i + 2]);
  return a;
}

/** A multiply half reduced to its colored content inside the pending shape
 *  mask: per pixel, white keeps the destination (alpha 0) and every other
 *  color lands verbatim at the mask's coverage — GDI's AND over a white page
 *  writes the source pixel itself wherever it is not white, not a blend. A
 *  distance-from-white ramp would turn light fills and antialiased ink into
 *  translucent washes over the page. An unconsumed multiply half falls back
 *  to its own non-white key. */
function maskedContent(
  img: HTMLImageElement,
  dx: number,
  dy: number,
  dw: number,
  dh: number,
  maskData: Uint8ClampedArray | undefined,
  maskWidth: number,
): HTMLCanvasElement | undefined {
  if (dw < 1 || dh < 1) return undefined;
  const c = document.createElement("canvas");
  c.width = dw;
  c.height = dh;
  const g = c.getContext("2d", { willReadFrequently: true })!;
  g.drawImage(img, 0, 0, dw, dh);
  const d = g.getImageData(0, 0, dw, dh);
  const a = d.data;
  for (let j = 0; j < dh; j++) {
    for (let i = 0; i < dw; i++) {
      const p = (j * dw + i) * 4;
      const m =
        maskData && dy + j >= 0 && dy + j < maskData.length / 4 / maskWidth
          ? maskData[((dy + j) * maskWidth + dx + i) * 4 + 3]
          : 255;
      a[p + 3] = Math.min(a[p], a[p + 1], a[p + 2]) >= 250 ? 0 : m;
    }
  }
  g.putImageData(d, 0, 0);
  return c;
}

/** A cropped picture (a:srcRect): Leafer paints whole sources only, so the
 *  sub-region renders through an offscreen canvas copy. Mirrors flip the
 *  cropped result (the xfrm flip applies to the blip, post-crop). Shared by
 *  drawing members and inline picture atoms. */
export function addCroppedImage(
  tree: IGroup,
  src: string,
  crop: { left: number; top: number; right: number; bottom: number },
  x: number,
  y: number,
  width: number,
  height: number,
  ctx: PaintContext,
  flipH?: boolean,
  flipV?: boolean,
): void {
  const fingerprint = `crop|${src}|${crop.left},${crop.top},${crop.right},${crop.bottom}`;
  addDerivedImage(
    tree,
    fingerprint,
    (deliver) => {
      const el = new Image();
      el.onload = () => {
        const sx = Math.round(crop.left * el.naturalWidth);
        const sy = Math.round(crop.top * el.naturalHeight);
        const sw = Math.max(1, el.naturalWidth - sx - Math.round(crop.right * el.naturalWidth));
        const sh = Math.max(1, el.naturalHeight - sy - Math.round(crop.bottom * el.naturalHeight));
        const canvas = document.createElement("canvas");
        canvas.width = sw;
        canvas.height = sh;
        canvas.getContext("2d")?.drawImage(el, sx, sy, sw, sh, 0, 0, sw, sh);
        deliver(canvas.toDataURL("image/png"));
      };
      el.onerror = () => deliver(undefined);
      el.src = src;
    },
    x,
    y,
    width,
    height,
    ctx,
    flipH,
    flipV,
  );
}
