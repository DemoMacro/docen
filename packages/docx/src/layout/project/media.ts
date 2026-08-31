// Picture media: bytes → data-URL encoding (the renderer's passthrough src)
// and the metafile replay caches. Project reruns on every editor
// transaction, so the caches key on bytes identity or a fingerprint head and
// megabyte WMFs are not re-scanned each pass.

import type { LayoutDrawingMember } from "@docen/layout";
import { emfPlusMembers, wmfMembers, wmfDibFallback, type SourceCrop } from "leafer-x-metafile";

import { toLayoutMembers } from "../metafile-members";

// ── picture media (renderer passthrough) ──

const MIME_BY_TYPE: Record<string, string> = {
  jpg: "image/jpeg",
  png: "image/png",
  gif: "image/gif",
  bmp: "image/bmp",
  tif: "image/tiff",
  ico: "image/x-icon",
  svg: "image/svg+xml",
};

/** WPS-authored svgBlip art names its gradients with NCName-invalid ids
 *  (`wps{guid}@#c1@#c2`): a strict CSS parser cannot resolve those paint
 *  references. The gradient def itself is well-formed (real stop children),
 *  and MS Office renders these as true gradients — its PDF export shows the
 *  full ramp — so rename each invalid id (def + url() refs) to a valid one
 *  instead of degrading the art; well-formed ids pass through untouched. */
function sanitizeSvgGradientIds(svg: string): string {
  const renames = new Map<string, string>();
  let i = 0;
  for (const m of svg.matchAll(/<linearGradient\b[^>]*\bid="([^"]+)"/g)) {
    const id = m[1]!;
    if (!/[{}@#]/.test(id) || renames.has(id)) continue;
    renames.set(id, `wpsGradient${i++}`);
  }
  for (const [oldId, newId] of renames) {
    svg = svg.split(`id="${oldId}"`).join(`id="${newId}"`);
    svg = svg.split(`url(#${oldId})`).join(`url(#${newId})`);
  }
  return svg;
}

/** Bytes → base64 (btoa is universal: browsers and Node ≥ 16). */
function base64Of(bytes: Uint8Array): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return btoa(bin);
}

// Encoded src cache, keyed by the bytes object itself: renderDocx hands out
// the same cached array every pass (see decodedBytesOf), so identity is
// stable across transactions and the megabyte btoa runs once per image,
// not once per keystroke. The bytes stay alive with their document node,
// which keeps the entry alive with them — no eviction needed.
const encodedSrcs = new WeakMap<Uint8Array, string>();

/** Raster bytes → data URL, memoized per bytes identity. SVG bytes first get
 *  their WPS-invalid gradient ids renamed (browser loaders cannot resolve
 *  the originals). */
function encodedDataUrl(mime: string, data: Uint8Array): string {
  const hit = encodedSrcs.get(data);
  if (hit) return hit;
  const body =
    mime === "image/svg+xml"
      ? new TextEncoder().encode(sanitizeSvgGradientIds(new TextDecoder().decode(data)))
      : data;
  const url = `data:${mime};base64,${base64Of(body)}`;
  encodedSrcs.set(data, url);
  return url;
}

/** PictureOptions (type + data) → a data URL the painter can load. Absent
 *  bytes (linked-only) yields undefined — the renderer draws an empty frame.
 *  Metafile types with no raster MIME (emf/wmf — browsers have no GDI
 *  rasterizer) fall back to the embedded DIB (see wmf-dib.ts). */
export function pictureSrc(pic: { type?: unknown; data?: unknown }): string | undefined {
  if (typeof pic.type !== "string") return undefined;
  const mime = MIME_BY_TYPE[pic.type];
  const { data } = pic;
  if (!mime) {
    if (typeof data === "string" || data instanceof Uint8Array) {
      return metafileFallback(pic.type, data);
    }
    return undefined;
  }
  if (typeof data === "string") {
    return data.startsWith("data:") ? data : `data:${mime};base64,${data}`;
  }
  if (data instanceof Uint8Array) {
    return encodedDataUrl(mime, data);
  }
  return undefined;
}

/** Metafile caches: project reruns on every editor transaction, and
 *  re-scanning megabyte WMFs each pass is pure waste. Direct-API callers
 *  handing out the same bytes object every pass hit WeakMaps keyed by those
 *  bytes — entry lifetime rides the picture's own. The editor's compiled
 *  options rebuild the picture objects each pass (identity dies in
 *  compileDocument) and its attrs carry data-URL strings, so the editor's
 *  hot path is the string memo: its bound must exceed a real document's
 *  metafile count or the working set thrashes (the 112-media corpus doc
 *  against a 32-slot memo re-replayed ~80 files every relayout). Replays
 *  are lightweight structure, so the bound is generous; the DIB backdrop
 *  memo stays small on purpose — its values are multi-megabyte BMP data
 *  URLs and mask-layer files are rare. */
function memoByFingerprint<V>(limit: number): (key: string, make: () => V) => V {
  const map = new Map<string, V>();
  return (key, make) => {
    if (map.has(key)) return map.get(key) as V;
    const value = make();
    map.set(key, value);
    if (map.size > limit) {
      const oldest = map.keys().next().value;
      if (oldest !== undefined) map.delete(oldest);
    }
    return value;
  };
}

const dibFallbackByIdentity = new WeakMap<Uint8Array, string | undefined>();
const dibFallbackOfString = memoByFingerprint<string | undefined>(16);
const wmfMembersByIdentity = new WeakMap<
  Uint8Array,
  Map<string, LayoutDrawingMember[] | undefined>
>();
const wmfMembersOfString = memoByFingerprint<LayoutDrawingMember[] | undefined>(192);

/** Cache fingerprint head for string data: the base64 payload prefix after a
 *  data-URL header (the header itself is constant, zero distinguishing
 *  entropy) — or the raw prefix when no header is present. */
function fingerprintHead(data: string): string {
  const start = data.startsWith("data:") ? data.indexOf(",") + 1 : 0;
  return data.slice(start, start + 24);
}

export function metafileFallback(type: string, data: string | Uint8Array): string | undefined {
  if (typeof data === "string") {
    return dibFallbackOfString(`${type}:${data.length}:${fingerprintHead(data)}`, () => {
      const bytes = base64ToBytes(data);
      return bytes ? wmfDibFallback(bytes) : undefined;
    });
  }
  if (dibFallbackByIdentity.has(data)) return dibFallbackByIdentity.get(data);
  const value = wmfDibFallback(data);
  dibFallbackByIdentity.set(data, value);
  return value;
}

/** A metafile picture's vector replay (wmf.ts), cached per bytes+box (the
 *  members scale with the box, so the size rides the key). Raster types
 *  (a real MIME) return undefined — they paint through `src` directly.
 *  Mask-layer files (SRCPAINT/SRCAND blts, no SRCCOPY) replay text and
 *  strokes but not their photo — the flat DIB backdrop carries it under
 *  the members (see wmf-dib.ts for the extraction).
 *
 * Dual-mode files carry their real art as an embedded GDI+ stream
 * (emf-plus.ts); that replay wins when present and already includes every
 * raster it draws, so no DIB backdrop is layered beneath it. */
export function metafileMembers(
  pic: { type?: unknown; data?: unknown },
  boxW: number,
  boxH: number,
  crop?: SourceCrop,
): LayoutDrawingMember[] | undefined {
  if (typeof pic.type !== "string" || MIME_BY_TYPE[pic.type]) return undefined;
  const { data } = pic;
  if (typeof data !== "string" && !(data instanceof Uint8Array)) return undefined;
  const boxKey = `${Math.round(boxW)}x${Math.round(boxH)}${
    crop ? `:${crop.left},${crop.top},${crop.right},${crop.bottom}` : ""
  }`;
  if (data instanceof Uint8Array) {
    let byBox = wmfMembersByIdentity.get(data);
    if (!byBox) {
      byBox = new Map();
      wmfMembersByIdentity.set(data, byBox);
    }
    if (byBox.has(boxKey)) return byBox.get(boxKey);
    const value = replayMetafile(pic.type as string, data, boxW, boxH, crop);
    byBox.set(boxKey, value);
    return value;
  }
  const key = `${pic.type}:${data.length}:${fingerprintHead(data)}:${boxKey}`;
  return wmfMembersOfString(key, () => {
    const bytes = base64ToBytes(data);
    return bytes ? replayMetafile(pic.type as string, bytes, boxW, boxH, crop) : undefined;
  });
}

/** EMF+ stream first, then the WMF record replay; a replay without any
 *  raster member gets the flat DIB backdrop layered beneath it. */
function replayMetafile(
  type: string,
  bytes: Uint8Array,
  boxW: number,
  boxH: number,
  crop?: SourceCrop,
): LayoutDrawingMember[] | undefined {
  const plus = emfPlusMembers(bytes, boxW, boxH, crop);
  if (plus) return toLayoutMembers(plus);
  const replay = wmfMembers(bytes, boxW, boxH, crop);
  if (!replay) return undefined;
  if (!replay.some((m) => m.kind === "picture")) {
    const backdrop = metafileFallback(type, bytes);
    if (backdrop) {
      replay.unshift({ kind: "picture", x: 0, y: 0, width: boxW, height: boxH, src: backdrop });
    }
  }
  return toLayoutMembers(replay);
}

function base64ToBytes(b64: string): Uint8Array | undefined {
  try {
    // Editor pictures carry full data URLs (`data:…;base64,…`), not bare
    // base64 — strip the prefix or atob chokes on the header characters.
    const comma = b64.indexOf(",");
    const payload = b64.startsWith("data:") && comma >= 0 ? b64.slice(comma + 1) : b64;
    const bin = atob(payload);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  } catch {
    return undefined;
  }
}
