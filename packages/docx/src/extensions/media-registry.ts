/**
 * Byte registry behind the blob: srcs carried by oversized images.
 *
 * resolveImage hands media above the inline limit a `blob:` URL instead of a
 * base64 data URL: encoding megabytes freezes the main thread for seconds on
 * media-heavy documents, and the base64 string outlives the parse at 1.33×
 * the bytes. The Blob behind the URL is zero-copy and paints natively, and
 * the registry lets renderDocx recover the embedded bytes synchronously on
 * save — no fetch, no re-decode.
 *
 * Entries live for the page's lifetime: the undo history and the clipboard
 * pane hold attrs copies whose srcs must stay resolvable, so there is no
 * safe release point. They retain raw bytes at 1.0× — strictly less than
 * the data URLs they replace.
 */

interface MediaEntry {
  /** Embedded bytes (shared with the Blob — callers must not mutate). */
  bytes: Uint8Array;
  /** office-open image type ("png"/"jpg"/…), not a MIME string. */
  type: string;
}

const registry = new Map<string, MediaEntry>();

/** Media at or below this size keeps the self-contained data URL — small
 * images are the ones that survive clipboard and markdown round-trips, and
 * encoding them is free. Above it, portability loses to the freeze. */
export const MEDIA_INLINE_LIMIT = 256 * 1024;

/** True when a src resolves through this registry (as opposed to data:/http:
 * URLs). A foreign blob: URL (pasted from another page) matches the prefix
 * but not the registry — mediaBytesOf returns undefined for it. */
export function isRegisteredMediaSrc(src: string): boolean {
  return src.startsWith("blob:");
}

/** Store `bytes` behind a fresh object URL. Node 16.7+ provides
 * URL.createObjectURL for Blobs, so headless round-trips register too. */
export function registerMediaBlob(bytes: Uint8Array, type: string): string {
  // The view respects byteOffset/length; the cast only bridges the DOM lib's
  // narrower ArrayBuffer-backing type (TS 5.7 typed arrays).
  const url = URL.createObjectURL(
    new Blob([bytes as unknown as BlobPart], { type: `image/${type}` }),
  );
  registry.set(url, { bytes, type });
  return url;
}

/** The entry behind a registered src, or undefined for anything else. */
export function mediaBytesOf(src: string): MediaEntry | undefined {
  return registry.get(src);
}
