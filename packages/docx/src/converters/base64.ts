/**
 * Encode bytes as base64. Node uses Buffer (fast path); browsers fall back to
 * chunked `String.fromCharCode` + `btoa`. The chunking matters: spreading a
 * whole image into `String.fromCharCode(...bytes)` makes every byte a function
 * argument and overflows the call stack on large images.
 *
 * (@office-open/core exposes decodeBase64 but not yet encodeBase64 in the
 * published 0.9.8 — swap to `encodeBase64` from `@office-open/core` once a
 * release ships it.)
 */
export function bytesToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== "undefined") return Buffer.from(bytes).toString("base64");
  const CHUNK = 0x4000; // 16 KiB — well under the argument/stack limit
  let binary = "";
  for (let i = 0; i < bytes.length; i += CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
  }
  return btoa(binary);
}
