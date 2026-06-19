import { encodeBase64 } from "@office-open/core";
import { imageMeta } from "image-meta";

import type { JSONContent } from "../core";

// ── Types ──

/**
 * A prepare step that transforms Tiptap JSON in place (e.g. fetch external resources).
 */
export type PrepareStep = (json: JSONContent) => Promise<void>;

/**
 * Fetch handler for external image URLs.
 *
 * Receives the URL, returns the image binary data.
 * Override to customize fetching (proxy, auth, caching, etc.).
 */
export type ImageFetchHandler = (url: string) => Promise<Uint8Array>;

// ── Built-in steps ──

/**
 * Default fetch handler using the global `fetch` API (Node 18+ and browsers).
 */
export async function fetchImageHandler(url: string): Promise<Uint8Array> {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to fetch image from ${url}: ${response.status} ${response.statusText}`);
  }
  const buffer = await response.arrayBuffer();
  return new Uint8Array(buffer);
}

/**
 * Create a prepare step that fetches external image URLs and converts them to data URLs.
 *
 * @param handler - Custom fetch handler (defaults to `fetchImageHandler`)
 *
 * @example
 * ```ts
 * // Use with defaults
 * await prepareDocument(json);
 *
 * // Custom handler
 * await prepareDocument(json, [prepareImages(myHandler)]);
 * ```
 */
export function prepareImages(handler: ImageFetchHandler = fetchImageHandler): PrepareStep {
  return async (json: JSONContent) => {
    await walkImages(json, handler);
  };
}

/**
 * Create a prepare step that fills in missing image dimensions by reading image
 * metadata. Intended to run after {@link prepareImages} so HTTP sources are
 * already embedded as data URLs.
 *
 * Only images lacking `width` or `height` are probed — the DOCX round-trip path
 * already carries `transformation.width/height`, so this mainly serves images
 * entering via HTML/Markdown (which carry no intrinsic size). Images whose
 * metadata can't be read are left untouched; `renderDocx` falls back to 600×400
 * (see extensions/image.ts).
 */
export function prepareImageSizes(): PrepareStep {
  return async (json: JSONContent) => {
    await walkImageSizes(json);
  };
}

// ── Pipeline ──

/** Built-in prepare steps, run when no custom steps are provided. */
const DEFAULT_STEPS: readonly PrepareStep[] = [prepareImages(), prepareImageSizes()];

/**
 * Run prepare steps on a Tiptap JSON document before compilation.
 *
 * Each step receives the JSON and may mutate it in place (e.g. replace external
 * URLs with embedded data). Steps run sequentially in order.
 *
 * Defaults to `[prepareImages(), prepareImageSizes()]` when no steps are provided.
 *
 * @example
 * ```ts
 * const json = parseHTML(html);
 * await prepareDocument(json);             // default: fetch images
 * const docOpts = compileDocument(json);
 * ```
 */
export async function prepareDocument(
  json: JSONContent,
  steps: readonly PrepareStep[] = DEFAULT_STEPS,
): Promise<void> {
  for (const step of steps) {
    await step(json);
  }
}

// ── Internals ──

async function toDataUrl(src: string, handler: ImageFetchHandler): Promise<string> {
  const data = await handler(src);
  const ext = src.split(".").pop()?.toLowerCase() ?? "png";
  const typeMap: Record<string, string> = {
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    png: "image/png",
    gif: "image/gif",
    bmp: "image/bmp",
    svg: "image/svg+xml",
    webp: "image/webp",
  };
  const mime = typeMap[ext] ?? "image/png";

  return `data:${mime};base64,${encodeBase64(data)}`;
}

async function walkImages(node: JSONContent, handler: ImageFetchHandler): Promise<void> {
  if (node.type === "image" && node.attrs) {
    const src = node.attrs.src as string | undefined;
    if (src && (src.startsWith("http://") || src.startsWith("https://"))) {
      try {
        node.attrs.src = await toDataUrl(src, handler);
      } catch (error) {
        console.warn(
          `Failed to fetch image: ${src}`,
          error instanceof Error ? error.message : error,
        );
      }
    }
  }

  const tasks: Promise<void>[] = [];
  for (const child of node.content ?? []) {
    tasks.push(walkImages(child, handler));
  }
  await Promise.all(tasks);
}

async function walkImageSizes(node: JSONContent): Promise<void> {
  if (node.type === "image" && node.attrs) {
    const attrs = node.attrs;
    const needsWidth = attrs.width == null;
    const needsHeight = attrs.height == null;
    if (needsWidth || needsHeight) {
      const bytes = decodeDataUrl(attrs.src as string | undefined);
      if (bytes) {
        try {
          const meta = imageMeta(bytes);
          if (needsWidth && typeof meta.width === "number") attrs.width = meta.width;
          if (needsHeight && typeof meta.height === "number") attrs.height = meta.height;
        } catch {
          // Unreadable or unsupported image — leave attrs; renderDocx falls back.
        }
      }
    }
  }

  const tasks: Promise<void>[] = [];
  for (const child of node.content ?? []) {
    tasks.push(walkImageSizes(child));
  }
  await Promise.all(tasks);
}

/** Decode a `data:image/...;base64,...` URL to bytes; null if not a data URL. */
function decodeDataUrl(src: string | undefined): Uint8Array | null {
  if (!src) return null;
  const match = src.match(/^data:image\/[\w.+-]+;base64,(.+)$/);
  if (!match) return null;
  const binary = atob(match[1]);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}
