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

// ── Pipeline ──

/** Built-in prepare steps, run when no custom steps are provided. */
const DEFAULT_STEPS: readonly PrepareStep[] = [prepareImages()];

/**
 * Run prepare steps on a Tiptap JSON document before compilation.
 *
 * Each step receives the JSON and may mutate it in place (e.g. replace external
 * URLs with embedded data). Steps run sequentially in order.
 *
 * Defaults to `[prepareImages()]` when no steps are provided.
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

  const base64 =
    typeof btoa !== "undefined"
      ? btoa(String.fromCharCode(...data))
      : Buffer.from(data).toString("base64");

  return `data:${mime};base64,${base64}`;
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
