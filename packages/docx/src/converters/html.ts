import { encodeBase64 } from "@office-open/core";
import {
  generateHTML as generateTiptapHTML,
  generateJSON as generateTiptapJSON,
} from "@tiptap/html";
import type { ParseOptions } from "@tiptap/pm/model";

import type { JSONContent, Extensions } from "../core";
import { docxExtensions } from "../core";

const defaultExtensions: Extensions = docxExtensions;

interface BackgroundMediaItem {
  data?: Uint8Array | Record<string, number>;
  type?: string;
}

interface BackgroundLike {
  color?: string;
  rawXml?: string;
  image?: { data?: BackgroundMediaItem };
  rawMedia?: BackgroundMediaItem[];
}

/**
 * Normalize a media payload to Uint8Array — JSON round-trips byte arrays as
 * plain objects ({0:byte,...}), so rebuild before base64-encoding.
 */
function toBytes(d: Uint8Array | Record<string, number> | undefined): Uint8Array | null {
  if (!d) return null;
  if (d instanceof Uint8Array) return d;
  return new Uint8Array(Object.values(d));
}

/**
 * Build a CSS background style from DocumentBackgroundOptions. VML/rawXml page
 * backgrounds (base fill + pattern image) reduce to a solid color (the w:color
 * base fill) plus a tiled image (rawMedia) — OOXML pattern semantics can't be
 * expressed in CSS, but base color + tile gives a visually close result.
 */
function extractBackgroundStyle(bg: unknown): string | undefined {
  if (!bg || typeof bg !== "object") return undefined;
  const b = bg as BackgroundLike;
  let color: string | undefined;
  if (b.color) color = b.color;
  else if (b.rawXml) {
    const m = b.rawXml.match(/w:color="([0-9A-Fa-f]{6})"/);
    if (m) color = m[1];
  }
  const media = b.image?.data ?? b.rawMedia?.[0];
  const bytes = toBytes(media?.data);
  const styles: string[] = [];
  if (color) styles.push(`background-color:#${color}`);
  if (bytes) {
    const type = media?.type ?? "png";
    styles.push(
      `background-image:url(data:image/${type};base64,${encodeBase64(bytes)})`,
      "background-repeat:repeat",
    );
  }
  return styles.length ? styles.join(";") : undefined;
}

/**
 * Parse HTML string to Tiptap JSON. Reads back the page background base color
 * from a generateHTML wrapper div (pattern image is lost — OOXML patterns have
 * no HTML equivalent; the DOCX round-trip preserves the full background).
 */
export function parseHTML(
  html: string,
  extensions?: Extensions,
  options?: ParseOptions,
): JSONContent {
  const doc = generateTiptapJSON(html, extensions ?? defaultExtensions, options);
  const m = html.match(/<div style="background-color:#([0-9A-Fa-f]{6})/);
  if (m) {
    if (!doc.attrs) doc.attrs = {};
    (doc.attrs as Record<string, unknown>).background = { color: m[1] };
  }
  return doc;
}

/**
 * Generate HTML string from Tiptap JSON. Wraps the content in a div carrying
 * the page background (color + pattern image) so the document background is
 * visible — Tiptap's DOMSerializer only serializes content nodes, not doc attrs.
 */
export function generateHTML(doc: JSONContent, extensions?: Extensions): string {
  const html = generateTiptapHTML(doc, extensions ?? defaultExtensions);
  const style = extractBackgroundStyle(
    (doc.attrs as Record<string, unknown> | undefined)?.background,
  );
  if (style) return `<div style="${style}">${html}</div>`;
  return html;
}
