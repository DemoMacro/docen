import type { RunOptions } from "@office-open/docx";
import { TextStyle as BaseTextStyle } from "@tiptap/extension-text-style";

import {
  attrNative,
  characterSpacingFromCss,
  characterSpacingToCss,
  normalizeColorToHex,
  resolveFontFamilyCss,
  shadingFromCss,
  shadingToCss,
  sizeFromCss,
  sizeToCss,
} from "./utils";

/**
 * TextStyle mark with office-open attrs.
 *
 * Attrs mirror RunStylePropertiesOptions (bold/italic/strike/subScript/
 * superScript handled by dedicated marks and therefore omitted). DOCX
 * round-trip is near-identity: renderDocx/parseDocx pass attrs through;
 * CSS conversion happens only in attribute-level renderHTML/parseHTML.
 */

/** Structural/semantic keys expressed elsewhere (run children/text, style name). */
const SKIP_KEYS = new Set([
  "children",
  "text",
  "style",
  "break",
  // Expressed by dedicated marks — must not pollute textStyle attrs:
  "bold",
  "italic",
  "strike",
  "doubleStrike",
  "subScript",
  "superScript",
]);

export const TextStyle = BaseTextStyle.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      // rStyle reference (e.g. "InternetLink") — the named character style.
      // renderHTML emits class="docx-char-{styleId}" so the injected document
      // CSS (generated from styles.xml characterStyles) applies. Round-trips
      // via OOXML run `style`.
      styleId: {
        default: null,
        parseHTML: (element: HTMLElement) => {
          const m = (element.getAttribute("class") || "").match(/(?:^|\s)docx-char-(\S+)/);
          return m ? m[1] : null;
        },
        renderHTML: (attributes: Record<string, unknown>) => {
          const id = attributes.styleId as string | null;
          return id ? { class: `docx-char-${id}` } : {};
        },
      },

      // Scalar OOXML run properties with CSS equivalents.
      // Attr values are office-open native (color hex, font name, size in
      // points, shading object); CSS conversion lives in renderHTML/parseHTML.
      color: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          normalizeColorToHex(element.style.color || undefined) ?? null,
        renderHTML: (attributes: Record<string, unknown>) => {
          // "auto"/unset emit no inline color: the text inherits the page
          // default (color: contrast-color(var(--docen-ink-bg)) on .docen-page),
          // so default text inverts against the nearest fill like Word. An
          // explicit hex overrides it.
          if (attributes.color === "auto") return {};
          const hex = normalizeColorToHex(attributes.color as string | undefined);
          return hex ? { style: `color:${hex}` } : {};
        },
      },
      characterSpacing: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          characterSpacingFromCss(element.style.letterSpacing || null),
        renderHTML: (attributes: Record<string, unknown>) => {
          const css = characterSpacingToCss(
            attributes.characterSpacing as number | null | undefined,
          );
          return css ? { style: `letter-spacing:${css}` } : {};
        },
      },
      font: {
        default: null,
        parseHTML: (element: HTMLElement) => element.style.fontFamily || null,
        renderHTML: (attributes: Record<string, unknown>) => {
          const name = resolveFontFamilyCss(attributes.font);
          return name ? { style: `font-family:${name}` } : {};
        },
      },
      rightToLeft: {
        default: null,
        parseHTML: (element: HTMLElement) => (element.dir === "rtl" ? true : null),
        renderHTML: (attributes: Record<string, unknown>) =>
          attributes.rightToLeft ? { style: "direction:rtl" } : {},
      },
      // RunOptions.size is in POINTS (office-open convention); CSS font-size is
      // derived in renderHTML and parsed back in parseHTML.
      size: {
        default: null,
        parseHTML: (element: HTMLElement) => sizeFromCss(element.style.fontSize),
        renderHTML: (attributes: Record<string, unknown>) => {
          const css = sizeToCss(attributes.size as number | null | undefined);
          return css ? { style: `font-size:${css}` } : {};
        },
      },
      // RunOptions.shading (OOXML <w:shd>) ↔ CSS background-color.
      shading: {
        default: null,
        parseHTML: (element: HTMLElement) => shadingFromCss(element.style.backgroundColor),
        renderHTML: (attributes: Record<string, unknown>) => {
          const css = shadingToCss(attributes.shading as { fill?: string } | null | undefined);
          // A run fill flips the ink against it (Word "auto"). Declared on the
          // run's own span so the text inherits the inverted color directly.
          return css ? { style: `background-color:${css};color:contrast-color(${css})` } : {};
        },
      },

      // Scalar OOXML run properties with no CSS equivalent (stored verbatim)
      underline: attrNative(),
      emphasisMark: attrNative(),
      highlight: attrNative(),
      smallCaps: attrNative(),
      allCaps: attrNative(),
      kern: attrNative(),
      position: attrNative(),
      effect: attrNative(),
      noProof: attrNative(),
      sizeComplexScript: attrNative(),
      highlightComplexScript: attrNative(),
      boldComplexScript: attrNative(),
      italicComplexScript: attrNative(),
      doubleStrike: attrNative(),
      emboss: attrNative(),
      imprint: attrNative(),
      revision: attrNative(),
      language: attrNative(),
      border: attrNative(),
      snapToGrid: attrNative(),
      vanish: attrNative(),
      specVanish: attrNative(),
      scale: attrNative(),
      math: attrNative(),
      outline: attrNative(),
      shadow: attrNative(),
      webHidden: attrNative(),
      fitText: attrNative(),
      complexScript: attrNative(),
      eastAsianLayout: attrNative(),
      contentPartRId: attrNative(),
    };
  },

  // Near-identity: attrs mirror RunStylePropertiesOptions. styleId (attr) ↔
  // OOXML run `style` (rStyle). CSS conversion happens only in attribute-level
  // renderHTML/parseHTML above.
  renderDocx: (attrs: Record<string, unknown>): Partial<RunOptions> => {
    const opts: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(attrs)) {
      if (SKIP_KEYS.has(key)) continue;
      if (value === null || value === undefined) continue;
      // styleId (attr) → OOXML run `style` (the rStyle / character-style reference).
      if (key === "styleId") {
        opts.style = value;
        continue;
      }
      opts[key] = value;
    }
    return opts as Partial<RunOptions>;
  },
  parseDocx: (opts: RunOptions): Record<string, unknown> | null => {
    const resolved = typeof opts === "string" ? { text: opts } : opts;
    const attrs: Record<string, unknown> = {};
    // rStyle "CodeChar" belongs to the `code` mark — skip its styleId and the
    // Consolas fallback font so they aren't double-applied on compile.
    if (resolved.style && resolved.style !== "CodeChar") attrs.styleId = resolved.style;
    for (const [key, value] of Object.entries(resolved)) {
      if (SKIP_KEYS.has(key)) continue;
      if (resolved.style === "CodeChar" && key === "font") continue;
      attrs[key] = value ?? null;
    }
    return Object.keys(attrs).length ? attrs : null;
  },
});
