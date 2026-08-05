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
 * Attrs mirror RunStylePropertiesOptions. bold/italic/strike/doubleStrike are
 * three-state booleans (true/false/null) carried here for lossless round-trip
 * and CSS cascade; the dedicated Bold/Italic/Strike marks only surface the
 * "true" state for editor interaction (<strong>/<em>/<s>). subScript/
 * superScript (OOXML vertAlign enum, no false state) stay on dedicated marks.
 * DOCX round-trip is near-identity: renderDocx/parseDocx pass attrs through;
 * CSS conversion happens only in attribute-level renderHTML/parseHTML.
 */

/** Structural/semantic keys expressed elsewhere (run children/text, style name). */
const SKIP_KEYS = new Set([
  "children",
  "text",
  "style",
  "break",
  // subScript/superScript are OOXML vertAlign enums with no "false" state, so
  // they stay on dedicated marks. bold/italic/strike/doubleStrike are
  // three-state booleans (true/false/null) — they ride NATIVE_RUN_ATTRS so the
  // "false" state (e.g. <w:b w:val="0"/> cancelling an inherited bold) both
  // round-trips and feeds the CSS cascade; their dedicated marks only surface
  // "true" for editor interaction.
  "subScript",
  "superScript",
]);

/** Scalar OOXML run properties with no CSS equivalent — stored verbatim via
 *  attrNative (default null, renderDocx/parseDocx pass through). Listed once and
 *  spread into TextStyle's attribute set so adding a run prop is a one-line edit. */
const NATIVE_RUN_ATTRS = [
  // Three-state booleans (true/false/null) that have no dedicated mark — ride
  // TextStyle for round-trip + stylesToCss cascade. bold/italic are defined
  // separately below with attribute-level renderHTML so their "false" state
  // cancels an inherited bold/italic on the run (CSS cascade).
  "strike",
  "underline",
  "emphasisMark",
  "highlight",
  "smallCaps",
  "allCaps",
  "kern",
  "position",
  "effect",
  "noProof",
  "sizeComplexScript",
  "highlightComplexScript",
  "boldComplexScript",
  "italicComplexScript",
  "doubleStrike",
  "emboss",
  "imprint",
  "revision",
  "language",
  "border",
  "snapToGrid",
  "vanish",
  "specVanish",
  "scale",
  "math",
  "outline",
  "shadow",
  "webHidden",
  "fitText",
  "complexScript",
  "eastAsianLayout",
  "contentPartRId",
] as const;

const nativeRunAttrs = (): Record<string, ReturnType<typeof attrNative>> =>
  Object.fromEntries(NATIVE_RUN_ATTRS.map((key) => [key, attrNative()]));

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

      // bold/italic are three-state. The dedicated Bold/Italic marks surface
      // "true" as <strong>/<em>; TextStyle carries the full state for round-trip,
      // and "false" emits an inline normal so a run cancelling an inherited
      // bold/italic (e.g. <w:b w:val="0"/> inside a bold style) renders un-bold
      // via CSS cascade (inline overrides the stylesToCss stylesheet). "true"
      // emits nothing here to avoid duplicating <strong>/<em>.
      bold: {
        default: null,
        renderHTML: (attrs: Record<string, unknown>) =>
          attrs.bold === false ? { style: "font-weight:normal" } : {},
      },
      italic: {
        default: null,
        renderHTML: (attrs: Record<string, unknown>) =>
          attrs.italic === false ? { style: "font-style:normal" } : {},
      },

      ...nativeRunAttrs(),
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
