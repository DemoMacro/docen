import type { JSONContent } from "@docen/docx";
import type { Editor } from "@docen/docx/core";

import {
  type WatermarkPictureSpec,
  type WatermarkTextSpec,
  WATERMARK_PRESETS,
  customTextWatermarkPara,
  isWatermarkNode,
  pictureWatermarkPara,
  probeImageSize,
  stampHeaderSlots,
  watermarkPara,
} from "../watermark";

/** The design commands' view of the host — resolved per call so the
 *  controller can be built before a document opens. */
export interface DesignHost {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The story bridge — the watermark dialog's commit hands focus back. */
  bridge(): { focus(): void } | undefined;
  /** The host element — the shadow-DOM root for the watermark dialog. */
  element(): HTMLElement;
}

/**
 * The Design tab's document-level stamps, split out of the host element: the
 * paragraph-spacing docDefaults presets, the page background, and the
 * watermark gallery/dialog (Word's watermark IS a behind-document,
 * page-centered shape stamped into every section's header carrier).
 */
export class DesignCommands {
  constructor(private readonly host: DesignHost) {}

  /** Design → Paragraph Spacing presets — stamp the styles' docDefaults
   *  paragraph spacing (styles.default.document.paragraph.spacing), the
   *  document-level default every paragraph without explicit spacing
   *  inherits. Word's preset values: default restores the factory 8pt-after /
   *  1.08-line spacing; the named presets are single-spaced with the after
   *  gap shrinking (none 0pt → compact 2pt → narrow 6pt → wide 16pt). */
  setParagraphSpacing(preset?: string): void {
    const editor = this.host.editor();
    if (!editor) return;
    const spacing =
      preset === "none"
        ? { before: 0, after: 0, line: 240, lineRule: "auto" }
        : preset === "compact"
          ? { after: 40, line: 240, lineRule: "auto" }
          : preset === "narrow"
            ? { after: 120, line: 240, lineRule: "auto" }
            : preset === "wide"
              ? { after: 320, line: 240, lineRule: "auto" }
              : preset === "default"
                ? { after: 160, line: 259, lineRule: "auto" }
                : null;
    if (!spacing) return;
    const styles = { ...((editor.state.doc.attrs.styles ?? {}) as Record<string, unknown>) };
    const defaults = { ...((styles.default ?? {}) as Record<string, unknown>) };
    const documentDefaults = { ...((defaults.document ?? {}) as Record<string, unknown>) };
    documentDefaults.paragraph = {
      ...((documentDefaults.paragraph ?? {}) as Record<string, unknown>),
      spacing,
    };
    defaults.document = documentDefaults;
    styles.default = defaults;
    editor.view.dispatch(editor.state.tr.setDocAttribute("styles", styles));
  }

  /** Design → Page Color: write the doc-level page background
   *  (doc.attrs.background → w:background on export; the stage paints it as
   *  the page frame color). "none" clears it (Word's No Color); a bare hex is
   *  the standard/custom swatch path; a theme-semantic pick carries its
   *  themeColor/tint/shade through so Word re-resolves on theme change. */
  setPageColor(
    value?: string | { themeColor: string; val: string; themeTint?: string; themeShade?: string },
  ): void {
    const editor = this.host.editor();
    if (!editor) return;
    const background =
      value == null || value === "none"
        ? null
        : typeof value === "string"
          ? { color: value }
          : {
              color: value.val,
              themeColor: value.themeColor,
              ...(value.themeTint ? { themeTint: value.themeTint } : {}),
              ...(value.themeShade ? { themeShade: value.themeShade } : {}),
            };
    editor.view.dispatch(editor.state.tr.setDocAttribute("background", background));
  }

  /** Design → Watermark presets (Word's gallery): one diagonal silver text
   *  shape stamped into every header slot — Word's watermark IS a
   *  behind-document, page-centered shape anchored in the header, so it
   *  repeats on every page of the slot. The shape carries Word's watermark
   *  name ("WordPictureWatermark"), which is also how Remove finds it. */
  setWatermark(preset?: string): void {
    const spec = preset && preset !== "remove" ? WATERMARK_PRESETS[preset] : undefined;
    this.#stampWatermark(spec ? watermarkPara(spec) : null);
  }

  /** Strip any existing watermark from every section's header slots, then
   *  append the given stamp paragraph (null removes). Word's watermark rides
   *  linked headers and reads on every page, so every section's carrier is
   *  stamped — earlier sections' slots live on their closing sectPr
   *  paragraphs, the final section's on the doc node. Shared by the gallery
   *  presets and the custom dialog's text/picture stamps. */
  #stampWatermark(para: JSONContent | null): void {
    const editor = this.host.editor();
    if (!editor) return;
    const { doc, tr } = editor.state;
    const stamp = (attrs: Record<string, unknown>): Record<string, unknown> => {
      const stamped = stampHeaderSlots(
        attrs.sectionHeaders as Record<string, JSONContent[] | undefined>,
        para,
      );
      // An all-empty sectionHeaders is the no-headers state (Word's Remove
      // Watermark leaves a blank header behind; an empty attrs object is the
      // cleaner equivalent here and drops the furniture strut).
      const anyContent = Object.values(stamped).some((paras) => paras.length > 0);
      return { ...attrs, sectionHeaders: anyContent ? stamped : {} };
    };
    doc.descendants((node, pos) => {
      if (
        node.type.name === "paragraph" &&
        (node.attrs as { sectionProperties?: unknown }).sectionProperties != null
      ) {
        tr.setNodeMarkup(pos, undefined, stamp(node.attrs as Record<string, unknown>));
      }
      return true;
    });
    tr.setDocAttribute(
      "sectionHeaders",
      stamp(doc.attrs as Record<string, unknown>).sectionHeaders,
    );
    editor.view.dispatch(tr);
  }

  /** The custom watermark dialog's OK (Word's 自定义水印): none clears, the
   *  text spec stamps the text shape, the picture spec probes the natural
   *  size then stamps the floating picture. */
  readonly onWatermarkOk = (
    event: CustomEvent<
      | { kind: "none" }
      | { kind: "text"; spec: WatermarkTextSpec }
      | { kind: "picture"; spec: WatermarkPictureSpec }
      | undefined
    >,
  ): void => {
    const detail = event.detail;
    if (!detail) return;
    if (detail.kind === "none") {
      this.#stampWatermark(null);
    } else if (detail.kind === "text") {
      this.#stampWatermark(customTextWatermarkPara(detail.spec));
    } else {
      void probeImageSize(detail.spec.src).then((natural) => {
        this.#stampWatermark(pictureWatermarkPara(detail.spec, natural));
      });
    }
    this.host.bridge()?.focus();
  };

  /** Design → Watermark → Custom Watermark: open the dialog prefilled from
   *  the current stamp (a text shape's run reads back text/color/size;
   *  a picture stamp selects the picture pane). */
  openWatermarkDialog(): void {
    const dialog = this.host.element().shadowRoot?.querySelector("docen-watermark-dialog") as {
      show(current?: unknown): void;
    } | null;
    const editor = this.host.editor();
    if (!dialog) return;
    if (!editor) {
      dialog.show();
      return;
    }
    // The stamp rides every section's carrier (see #stampWatermark) — read
    // them in document order and prefill from the first stamp found.
    const groups: Array<Record<string, JSONContent[] | undefined>> = [];
    editor.state.doc.descendants((node) => {
      if (
        node.type.name === "paragraph" &&
        (node.attrs as { sectionProperties?: unknown }).sectionProperties != null
      ) {
        groups.push(
          (node.attrs as { sectionHeaders?: Record<string, JSONContent[] | undefined> })
            .sectionHeaders ?? {},
        );
      }
      return true;
    });
    groups.push(
      (editor.state.doc.attrs.sectionHeaders ?? {}) as Record<string, JSONContent[] | undefined>,
    );
    let current: unknown = null;
    for (const headers of groups) {
      current = this.#watermarkSpecOf(headers);
      if (current) break;
    }
    dialog.show(current);
  }

  /** One section's header slots read back as the dialog's prefill — the first
   *  watermark shape/picture in the default slot, null when none. */
  #watermarkSpecOf(headers: Record<string, JSONContent[] | undefined>): unknown {
    for (const para of headers.default ?? []) {
      for (const child of ((para as JSONContent).content ?? []) as JSONContent[]) {
        if (child.type === "wpsShape" && isWatermarkNode(child)) {
          const shape = (child.attrs as { wpsShape?: Record<string, unknown> }).wpsShape ?? {};
          const run = ((child as JSONContent).content?.[0]?.content ?? []) as Array<{
            marks?: Array<{ type: string; attrs: Record<string, unknown> }>;
          }>;
          const style = run[0]?.marks?.find((m) => m.type === "textStyle")?.attrs ?? {};
          return {
            kind: "text",
            text: (child.content?.[0]?.content?.[0] as { text?: string })?.text ?? "",
            font: (style.font as string) ?? null,
            size: (style.size as number) ?? null,
            color: (style.color as string) ?? "C0C0C0",
            diagonal:
              (shape.transformation as { rotation?: number })?.rotation != null &&
              (shape.transformation as { rotation?: number }).rotation! < 0,
            semiTransparent: false,
          };
        }
        if (child.type === "image" && isWatermarkNode(child)) {
          const attrs = child.attrs as { blipEffects?: { luminance?: unknown } };
          return {
            kind: "picture",
            hasImage: true,
            washout: !!attrs.blipEffects?.luminance,
            scale: "auto",
          };
        }
      }
    }
    return null;
  }
}
