import { convertCssLengthToPixels, convertPixelsToTwip, TEXT_ALIGN_MAP } from "@docen/utils";
import type { ParagraphNode } from "@docen/extensions/types";

/**
 * Apply paragraph style attributes to options
 */

export const applyParagraphStyleAttributes = <T extends Record<string, unknown>>(
  options: T,
  attrs?: ParagraphNode["attrs"],
): T => {
  if (!attrs) return options;

  let result = { ...options };

  if (attrs.indentLeft || attrs.indentRight || attrs.indentFirstLine) {
    result = {
      ...result,
      indent: {
        ...(attrs.indentLeft && {
          left: convertPixelsToTwip(convertCssLengthToPixels(attrs.indentLeft)),
        }),
        ...(attrs.indentRight && {
          right: convertPixelsToTwip(convertCssLengthToPixels(attrs.indentRight)),
        }),
        ...(attrs.indentFirstLine && {
          firstLine: convertPixelsToTwip(convertCssLengthToPixels(attrs.indentFirstLine)),
        }),
      },
    };
  }

  if (attrs.spacingBefore || attrs.spacingAfter) {
    result = {
      ...result,
      spacing: {
        ...(attrs.spacingBefore && {
          before: convertPixelsToTwip(convertCssLengthToPixels(attrs.spacingBefore)),
        }),
        ...(attrs.spacingAfter && {
          after: convertPixelsToTwip(convertCssLengthToPixels(attrs.spacingAfter)),
        }),
      },
    };
  }

  if (attrs.textAlign) {
    result = {
      ...result,
      alignment:
        TEXT_ALIGN_MAP.tiptapToDocx[attrs.textAlign as keyof typeof TEXT_ALIGN_MAP.tiptapToDocx],
    };
  }

  return result;
};
