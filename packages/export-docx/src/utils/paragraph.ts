import { convertCssLengthToPixels, convertPixelsToTwip } from "./conversion";
import type { ParagraphNode } from "@docen/tiptap-extensions/types";

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
    const ALIGNMENT_MAP: Record<string, "left" | "right" | "center" | "both"> = {
      left: "left",
      right: "right",
      center: "center",
      justify: "both",
    } as const;
    result = {
      ...result,
      alignment: ALIGNMENT_MAP[attrs.textAlign],
    };
  }

  return result;
};
