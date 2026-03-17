import { convertCssLengthToPixels, convertPixelsToTwip, TEXT_ALIGN_MAP } from "@docen/utils";
import type { ParagraphNode } from "@docen/extensions/types";
import { convertBorder, convertShading } from "./conversion";

/**
 * Apply paragraph style attributes to options
 */

export const applyParagraphStyleAttributes = <T extends Record<string, unknown>>(
  options: T,
  attrs?: ParagraphNode["attrs"],
): T => {
  if (!attrs) return options;

  // Build result incrementally with single object creation
  const result: Record<string, unknown> = { ...options };

  // Handle indentation (single object creation)
  if (attrs.indentLeft || attrs.indentRight || attrs.indentFirstLine) {
    result.indent = {
      ...(attrs.indentLeft && {
        left: convertPixelsToTwip(convertCssLengthToPixels(attrs.indentLeft)),
      }),
      ...(attrs.indentRight && {
        right: convertPixelsToTwip(convertCssLengthToPixels(attrs.indentRight)),
      }),
      ...(attrs.indentFirstLine && {
        firstLine: convertPixelsToTwip(convertCssLengthToPixels(attrs.indentFirstLine)),
      }),
    };
  }

  // Handle spacing (single object creation)
  if (attrs.spacingBefore || attrs.spacingAfter) {
    result.spacing = {
      ...(attrs.spacingBefore && {
        before: convertPixelsToTwip(convertCssLengthToPixels(attrs.spacingBefore)),
      }),
      ...(attrs.spacingAfter && {
        after: convertPixelsToTwip(convertCssLengthToPixels(attrs.spacingAfter)),
      }),
    };
  }

  // Handle alignment (direct assignment)
  if (attrs.textAlign) {
    result.alignment =
      TEXT_ALIGN_MAP.tiptapToDocx[attrs.textAlign as keyof typeof TEXT_ALIGN_MAP.tiptapToDocx];
  }

  // Apply shading (background color)
  if (attrs.shading) {
    result.shading = convertShading(attrs.shading);
  }

  // Apply borders (single object creation)
  if (attrs.borderTop || attrs.borderBottom || attrs.borderLeft || attrs.borderRight) {
    result.border = {
      ...(attrs.borderTop && { top: convertBorder(attrs.borderTop) }),
      ...(attrs.borderBottom && { bottom: convertBorder(attrs.borderBottom) }),
      ...(attrs.borderLeft && { left: convertBorder(attrs.borderLeft) }),
      ...(attrs.borderRight && { right: convertBorder(attrs.borderRight) }),
    };
  }

  return result as T;
};
