/**
 * Shared utility functions for the export-docx package
 */

/**
 * Extract image type from URL or base64 data
 */
export function getImageTypeFromSrc(
  src: string,
): "png" | "jpeg" | "gif" | "bmp" | "tiff" {
  if (src.startsWith("data:")) {
    const match = src.match(/data:image\/(\w+);/);
    if (match) {
      const type = match[1].toLowerCase();
      switch (type) {
        case "jpg":
        case "jpeg":
          return "jpeg";
        case "png":
          return "png";
        case "gif":
          return "gif";
        case "bmp":
          return "bmp";
        case "tiff":
          return "tiff";
        default:
          return "png";
      }
    }
  } else {
    const extension = src.split(".").pop()?.toLowerCase();
    switch (extension) {
      case "jpg":
      case "jpeg":
        return "jpeg";
      case "png":
        return "png";
      case "gif":
        return "gif";
      case "bmp":
        return "bmp";
      case "tiff":
        return "tiff";
      default:
        return "png";
    }
  }
  return "png";
}

/**
 * Create floating options for full-width images
 */
export function createFloatingOptions() {
  return {
    horizontalPosition: {
      relative: "page",
      align: "center",
    },
    verticalPosition: {
      relative: "page",
      align: "top",
    },
    lockAnchor: true,
    behindDocument: false,
    inFrontOfText: false,
  };
}
