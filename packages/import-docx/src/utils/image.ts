/**
 * Image processing utilities for cross-platform image cropping
 * Supports both browser (native Canvas) and Node.js (@napi-rs/canvas)
 *
 * Cross-platform canvas factory implementation inspired by @unjs/unpdf
 * @see https://github.com/unjs/unpdf
 */

import type { Canvas } from "@napi-rs/canvas";

/**
 * Detect current environment
 */
export const isNode = globalThis.process?.release?.name === "node";
export const isBrowser = typeof window !== "undefined";

/**
 * Handle interop for module default exports (from unpdf)
 */
async function interopDefault<T>(
  m: T | Promise<T>,
): Promise<T extends { default: infer U } ? U : T> {
  const resolved = await m;
  return (resolved as any).default || resolved;
}

interface CanvasFactoryContext {
  canvas?: HTMLCanvasElement | Canvas;
  context?: CanvasRenderingContext2D | import("@napi-rs/canvas").CanvasRenderingContext2D;
}

let resolvedCanvasModule: typeof import("@napi-rs/canvas") | undefined;

/**
 * Base canvas factory for cross-platform canvas creation
 */
class BaseCanvasFactory {
  #enableHWA = false;

  constructor({ enableHWA = false } = {}) {
    this.#enableHWA = enableHWA;
  }

  create(width: number, height: number) {
    const canvas = this._createCanvas(width, height);

    return {
      canvas,
      context: canvas.getContext("2d", {
        willReadFrequently: !this.#enableHWA,
      }),
    };
  }

  reset({ canvas }: CanvasFactoryContext, width: number, height: number) {
    if (!canvas) {
      throw new Error("Canvas is not specified");
    }

    canvas.width = width;
    canvas.height = height;
  }

  destroy(context: CanvasFactoryContext) {
    if (!context.canvas) {
      throw new Error("Canvas is not specified");
    }

    // Zeroing the width and height cause Firefox to release graphics
    // resources immediately, which can greatly reduce memory consumption.
    context.canvas.width = 0;
    context.canvas.height = 0;
    context.canvas = undefined;
    context.context = undefined;
  }

  // eslint-disable-next-line unused-imports/no-unused-vars
  _createCanvas(width: number, height: number): HTMLCanvasElement | Canvas {
    throw new Error("Not implemented");
  }
}

/**
 * Browser canvas factory using native HTMLCanvasElement
 */
export class DOMCanvasFactory extends BaseCanvasFactory {
  _document: Document;

  constructor({ ownerDocument = globalThis.document, enableHWA = false } = {}) {
    super({ enableHWA });
    this._document = ownerDocument;
  }

  _createCanvas(width: number, height: number) {
    const canvas = this._document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    return canvas;
  }
}

/**
 * Node.js canvas factory using @napi-rs/canvas
 */
export class NodeCanvasFactory extends BaseCanvasFactory {
  constructor({ enableHWA = false } = {}) {
    super({ enableHWA });
  }

  _createCanvas(width: number, height: number) {
    if (!resolvedCanvasModule) {
      throw new Error("@napi-rs/canvas module is not resolved");
    }

    return resolvedCanvasModule.createCanvas(width, height);
  }
}

/**
 * Resolve canvas module (from unpdf)
 */
export async function resolveCanvasModule(
  canvasImport: () => Promise<typeof import("@napi-rs/canvas")>,
) {
  resolvedCanvasModule ??= await interopDefault(canvasImport());
}

/**
 * Create appropriate canvas factory for current environment
 *
 * @param canvasImport - Dynamic import function for @napi-rs/canvas (required in Node.js)
 * @returns CanvasFactory instance
 */
export async function createCanvasFactory(
  canvasImport?: () => Promise<typeof import("@napi-rs/canvas")>,
) {
  if (isBrowser) return DOMCanvasFactory;

  if (isNode) {
    if (!canvasImport) {
      throw new Error(
        "In Node.js environment, @napi-rs/canvas is required for image cropping. " +
          "Please provide canvasImport parameter or install it: pnpm add @napi-rs/canvas",
      );
    }

    await resolveCanvasModule(canvasImport);
    return NodeCanvasFactory;
  }

  throw new Error("Unsupported environment for canvas operations");
}

/**
 * Crop rectangle from DOCX a:srcRect attributes
 * Values are in 1/100000 of a percentage (0-100000)
 */
export interface CropRect {
  left?: number;
  top?: number;
  right?: number;
  bottom?: number;
}

/**
 * Crop image options
 */
export interface CropImageOptions {
  /**
   * Dynamic import function for @napi-rs/canvas
   * Required in Node.js environment, ignored in browser
   */
  canvasImport?: () => Promise<typeof import("@napi-rs/canvas")>;

  /**
   * Enable or disable image cropping
   * @default true
   */
  enabled?: boolean;
}

/**
 * Crop image if crop information is provided
 *
 * @param imageData - Original image data as Uint8Array
 * @param crop - Crop rectangle (DOCX format: 0-100000)
 * @param options - Cropping options
 * @returns Cropped image data, or original if no crop or error occurs
 */
export async function cropImageIfNeeded(
  imageData: Uint8Array,
  crop: CropRect | undefined,
  options: CropImageOptions = {},
): Promise<Uint8Array> {
  // No crop information, return original
  if (!crop || (!crop.left && !crop.top && !crop.right && !crop.bottom)) {
    return imageData;
  }

  // Cropping explicitly disabled, return original
  if (options.enabled === false) {
    return imageData;
  }

  try {
    const CanvasFactory = await createCanvasFactory(options.canvasImport);
    const img = await loadImage(imageData, CanvasFactory);

    // Calculate crop region (DOCX unit is 1/100000 of percentage)
    const left = ((crop.left || 0) / 100000) * img.width;
    const top = ((crop.top || 0) / 100000) * img.height;
    const right = ((crop.right || 0) / 100000) * img.width;
    const bottom = ((crop.bottom || 0) / 100000) * img.height;

    const croppedWidth = Math.round(img.width - left - right);
    const croppedHeight = Math.round(img.height - top - bottom);

    // Validate dimensions
    if (croppedWidth <= 0 || croppedHeight <= 0) {
      console.warn("Invalid crop dimensions, returning original image");
      return imageData;
    }

    // Create cropped canvas
    const drawingContext = new CanvasFactory().create(croppedWidth, croppedHeight);

    if (!drawingContext.context) {
      throw new Error("Failed to get 2D context from canvas");
    }

    // Crop and draw
    (drawingContext.context as CanvasRenderingContext2D).drawImage(
      img as HTMLImageElement,
      left,
      top,
      croppedWidth,
      croppedHeight,
      0,
      0,
      croppedWidth,
      croppedHeight,
    );

    // Convert back to buffer
    const dataUrl = drawingContext.canvas.toDataURL();
    const response = await fetch(dataUrl);
    const buffer = await response.arrayBuffer();

    return new Uint8Array(buffer);
  } catch (error) {
    // Crop failed, return original (graceful degradation)
    console.warn("Image cropping failed, returning original image:", error);
    return imageData;
  }
}

/**
 * Load image from buffer (environment-agnostic)
 *
 * @param data - Image data as Uint8Array
 * @param _CanvasFactory - Canvas factory class (unused, for compatibility)
 * @returns Loaded canvas image
 */
async function loadImage(
  data: Uint8Array,
  _CanvasFactory: typeof DOMCanvasFactory | typeof NodeCanvasFactory,
): Promise<HTMLImageElement | import("@napi-rs/canvas").Image> {
  if (isBrowser) {
    // Browser: use Image + Blob
    const blob = new Blob([data.buffer as ArrayBuffer]);
    const url = URL.createObjectURL(blob);

    try {
      const img = new Image();

      return new Promise((resolve, reject) => {
        img.onload = () => {
          URL.revokeObjectURL(url);
          resolve(img);
        };
        img.onerror = () => {
          URL.revokeObjectURL(url);
          reject(new Error("Failed to load image"));
        };
        img.src = url;
      });
    } catch (error) {
      URL.revokeObjectURL(url);
      throw error;
    }
  } else {
    // Node.js: use @napi-rs/canvas loadImage
    if (!resolvedCanvasModule) {
      throw new Error("@napi-rs/canvas module is not resolved");
    }

    return await resolvedCanvasModule.loadImage(Buffer.from(data));
  }
}
