/** Pixel-level picture tools — the browser-canvas transforms behind Word's
 *  Compress Pictures and Set Transparent Color. Both re-encode the source
 *  into a new data URL the editor writes back through `picture-pixels`
 *  (the frame keeps its size — only the pixels change). */

/** A source rectangle as the image attrs carry it (integer percent insets,
 *  10 = 10% from that edge). */
export interface CropRect {
  left?: number;
  top?: number;
  right?: number;
  bottom?: number;
}

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("image decode failed"));
    img.src = src;
  });
}

function toDataUrl(canvas: HTMLCanvasElement, hasAlpha: boolean): string {
  // Lossy JPEG only where transparency cannot exist; quality 0.85 matches
  // the visual bar Word's 220ppi preset lands at.
  return hasAlpha ? canvas.toDataURL("image/png") : canvas.toDataURL("image/jpeg", 0.85);
}

function cropFractions(crop: CropRect | undefined): {
  l: number;
  t: number;
  r: number;
  b: number;
} {
  return {
    l: Math.min(Math.max(crop?.left ?? 0, 0), 100) / 100,
    t: Math.min(Math.max(crop?.top ?? 0, 0), 100) / 100,
    r: Math.min(Math.max(crop?.right ?? 0, 0), 100) / 100,
    b: Math.min(Math.max(crop?.bottom ?? 0, 0), 100) / 100,
  };
}

/** Resample the picture to `target` CSS pixels at most (never up — a target
 *  past the natural size keeps the natural pixels, Word's behavior when the
 *  source dpi already beats the preset). `dropCrop` bakes the crop into the
 *  pixels so the written-back attrs can drop the srcRect entirely. */
export async function compressPictureSrc(
  src: string,
  target: { width: number; height: number },
  crop?: CropRect,
  dropCrop?: boolean,
): Promise<string> {
  const img = await loadImage(src);
  const c = cropFractions(crop);
  // The baked pixels cover the cropped area only when dropping the crop;
  // otherwise the full source stays (the srcRect keeps applying on top).
  const sx = dropCrop ? c.l * img.naturalWidth : 0;
  const sy = dropCrop ? c.t * img.naturalHeight : 0;
  const sw = (dropCrop ? 1 - c.l - c.r : 1) * img.naturalWidth;
  const sh = (dropCrop ? 1 - c.t - c.b : 1) * img.naturalHeight;
  const width = Math.max(1, Math.round(Math.min(target.width, sw)));
  const height = Math.max(1, Math.round(Math.min(target.height, sh)));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("no 2d context");
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = "high";
  ctx.drawImage(img, sx, sy, sw, sh, 0, 0, width, height);
  const data = ctx.getImageData(0, 0, width, height).data;
  let hasAlpha = false;
  for (let i = 3; i < data.length; i += 4) {
    if (data[i]! < 255) {
      hasAlpha = true;
      break;
    }
  }
  return toDataUrl(canvas, hasAlpha);
}

/** Sample the picture at a display-relative point (`nx`/`ny` in 0..1 of the
 *  visible frame) and clear every pixel matching that color — Word's Set
 *  Transparent Color. The output is always PNG (transparency needs it); the
 *  sample point un-maps through the crop so a cropped picture still reads
 *  the color the user sees. */
export async function pickTransparentColor(
  src: string,
  nx: number,
  ny: number,
  crop?: CropRect,
): Promise<string> {
  const img = await loadImage(src);
  const canvas = document.createElement("canvas");
  canvas.width = img.naturalWidth;
  canvas.height = img.naturalHeight;
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  if (!ctx) throw new Error("no 2d context");
  ctx.drawImage(img, 0, 0);
  const image = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const data = image.data;
  const c = cropFractions(crop);
  const px = Math.round((c.l + nx * (1 - c.l - c.r)) * (canvas.width - 1));
  const py = Math.round((c.t + ny * (1 - c.t - c.b)) * (canvas.height - 1));
  const at = (py * canvas.width + px) * 4;
  const [r0, g0, b0] = [data[at]!, data[at + 1]!, data[at + 2]!];
  // A tight tolerance rides Word's built-in one: photo noise near the pick
  // must survive, flat fills must fall. ~18 per channel.
  const TOL = 32;
  const within = (v: number, r: number, g: number, b: number): boolean =>
    (v - r) * (v - r) + (v - g) * (v - g) + (v - b) * (v - b) <= TOL * TOL;
  for (let i = 0; i < data.length; i += 4) {
    if (within(data[i]!, r0, g0, b0)) data[i + 3] = 0;
  }
  ctx.putImageData(image, 0, 0);
  return canvas.toDataURL("image/png");
}
