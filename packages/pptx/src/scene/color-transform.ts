/** DrawingML EG_ColorTransform evaluation. Transform values are the parsed
 *  option shape (integer percent and degrees); transforms compose in source
 *  order, exactly as serialized child order specifies. */

import type { ColorTransformOptions } from "@office-open/core/drawing";

export interface ResolvedColor {
  color: string;
  /** The alpha channel after alpha/alphaMod/alphaOff (0–1). */
  alpha: number;
}

type Rgb = [red: number, green: number, blue: number];
type Hsl = [hue: number, saturation: number, lightness: number];

const clamp = (value: number, min = 0, max = 1) => Math.min(max, Math.max(min, value));
const channel = (value: number) => Math.round(clamp(value, 0, 255));
const percent = (value: number | undefined, fallback = 0) =>
  typeof value === "number" ? value / 100 : fallback;
const normalizeHue = (value: number) => ((value % 360) + 360) % 360;

function hueToRgb(offset: number, temp1: number, temp2: number): number {
  const epsilon = 1e-10;
  if (offset < 0) offset += 1;
  if (offset > 1) offset -= 1;
  if (offset < 1 / 6 - epsilon) return temp1 + (temp2 - temp1) * 6 * offset;
  if (offset < 1 / 2 - epsilon) return temp2;
  if (offset < 2 / 3 - epsilon) return temp1 + (temp2 - temp1) * (2 / 3 - offset) * 6;
  return temp1;
}

function hslToRgb([hue, saturation, lightness]: Hsl): Rgb {
  if (saturation === 0) {
    const value = lightness * 255;
    return [value, value, value];
  }
  const temp2 =
    lightness < 0.5
      ? lightness * (1 + saturation)
      : lightness + saturation - lightness * saturation;
  const temp1 = 2 * lightness - temp2;
  const base = normalizeHue(hue) / 360;
  return [
    hueToRgb(base + 1 / 3, temp1, temp2) * 255,
    hueToRgb(base, temp1, temp2) * 255,
    hueToRgb(base - 1 / 3, temp1, temp2) * 255,
  ];
}

function rgbToHsl([red, green, blue]: Rgb): Hsl {
  const r = red / 255;
  const g = green / 255;
  const b = blue / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const lightness = (max + min) / 2;
  if (max === min) return [0, 0, lightness];
  const delta = max - min;
  const saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let hue: number;
  if (max === r) hue = ((g - b) / delta + (g < b ? 6 : 0)) * 60;
  else if (max === g) hue = ((b - r) / delta + 2) * 60;
  else hue = ((r - g) / delta + 4) * 60;
  return [hue, saturation, lightness];
}

function parseHex(color: string): Rgb {
  const value = color.replace("#", "");
  const expanded =
    value.length === 3 ? value.replace(/([0-9a-f])/gi, "$1$1") : value.padStart(6, "0").slice(0, 6);
  return [
    Number.parseInt(expanded.slice(0, 2), 16),
    Number.parseInt(expanded.slice(2, 4), 16),
    Number.parseInt(expanded.slice(4, 6), 16),
  ];
}

function hexOf([red, green, blue]: Rgb): string {
  return [red, green, blue]
    .map((value) => channel(value).toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}

function withRgb(rgb: Rgb, transform: (rgb: Rgb) => Rgb): Rgb {
  return transform(rgb);
}

function withHsl(rgb: Rgb, transform: (hsl: Hsl) => Hsl): Rgb {
  return hslToRgb(transform(rgbToHsl(rgb)));
}

function gammaOf(rgb: Rgb, exponent: number): Rgb {
  return rgb.map((value) => (value / 255) ** exponent * 255) as Rgb;
}

/** Evaluate every EG_ColorTransform key; alpha is returned separately so a
 *  fill's opacity never bleeds into its stroke. Unknown values are no-ops. */
export function transformColor(
  color: string,
  transforms: ColorTransformOptions | undefined,
): ResolvedColor {
  let rgb = parseHex(color);
  let alpha = 1;
  if (!transforms) return { color: hexOf(rgb), alpha };

  for (const [name, raw] of Object.entries(transforms)) {
    if (raw === undefined) continue;
    switch (name) {
      case "tint": {
        const amount = percent(raw);
        rgb = rgb.map((value) => value + (255 - value) * amount) as Rgb;
        break;
      }
      case "shade": {
        const amount = percent(raw, 1);
        rgb = rgb.map((value) => value * amount) as Rgb;
        break;
      }
      case "alpha":
        alpha = percent(raw);
        break;
      case "alphaOff":
        alpha += percent(raw);
        break;
      case "alphaMod":
        alpha *= percent(raw, 1);
        break;
      case "comp": {
        const max = Math.max(...rgb);
        const min = Math.min(...rgb);
        rgb = rgb.map((value) => max + min - value) as Rgb;
        break;
      }
      case "inv":
        rgb = rgb.map((value) => 255 - value) as Rgb;
        break;
      case "gray": {
        const luminance = 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2];
        rgb = [luminance, luminance, luminance];
        break;
      }
      case "gamma":
        rgb = withRgb(rgb, (value) => gammaOf(value, 1 / 2.2));
        break;
      case "invGamma":
        rgb = withRgb(rgb, (value) => gammaOf(value, 2.2));
        break;
      case "hue":
      case "hueOff":
      case "hueMod":
      case "sat":
      case "satOff":
      case "satMod":
      case "lum":
      case "lumOff":
      case "lumMod":
        rgb = withHsl(rgb, ([hue, saturation, lightness]) => {
          switch (name) {
            case "hue":
              return [typeof raw === "number" ? raw : hue, saturation, lightness];
            case "hueOff":
              return [hue + (typeof raw === "number" ? raw : 0), saturation, lightness];
            case "hueMod":
              return [hue * percent(raw, 1), saturation, lightness];
            case "sat":
              return [hue, percent(raw), lightness];
            case "satOff":
              return [hue, saturation + percent(raw), lightness];
            case "satMod":
              return [hue, saturation * percent(raw, 1), lightness];
            case "lum":
              return [hue, saturation, percent(raw)];
            case "lumOff":
              return [hue, saturation, lightness + percent(raw)];
            default:
              return [hue, saturation, lightness * percent(raw, 1)];
          }
        });
        break;
      case "red":
      case "redOff":
      case "redMod":
      case "green":
      case "greenOff":
      case "greenMod":
      case "blue":
      case "blueOff":
      case "blueMod": {
        const index = name.startsWith("red") ? 0 : name.startsWith("green") ? 1 : 2;
        rgb = withRgb(rgb, (value) => {
          const next = [...value] as Rgb;
          if (name.endsWith("Off")) next[index] += 255 * percent(raw);
          else if (name.endsWith("Mod")) next[index] *= percent(raw, 1);
          else next[index] = 255 * percent(raw);
          return next;
        });
        break;
      }
    }
  }

  return { color: hexOf(rgb), alpha: clamp(alpha) };
}
