// Media frame (p:pic with a:videoFile / a:audioFile) projection. The bytes
// stay in the source model for round-trip; the canvas carries the poster and
// enough identity to paint a stable player, plus the playable data URL for
// the browser-decodable formats (wmv/avi/wma have no browser decoder).

import type { LayoutDrawingMember, LayoutPictureCrop } from "@docen/layout";
import type { SourceRectangleOptions } from "@office-open/core/drawing";
import type { AudioFrameOptions, VideoFrameOptions } from "@office-open/pptx";

import { emuOf, type Xform } from "./geometry";
import { pictureSrcOf } from "./pictures";

/** The container formats the browser can play, by media family. */
const PLAYABLE_MIME: Partial<Record<string, string>> = {
  mp4: "video/mp4",
  mov: "video/quicktime",
  mp3: "audio/mpeg",
  wav: "audio/wav",
  aac: "audio/aac",
};

function base64Of(bytes: Uint8Array): string {
  let binary = "";
  for (let index = 0; index < bytes.length; index += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
  }
  return btoa(binary);
}

function posterUrlOf(data: unknown, type: string | undefined): string | undefined {
  if (typeof data === "string") {
    if (data.startsWith("data:")) return data;
    return pictureSrcOf(data, type);
  }
  return pictureSrcOf(data, type);
}

function cropOf(rect: SourceRectangleOptions | undefined): LayoutPictureCrop | undefined {
  const left = rect?.left ?? 0;
  const top = rect?.top ?? 0;
  const right = rect?.right ?? 0;
  const bottom = rect?.bottom ?? 0;
  if (!left && !top && !right && !bottom) return undefined;
  return { left: left / 100, top: top / 100, right: right / 100, bottom: bottom / 100 };
}

function playbackOf(options: VideoFrameOptions | AudioFrameOptions) {
  const playback = options.playback;
  if (!playback && !options.trim) return undefined;
  const { loop, mute, volume } = playback ?? {};
  return {
    ...(loop !== undefined ? { loop } : {}),
    ...(mute !== undefined ? { muted: mute } : {}),
    ...(volume !== undefined ? { volume: Math.max(0, Math.min(1, volume)) } : {}),
    ...(options.trim ? { trim: options.trim } : {}),
  };
}

export function mediaMember(
  media: "video" | "audio",
  options: VideoFrameOptions | AudioFrameOptions,
  t: Xform,
  childPath: readonly number[] | undefined,
): LayoutDrawingMember {
  const src = posterUrlOf(options.poster, options.posterType);
  const playableMime = PLAYABLE_MIME[options.type ?? ""];
  const playableSrc =
    playableMime && options.data != null
      ? typeof options.data === "string"
        ? options.data.startsWith("data:")
          ? options.data
          : `data:${playableMime};base64,${options.data}`
        : options.data instanceof Uint8Array
          ? `data:${playableMime};base64,${base64Of(options.data)}`
          : undefined
      : undefined;
  return {
    kind: "mediaFrame",
    x: t.sx * emuOf(options.x) + t.dx,
    y: t.sy * emuOf(options.y) + t.dy,
    width: t.sx * emuOf(options.width),
    height: t.sy * emuOf(options.height),
    ...(childPath ? { childPath } : {}),
    media,
    ...(src ? { src } : {}),
    ...(cropOf(options.sourceRectangle) ? { crop: cropOf(options.sourceRectangle)! } : {}),
    ...(options.fileName ? { fileName: options.fileName } : {}),
    ...(playableSrc && playableMime ? { playable: { src: playableSrc, mime: playableMime } } : {}),
    ...(playbackOf(options) ? { playback: playbackOf(options)! } : {}),
  };
}
