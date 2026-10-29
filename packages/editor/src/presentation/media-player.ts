// The media playback overlay: one <video>/<audio> element parked over the
// canvas host and positioned to the frame's screen box. The canvas keeps
// painting the poster — native controls ride the DOM layer above it (the
// same layer as the selection overlay), so play/pause/seek come free, and
// the canvas composition never reflows.

export interface MediaPlayback {
  media: "video" | "audio";
  src: string;
  /** Screen-px box of the media frame (already scaled/offset). */
  x: number;
  y: number;
  width: number;
  height: number;
  loop?: boolean;
  muted?: boolean;
  /** Normalized 0-1 browser volume. */
  volume?: number;
  trim?: { start?: number; end?: number };
  crop?: { left: number; top: number; right: number; bottom: number };
}

export class MediaPlayer {
  readonly el: HTMLDivElement;
  #media: HTMLMediaElement | null = null;
  #src: string | null = null;

  constructor(host: HTMLElement) {
    this.el = document.createElement("div");
    // The bridge's takeFocus treats a click inside a `data-docen-overlay`
    // widget as the widget's own — the controls must not drop a caret or
    // fall through to the canvas as a slide advance.
    this.el.setAttribute("data-docen-overlay", "");
    Object.assign(this.el.style, {
      position: "absolute",
      display: "none",
      zIndex: "7",
    } satisfies Partial<CSSStyleDeclaration>);
    host.append(this.el);
  }

  /** Whether the open player plays this exact source. */
  isOpenFor(src: string): boolean {
    return this.#media != null && this.#src === src;
  }

  get playing(): boolean {
    return this.#media != null && !this.#media.paused;
  }

  show(playback: MediaPlayback): void {
    this.#media?.pause();
    this.el.replaceChildren();
    this.#src = playback.src;
    const media =
      playback.media === "video"
        ? document.createElement("video")
        : document.createElement("audio");
    media.src = playback.src;
    media.controls = true;
    media.loop = playback.loop === true;
    media.muted = playback.muted === true;
    if (playback.volume !== undefined) media.volume = playback.volume;
    // A video fills its frame (letterboxed, black-backed like PowerPoint's
    // player); an audio frame keeps its poster — only a compact control bar
    // docks at the frame's bottom.
    if (playback.media === "video") {
      Object.assign(this.el.style, {
        left: `${playback.x}px`,
        top: `${playback.y}px`,
        width: `${playback.width}px`,
        height: `${playback.height}px`,
      });
      Object.assign(media.style, {
        width: "100%",
        height: "100%",
        objectFit: playback.crop ? "fill" : "contain",
        background: "#000",
      });
      if (playback.crop) {
        const { left, top, right, bottom } = playback.crop;
        const width = Math.max(0.01, 1 - left - right);
        const height = Math.max(0.01, 1 - top - bottom);
        Object.assign(media.style, {
          transformOrigin: `${(left / width) * 100}% ${(top / height) * 100}%`,
          transform: `scale(${1 / width}, ${1 / height})`,
        });
      }
    } else {
      const barWidth = Math.min(playback.width, 320);
      Object.assign(this.el.style, {
        left: `${playback.x + (playback.width - barWidth) / 2}px`,
        top: `${playback.y + playback.height - 40}px`,
        width: `${barWidth}px`,
        height: "36px",
      });
      Object.assign(media.style, { width: "100%", height: "100%" });
    }
    this.el.style.display = "block";
    this.el.append(media);
    this.#media = media;
    const trim = playback.trim;
    if (trim?.start !== undefined)
      media.addEventListener(
        "loadedmetadata",
        () => {
          if (this.#media === media && media.currentTime < trim.start!)
            media.currentTime = trim.start!;
        },
        { once: true },
      );
    if (trim?.end !== undefined)
      media.addEventListener("timeupdate", () => {
        if (this.#media !== media || media.currentTime < trim.end!) return;
        if (media.loop) {
          if (trim.start !== undefined) media.currentTime = trim.start;
          return;
        }
        media.pause();
        media.currentTime = trim.end!;
      });
    // Play right away — PowerPoint's player starts on open. A rejected play
    // (a strict autoplay policy) just leaves the controls waiting.
    void media.play().catch(() => {});
  }

  toggle(): void {
    const media = this.#media;
    if (!media) return;
    if (media.paused) void media.play();
    else media.pause();
  }

  hide(): void {
    if (!this.#media) return;
    this.#media.pause();
    this.el.replaceChildren();
    this.el.style.display = "none";
    this.#media = null;
    this.#src = null;
  }

  destroy(): void {
    this.hide();
    this.el.remove();
  }
}
