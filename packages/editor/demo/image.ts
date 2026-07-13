/**
 * Image demo — a single `<docen-image>` with Compound Buttons to switch preset
 * attrs (plain / rotated / cropped / outlined). Uses only Fluent UI components.
 */
import type { RenderImageInput } from "@docen/core";

const SAMPLE_SRC = "https://placehold.co/600x400/png";

interface Preset {
  id: string;
  label: string;
  description: string;
  attrs: RenderImageInput;
}

const presets: Preset[] = [
  {
    id: "plain",
    label: "Plain",
    description: "400×300, no transform",
    attrs: { src: SAMPLE_SRC, width: 400, height: 300 },
  },
  {
    id: "rotated",
    label: "Rotated",
    description: "15° clockwise rotation",
    attrs: { src: SAMPLE_SRC, width: 400, height: 300, rotation: 15 },
  },
  {
    id: "cropped",
    label: "Cropped",
    description: "srcRect: 20% left, 10% top",
    attrs: { src: SAMPLE_SRC, width: 400, height: 300, crop: { left: 20000, top: 10000 } },
  },
  {
    id: "outlined",
    label: "Outlined",
    description: "2px red border",
    attrs: {
      src: SAMPLE_SRC,
      width: 400,
      height: 300,
      outline: { color: { value: "#ef4444" }, width: 19050 },
    },
  },
];

export const mountImageDemo = (stage: HTMLElement): void => {
  const presetRow = document.createElement("div");
  presetRow.className = "demo-row";

  const image = document.createElement("docen-image") as HTMLElement & {
    attrs: string;
    selected: boolean;
    addEventListener: (type: string, listener: (e: CustomEvent) => void) => void;
  };
  image.addEventListener("click", () => {
    image.selected = !image.selected;
  });

  const applyPreset = (preset: Preset): void => {
    image.attrs = JSON.stringify(preset.attrs);
  };

  for (const preset of presets) {
    const btn = document.createElement("fluent-compound-button");
    btn.textContent = preset.label;
    btn.setAttribute("secondary", preset.description);
    btn.setAttribute("appearance", "outline");
    btn.addEventListener("click", () => applyPreset(preset));
    presetRow.append(btn);
  }

  const readout = document.createElement("p");
  readout.className = "demo-readout";
  readout.textContent = "Interact with the image — the edited geometry appears here.";
  image.addEventListener("change", (event: CustomEvent) => {
    readout.textContent = `geometry: ${JSON.stringify(event.detail)}`;
  });

  applyPreset(presets[0]!);

  stage.append(presetRow, image, readout);
  requestAnimationFrame(() => {
    image.selected = true;
  });
};
