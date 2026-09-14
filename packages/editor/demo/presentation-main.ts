/**
 * Presentation demo entry — mounts `<docen-presentation>` (workspace route:
 * title-bar/ribbon/status-bar chrome + parse → project → LeaferJS slides).
 * Files open through the element's own title-bar menu.
 */
import { DocenPresentation, registerComponents } from "@docen/editor";
import { generatePresentation, type PresentationOptions, type SlideChild } from "@docen/pptx";

type ShapeChild = Extract<SlideChild, { shape: unknown }>["shape"];

const inEMU = (v: number): number => Math.round(v * 914400);

const textCard = (text: string, size: number, fill: string, y: number): ShapeChild => ({
  x: inEMU(0.9),
  y: inEMU(y),
  width: inEMU(11),
  height: inEMU(0.8),
  textBody: {
    anchor: "center",
    paragraphs: [{ children: [{ text, size, bold: true, fill }] }],
  },
});

const rectCard = (x: number, y: number, w: number, h: number, fill: string): ShapeChild => ({
  x: inEMU(x),
  y: inEMU(y),
  width: inEMU(w),
  height: inEMU(h),
  properties: { geometry: "rect", fill },
});

const demoDeck = (): PresentationOptions => ({
  slides: [
    // Title slide — accent bar, big title, subtitle.
    {
      children: [
        {
          shape: {
            x: inEMU(0.9),
            y: inEMU(2.1),
            width: inEMU(0.14),
            height: inEMU(1.3),
            properties: { geometry: "rect", fill: "4472C4" },
          },
        },
        {
          shape: {
            x: inEMU(1.3),
            y: inEMU(2.05),
            width: inEMU(10.4),
            height: inEMU(1.4),
            textBody: {
              anchor: "center",
              paragraphs: [
                {
                  children: [{ text: "Docen Presentation", size: 44, bold: true, fill: "262626" }],
                },
              ],
            },
          },
        },
        {
          shape: {
            x: inEMU(1.3),
            y: inEMU(3.55),
            width: inEMU(10.4),
            height: inEMU(0.7),
            textBody: {
              paragraphs: [
                {
                  children: [
                    { text: "A full PPTX editor rendered on canvas", size: 18, fill: "595959" },
                  ],
                },
              ],
            },
          },
        },
      ],
    },
    // Shape gallery — the variants the editor can select and transform.
    {
      children: [
        { shape: textCard("Shapes", 28, "262626", 0.55) },
        { shape: rectCard(0.9, 2, 2.7, 1.9, "4472C4") },
        {
          shape: {
            x: inEMU(4.1),
            y: inEMU(2),
            width: inEMU(1.9),
            height: inEMU(1.9),
            properties: { geometry: "ellipse", fill: "ED7D31" },
          },
        },
        {
          shape: {
            x: inEMU(6.5),
            y: inEMU(2),
            width: inEMU(2.7),
            height: inEMU(1.9),
            properties: { geometry: "roundRect", fill: "70AD47" },
          },
        },
        {
          line: {
            x1: inEMU(0.95),
            y1: inEMU(4.7),
            x2: inEMU(12),
            y2: inEMU(4.15),
            properties: { outline: { width: "3pt", color: "4472C4" } },
          },
        },
        {
          shape: {
            x: inEMU(0.9),
            y: inEMU(5.4),
            width: inEMU(11),
            height: inEMU(0.6),
            textBody: {
              paragraphs: [
                {
                  children: [
                    {
                      text: "Select, drag, resize and rotate — everything stays on the slide JSON.",
                      size: 14,
                      fill: "595959",
                    },
                  ],
                },
              ],
            },
          },
        },
      ],
    },
    // Hands-on slide — the blue card is the drag/resize/rotate demo target.
    {
      children: [
        { shape: textCard("Try it out", 28, "262626", 0.55) },
        {
          shape: {
            x: inEMU(0.9),
            y: inEMU(1.6),
            width: inEMU(11),
            height: inEMU(0.6),
            textBody: {
              paragraphs: [
                {
                  children: [
                    {
                      text: "Drag the blue card; grab the knob above it to rotate. Ctrl+Z undoes.",
                      size: 14,
                      fill: "595959",
                    },
                  ],
                },
              ],
            },
          },
        },
        { shape: rectCard(2.2, 3, 2.67, 2, "4472C4") },
        { shape: rectCard(6.2, 3.1, 1.6, 1.6, "70AD47") },
      ],
    },
  ],
});

void registerComponents().then(async () => {
  const el = document.createElement("docen-presentation") as DocenPresentation;
  el.setAttribute("filename", "Demo.pptx");
  const bytes = await generatePresentation(demoDeck(), { type: "uint8array" });
  await el.openPresentation(bytes);
  document.body.append(el);
});
