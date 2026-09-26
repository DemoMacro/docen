/**
 * Presentation demo entry — mounts `<docen-presentation>` (workspace route:
 * title-bar/ribbon/status-bar chrome + parse → project → LeaferJS slides).
 * Files open through the element's own title-bar menu.
 */
import { DocenPresentation, registerComponents } from "@docen/editor";
import { generatePresentation, type PresentationOptions, type SlideChild } from "@docen/pptx";

type ShapeChild = Extract<SlideChild, { shape: unknown }>["shape"];

const inEMU = (v: number): number => Math.round(v * 914400);

// A tiny 8 kHz mono sine — real playable bytes so the media slide demos
// native playback without shipping an asset.
const demoWav = (): Uint8Array => {
  const rate = 8000;
  const samples = rate;
  const bytes = new Uint8Array(44 + samples);
  const view = new DataView(bytes.buffer);
  const ascii = (offset: number, text: string) => {
    for (let i = 0; i < text.length; i++) view.setUint8(offset + i, text.charCodeAt(i));
  };
  ascii(0, "RIFF");
  view.setUint32(4, 36 + samples, true);
  ascii(8, "WAVE");
  ascii(12, "fmt ");
  view.setUint32(16, 16, true);
  view.setUint16(20, 1, true);
  view.setUint16(22, 1, true);
  view.setUint32(24, rate, true);
  view.setUint32(28, rate, true);
  view.setUint16(32, 1, true);
  view.setUint16(34, 8, true);
  ascii(36, "data");
  view.setUint32(40, samples, true);
  for (let i = 0; i < samples; i++)
    bytes[44 + i] = 128 + Math.round(100 * Math.sin((2 * Math.PI * 440 * i) / rate));
  return bytes;
};

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

// The demo master's theme: the background fill styles the slides' bgRef
// references resolve through (a flat color and a fade into the accent).
const demoTheme = {
  colorScheme: { accent1: "4472C4", dark1: "262626", light1: "FFFFFF" },
  formatScheme: {
    backgroundFillStyles: [
      { type: "solid", color: { value: "phClr" } },
      {
        type: "gradient",
        shade: { angle: 90 },
        stops: [
          { position: 0, color: { value: "phClr" } },
          { position: 1, color: { value: "accent1" } },
        ],
      },
      { type: "solid", color: { value: "phClr" } },
    ],
    fillStyles: [],
    lineStyles: [],
    effectStyles: [],
  },
};

const demoDeck = (): PresentationOptions => ({
  masters: [{ name: "demo", theme: demoTheme }] as PresentationOptions["masters"],
  slides: [
    // Title slide — accent bar, big title, subtitle.
    {
      // The bgRef resolves through the master's fade style (white → accent).
      background: { reference: { index: 1002, color: { value: "bg1" } } },
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
    // Graphic-frame table — spans, fills, cell borders, margins, anchors.
    {
      children: [
        { shape: textCard("Tables", 28, "262626", 0.55) },
        {
          table: {
            x: inEMU(0.9),
            y: inEMU(1.7),
            width: inEMU(11),
            columnWidths: [inEMU(3.4), inEMU(3.8), inEMU(3.8)],
            rows: [
              {
                cells: [
                  { text: "Region", fill: "4472C4" },
                  { text: "Q1", fill: "4472C4", verticalAlign: "center" },
                  { text: "Q2", fill: "4472C4", verticalAlign: "center" },
                ],
              },
              {
                cells: [{ text: "North", columnSpan: 2, fill: "E8F0FE" }, { text: "128" }],
              },
              {
                cells: [
                  {
                    text: "Merged",
                    rowSpan: 2,
                    fill: "FFF2CC",
                    verticalAlign: "center",
                    borders: {
                      left: { width: "2pt", color: "C00000" },
                      right: { width: "1pt", color: "C00000", dashStyle: "dash" },
                    },
                  },
                  { text: "South", margins: { left: 300000 } },
                  { text: "97", verticalAlign: "bottom" },
                ],
              },
              { cells: [{ text: "West" }, { text: "64" }] },
            ],
          },
        },
      ],
    },
    // Table styles — the tblPr flags resolving through the themed default
    // family (header fill, white bold text, light-accent banding) and the
    // no-style grid GUID as the bare comparison.
    {
      children: [
        { shape: textCard("Table styles", 28, "262626", 0.55) },
        {
          table: {
            x: inEMU(0.9),
            y: inEMU(1.7),
            width: inEMU(11),
            columnWidths: [inEMU(3.7), inEMU(3.7), inEMU(3.6)],
            firstRow: true,
            bandRow: true,
            rows: [
              {
                cells: [{ text: "Region" }, { text: "Q1" }, { text: "Q2" }],
              },
              {
                cells: [{ text: "North" }, { text: "128" }, { text: "116" }],
              },
              {
                cells: [{ text: "South" }, { text: "97" }, { text: "105" }],
              },
              {
                cells: [{ text: "West" }, { text: "64" }, { text: "71" }],
              },
            ],
          },
        },
        {
          table: {
            x: inEMU(0.9),
            y: inEMU(5),
            width: inEMU(11),
            height: inEMU(1.6),
            columnWidths: [inEMU(3.7), inEMU(3.7), inEMU(3.6)],
            tableStyleId: "{5940675A-B579-460E-94D1-54222C63F5DA}",
            rows: [
              {
                cells: [{ text: "No Style, Table Grid" }, { text: "a" }, { text: "b" }],
              },
              {
                cells: [{ text: "keeps the plain grid" }, { text: "c" }, { text: "d" }],
              },
            ],
          },
        },
      ],
    },
    // Background fills — a linear gradient slide and a radial one; the
    // default solid path stays the plain first slide.
    {
      background: {
        fill: {
          type: "gradient",
          angle: 45,
          stops: [
            { position: 0, color: "1F3864" },
            { position: 1, color: "4472C4" },
          ],
        },
      },
      children: [
        {
          shape: {
            x: inEMU(1.3),
            y: inEMU(3),
            width: inEMU(10.4),
            height: inEMU(1.2),
            textBody: {
              anchor: "center",
              paragraphs: [
                {
                  children: [{ text: "Gradient background", size: 40, bold: true, fill: "FFFFFF" }],
                },
              ],
            },
          },
        },
      ],
    },
    {
      background: {
        fill: {
          type: "gradient",
          options: {
            stops: [
              { position: 0, color: { value: "FFFFFF" } },
              { position: 1, color: { value: "FFD966" } },
            ],
            shade: { path: "circle" },
          },
        },
      },
      children: [
        {
          shape: {
            x: inEMU(1.3),
            y: inEMU(3),
            width: inEMU(10.4),
            height: inEMU(1.2),
            textBody: {
              anchor: "center",
              paragraphs: [
                {
                  children: [{ text: "Radial background", size: 40, bold: true, fill: "1F3864" }],
                },
              ],
            },
          },
        },
      ],
    },
    // Media playback — a real WAV the editor plays natively on click.
    {
      children: [
        { shape: textCard("Media", 28, "262626", 0.55) },
        {
          audio: {
            x: inEMU(4),
            y: inEMU(3.2),
            width: inEMU(6),
            height: inEMU(1),
            data: demoWav(),
            type: "wav",
            fileName: "tone.wav",
          },
        },
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
