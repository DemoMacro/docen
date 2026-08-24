/**
 * Canvas demo — the M-R1 rendering base in isolation: open a .docx, run the
 * full persistence-side pipeline (parseDOCX → compileDocument →
 * projectDocumentOptions → layoutFlow), and paint the pages with LeaferJS.
 * No editing yet — this validates the projection and the stage before the
 * text/object editing milestones wire in.
 *
 * Fonts are deliberately NOT registered: measurement (pretext's canvas
 * measureText) and painting (LeaferJS) both resolve family names through the
 * browser's own font stack — the document's 宋体 renders with the system's
 * 宋体, the same engine on both sides, so measure == render by construction.
 */
import { compileDocument, parseDOCX } from "@docen/docx";
import { projectDocumentOptions } from "@docen/docx/layout";
import { browserFontMetrics, TextMeasurer, layoutFlow } from "@docen/layout";

import { CanvasStage } from "../src/document/canvas/stage";

export const mountCanvasDemo = (stage: HTMLElement): void => {
  const measurer = new TextMeasurer(browserFontMetrics);
  let canvas: CanvasStage | null = null;

  // Toolbar: the open-file control and a status line.
  const bar = document.createElement("div");
  Object.assign(bar.style, {
    display: "flex",
    alignItems: "center",
    gap: "12px",
    padding: "8px 16px",
    borderBottom: "1px solid #e5e5e5",
    fontFamily: "Inter, sans-serif",
    fontSize: "13px",
  } satisfies Partial<CSSStyleDeclaration>);
  const file = document.createElement("input");
  file.type = "file";
  file.accept = ".docx";
  const status = document.createElement("span");
  status.textContent = "Open a .docx to render it on canvas";
  bar.append(file, status);

  // The scroll surface: pages stack in a centered fit-content column inside.
  const stageHost = document.createElement("div");
  Object.assign(stageHost.style, { height: "100%" } satisfies Partial<CSSStyleDeclaration>);
  const scroller = document.createElement("div");
  Object.assign(scroller.style, {
    flex: "1",
    overflow: "auto",
    background: "#e8e8e8",
  } satisfies Partial<CSSStyleDeclaration>);
  scroller.append(stageHost);

  const column = document.createElement("div");
  Object.assign(column.style, {
    display: "flex",
    flexDirection: "column",
    height: "100%",
  } satisfies Partial<CSSStyleDeclaration>);
  column.append(bar, scroller);
  stage.append(column);

  file.addEventListener("change", () => {
    const picked = file.files?.[0];
    if (!picked) return;
    void picked.arrayBuffer().then((buffer) => {
      const { blocks, flow } = projectDocumentOptions(compileDocument(parseDOCX(buffer)));
      const pages = layoutFlow(blocks, flow, measurer);
      canvas ??= new CanvasStage(stageHost, { metrics: browserFontMetrics, flow });
      canvas.sync(pages, flow);
      status.textContent = `${picked.name} — ${pages.length} page${pages.length === 1 ? "" : "s"}`;
    });
  });
};
