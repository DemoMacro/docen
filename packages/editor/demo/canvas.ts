/**
 * Canvas demo — the M-R1 rendering base with the editing bridge: open a
 * .docx, run the full pipeline (parseDOCX → compileDocument →
 * projectDocumentOptions → layoutFlow), paint the pages with LeaferJS, and
 * keep editing — a viewless Tiptap editor drives the same pipeline from its
 * transactions (raf-merged full re-flow), with text input arriving through
 * the bridge's textarea.
 *
 * Fonts are deliberately NOT registered: measurement (pretext's canvas
 * measureText) and painting (LeaferJS) both resolve family names through the
 * browser's own font stack — the document's 宋体 renders with the system's
 * 宋体, the same engine on both sides, so measure == render by construction.
 */
import { compileDocument, parseDOCX, type JSONContent } from "@docen/docx";
import { projectDocumentOptions } from "@docen/docx/layout";
import { browserFontMetrics, TextMeasurer, layoutFlow } from "@docen/layout";

import { mountEditBridge, type EditBridge } from "../src/document/canvas/edit-bridge";
import { CanvasStage } from "../src/document/canvas/stage";

export const mountCanvasDemo = (stage: HTMLElement): void => {
  const measurer = new TextMeasurer(browserFontMetrics);
  let canvas: CanvasStage | null = null;
  let bridge: EditBridge | null = null;

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
  // position:relative anchors the edit bridge's textarea overlay.
  const stageHost = document.createElement("div");
  Object.assign(stageHost.style, {
    height: "100%",
    position: "relative",
  } satisfies Partial<CSSStyleDeclaration>);
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
      bridge?.destroy();
      bridge = null;
      const json = parseDOCX(buffer);
      // The render entry both sides share: the initial paint and every edited
      // transaction flow through the same full pipeline.
      const render = (doc: JSONContent): void => {
        const { blocks, flow, furniture } = projectDocumentOptions(compileDocument(doc));
        const pages = layoutFlow(blocks, flow, measurer);
        canvas ??= new CanvasStage(stageHost, { metrics: browserFontMetrics, flow, furniture });
        canvas.sync(pages, flow);
        status.textContent = `${picked.name} — ${pages.length} page${pages.length === 1 ? "" : "s"}`;
      };
      render(json);
      bridge = mountEditBridge({ host: stageHost, content: json, onDoc: render });
      // Debug handle for interactive verification (demo-only).
      Object.assign(window, { docenCanvasDebug: { bridge } });
    });
  });
};
