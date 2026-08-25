/**
 * Editor demo entry — registers Fluent components + theme, then routes between
 * the demos via fluent-tablist. The Document tab is the canvas route (the
 * DOM contenteditable route is gone).
 *
 * Layout is a full-height flex column (declared in index.html): the tablist is
 * a fixed-height header, the stage fills the rest.
 */
import { applyTheme, registerComponents } from "@docen/editor";

import { mountCanvasDemo } from "./canvas";
import { mountImageDemo } from "./image";

// registerComponents is async (it dynamically imports + defines the web
// components). Chain via .then — not top-level await, so this file stays
// tsc-clean under the repo tsconfig — the Image demo needs <docen-image>
// defined before it mounts. `void` marks the floating promise.
void registerComponents().then(() => {
  applyTheme("light");

  const app = document.createElement("div");
  app.id = "app";

  const nav = document.createElement("fluent-tablist");
  nav.setAttribute("aria-label", "editor demos");

  const tabs: { id: string; label: string }[] = [
    { id: "document", label: "Document" },
    { id: "image", label: "Image" },
  ];
  for (const tab of tabs) {
    const el = document.createElement("fluent-tab");
    el.id = tab.id;
    el.textContent = tab.label;
    nav.append(el);
  }

  const stage = document.createElement("main");

  app.append(nav, stage);
  document.body.append(app);

  type Route = "document" | "image";
  let current: Route = "document";

  const render = (route: Route): void => {
    stage.replaceChildren();
    if (route === "document") mountCanvasDemo(stage);
    else mountImageDemo(stage);
  };

  nav.addEventListener("change", (event: Event) => {
    const detail = (event as CustomEvent).detail as { id?: string } | undefined;
    const id = (detail?.id ?? (event.target as HTMLElement)?.id) as Route | undefined;
    if (!id || id === current) return;
    current = id;
    render(id);
  });

  // Default tab + route.
  document.getElementById("document")?.setAttribute("aria-selected", "true");
  render("document");
});
