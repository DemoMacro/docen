/**
 * Editor demo entry — registers Fluent components + theme, then routes between
 * the Document and Image demos via fluent-tablist.
 *
 * Layout is a full-height flex column (declared in index.html): the tablist is
 * a fixed-height header, the stage fills the rest.
 */
import { applyTheme, registerComponents } from "@docen/editor";

import { mountDocumentDemo } from "./document";
import { mountImageDemo } from "./image";

// registerComponents is async (it dynamically imports + defines the web
// components). Chain via .then — not top-level await, so this file stays
// tsc-clean under the repo tsconfig — and only build the UI once <docen-document>
// is upgraded. Otherwise mountDocumentDemo creates one before the custom element
// is defined, and doc.addAddin(...) throws (the element is still an unknown
// HTMLElement lacking AddinHost's methods). `void` marks the floating promise.
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
    if (route === "document") mountDocumentDemo(stage);
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
