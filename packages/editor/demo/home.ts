/**
 * Demo home — the editor picker. Registers the Fluent components, then mounts
 * one compound button per editor (the workbook entry stays disabled until its
 * editor lands).
 */
import { applyTheme, registerComponents } from "@docen/editor";

const ENTRIES: { name: string; description: string; href?: string }[] = [
  { name: "Document", description: "Word documents on the canvas engine", href: "/document.html" },
  {
    name: "Presentation",
    description: "Slide decks on the canvas engine",
    href: "/presentation.html",
  },
  { name: "Workbook", description: "Spreadsheets" },
];

void registerComponents().then(() => {
  applyTheme("light");
  const host = document.createElement("div");
  host.className = "editors";
  for (const entry of ENTRIES) {
    const card = document.createElement("fluent-compound-button");
    card.setAttribute("appearance", "outline");
    card.setAttribute("size", "large");
    const description = document.createElement("span");
    description.slot = "description";
    description.textContent = entry.description;
    card.append(entry.name, description);
    if (entry.href) {
      card.addEventListener("click", () => {
        location.href = entry.href!;
      });
    } else {
      card.setAttribute("disabled", "");
    }
    host.append(card);
  }
  document.body.append(host);
});
