/**
 * Presentation demo entry — mounts `<docen-presentation>` (workspace route:
 * title-bar/ribbon/status-bar chrome + parse → project → LeaferJS slides).
 * Files open through the element's own title-bar menu.
 */
import { registerComponents } from "@docen/editor";

void registerComponents().then(() => {
  document.body.append(document.createElement("docen-presentation"));
});
