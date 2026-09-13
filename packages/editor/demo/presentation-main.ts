/**
 * Presentation demo entry — mounts `<docen-presentation>` (viewer route:
 * parse → project → LeaferJS slides) with a file opener.
 */
import { registerComponents, type DocenPresentation } from "@docen/editor";

void registerComponents().then(() => {
  const el = document.createElement("docen-presentation") as DocenPresentation;
  document.body.append(el);

  const file = document.querySelector<HTMLInputElement>("#file")!;
  file.addEventListener("change", () => {
    const picked = file.files?.[0];
    if (!picked) return;
    void picked.arrayBuffer().then((buffer) => el.openPresentation(buffer));
  });
});
