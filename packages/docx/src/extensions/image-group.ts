import { encodeBase64 } from "@office-open/core";

import { Node } from "../core";

/**
 * ImageGroup — inline atom node carrying a DOCX drawing group (wpg:
 * wordprocessingGroup) as an opaque blob. A group bundles pictures, shapes, and
 * nested groups behind a shared transform; the editor doesn't model its
 * interior — the full WpgGroupRunOptions round-trips verbatim in attrs.wpgGroup.
 *
 * renderHTML paints the first picture (common case: a single-logo header group)
 * so the group is visible; shapes/nested groups fall back to a placeholder.
 * Structured group/shape editing is a separate phase.
 */

const EMU_PER_PX = 9525;

interface GroupChild {
  type?: string;
  data?: Uint8Array | Record<string, number>;
  transformation?: { width?: number; height?: number };
}

interface WpgGroup {
  children?: GroupChild[];
}

/** Extract the first picture (pic/MediaData) from a group for display. */
function firstPicture(group: WpgGroup | null): {
  src: string;
  width?: number;
  height?: number;
} | null {
  const pic = group?.children?.find((c) => {
    const t = c.type;
    return typeof t === "string" && t !== "wps" && t !== "wpg";
  });
  if (!pic?.data) return null;
  const bytes = pic.data instanceof Uint8Array ? pic.data : new Uint8Array(Object.values(pic.data));
  const width = pic.transformation?.width;
  const height = pic.transformation?.height;
  return {
    src: `data:image/${pic.type};base64,${encodeBase64(bytes)}`,
    width: typeof width === "number" ? Math.round(width / EMU_PER_PX) : undefined,
    height: typeof height === "number" ? Math.round(height / EMU_PER_PX) : undefined,
  };
}

const attrWpgGroup = () => ({
  default: null,
  rendered: false,
  parseHTML: (element: HTMLElement) => {
    const raw = element.getAttribute("data-wpg-group");
    if (!raw) return null;
    try {
      return JSON.parse(raw);
    } catch {
      return null;
    }
  },
});

export const ImageGroup = Node.create({
  name: "imageGroup",
  group: "inline",
  inline: true,
  atom: true,

  addAttributes() {
    return {
      wpgGroup: attrWpgGroup(),
    };
  },

  parseHTML() {
    return [{ tag: "span[data-image-group]" }];
  },

  renderHTML({ node }: { node: { attrs: Record<string, unknown> } }) {
    // The full WpgGroupRunOptions (incl. picture bytes) is NOT serialized to
    // HTML — bytes-as-object is ~6× base64 and bloats the markup. The group
    // round-trips via DOCX (attrs.wpgGroup in JSON); HTML renders only the first
    // picture for display, so parseHTML reads wpgGroup back as null.
    const attrs: Record<string, unknown> = {
      "data-image-group": "",
      contenteditable: "false",
    };
    const pic = firstPicture(node.attrs.wpgGroup as WpgGroup | null);
    if (pic) {
      const imgAttrs: Record<string, unknown> = { src: pic.src, alt: "" };
      if (pic.width && pic.height) {
        imgAttrs.style = `width:${pic.width}px;height:${pic.height}px`;
      }
      return ["span", attrs, ["img", imgAttrs]];
    }
    return ["span", attrs, "♢"];
  },
});
