import { fromXml } from "xast-util-from-xml";

/**
 * Extract raw image data from DOCX relationships
 * Returns Map of relationship ID to raw image data (Uint8Array)
 */
export function extractImages(files: Record<string, Uint8Array>): Map<string, Uint8Array> {
  const images = new Map<string, Uint8Array>();

  const relsXml = files["word/_rels/document.xml.rels"];
  if (!relsXml) return images;

  const relsXast = fromXml(new TextDecoder().decode(relsXml));

  if (relsXast.type === "root") {
    for (const child of relsXast.children) {
      if (child.type === "element" && child.name === "Relationships") {
        const relationships = child;
        for (const relChild of relationships.children) {
          if (relChild.type === "element" && relChild.name === "Relationship") {
            const rel = relChild;
            const type = rel.attributes.Type;
            const imageRelType =
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

            if (type && type === imageRelType) {
              const rId = rel.attributes.Id;
              const target = rel.attributes.Target;
              if (rId && target) {
                // Extract image from media folder
                const imagePath = "word/" + (target as string);
                const imageData = files[imagePath];
                if (imageData) {
                  images.set(rId as string, imageData);
                }
              }
            }
          }
        }
        break;
      }
    }
  }

  return images;
}
