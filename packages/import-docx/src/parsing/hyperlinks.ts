import { fromXml } from "xast-util-from-xml";

/**
 * Extract hyperlinks from DOCX relationships
 * Returns Map of relationship ID to hyperlink target URL
 */
export function extractHyperlinks(files: Record<string, Uint8Array>): Map<string, string> {
  const hyperlinks = new Map<string, string>();
  const relsXml = files["word/_rels/document.xml.rels"];
  if (!relsXml) return hyperlinks;

  const relsXast = fromXml(new TextDecoder().decode(relsXml));

  // Find Relationships element first (CRITICAL FIX)
  if (relsXast.type === "root") {
    for (const child of relsXast.children) {
      if (child.type === "element" && child.name === "Relationships") {
        const relationships = child;
        // Now iterate through Relationship elements
        for (const relChild of relationships.children) {
          if (relChild.type === "element" && relChild.name === "Relationship") {
            const rel = relChild;
            const type = rel.attributes.Type;
            const hyperlinkRelType =
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
            if (type && type === hyperlinkRelType) {
              const rId = rel.attributes.Id;
              const target = rel.attributes.Target;
              if (rId && target) {
                hyperlinks.set(rId as string, target as string);
              }
            }
          }
        }
        break;
      }
    }
  }
  return hyperlinks;
}
