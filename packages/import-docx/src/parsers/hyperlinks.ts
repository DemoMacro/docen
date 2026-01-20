import { fromXml } from "xast-util-from-xml";
import { findChild, findDeepChildren } from "@docen/utils";

const HYPERLINK_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

/**
 * Extract hyperlinks from DOCX relationships
 * Returns Map of relationship ID to hyperlink target URL
 */
export function extractHyperlinks(files: Record<string, Uint8Array>): Map<string, string> {
  const hyperlinks = new Map<string, string>();
  const relsXml = files["word/_rels/document.xml.rels"];
  if (!relsXml) return hyperlinks;

  const relsXast = fromXml(new TextDecoder().decode(relsXml));
  const relationships = findChild(relsXast, "Relationships");
  if (!relationships) return hyperlinks;

  const rels = findDeepChildren(relationships, "Relationship");
  for (const rel of rels) {
    if (rel.attributes.Type === HYPERLINK_REL_TYPE && rel.attributes.Id && rel.attributes.Target) {
      hyperlinks.set(rel.attributes.Id as string, rel.attributes.Target as string);
    }
  }

  return hyperlinks;
}
