import { fromXml } from "xast-util-from-xml";
import type { ListTypeMap, ListInfo } from "./types";
import { findChild, findDeepChildren } from "../utils/xml";

/**
 * Parse numbering.xml to build list type map
 */
export function parseNumberingXml(files: Record<string, Uint8Array>): ListTypeMap {
  const listTypeMap = new Map<string, ListInfo>();
  const abstractNumStarts = new Map<string, number>();
  const numberingXml = files["word/numbering.xml"];
  if (!numberingXml) return listTypeMap;

  const numberingXast = fromXml(new TextDecoder().decode(numberingXml));
  const abstractNumFormats = new Map<string, string>();

  const numbering = findChild(numberingXast, "w:numbering");
  if (!numbering) return listTypeMap;

  // First pass: collect all abstractNum definitions
  const abstractNums = findDeepChildren(numbering, "w:abstractNum");
  for (const abstractNum of abstractNums) {
    const abstractNumId = abstractNum.attributes["w:abstractNumId"] as string;
    const lvl = findChild(abstractNum, "w:lvl");
    if (!lvl) continue;

    const numFmt = findChild(lvl, "w:numFmt");
    if (numFmt?.attributes["w:val"]) {
      abstractNumFormats.set(abstractNumId, numFmt.attributes["w:val"] as string);
    }

    const start = findChild(lvl, "w:start");
    if (start?.attributes["w:val"]) {
      abstractNumStarts.set(abstractNumId, parseInt(start.attributes["w:val"] as string, 10));
    }
  }

  // Second pass: map numId to list type
  const nums = findDeepChildren(numbering, "w:num");
  for (const num of nums) {
    const numId = num.attributes["w:numId"] as string;
    const abstractNumId = findChild(num, "w:abstractNumId");
    if (!abstractNumId?.attributes["w:val"]) continue;

    const abstractNumIdVal = abstractNumId.attributes["w:val"] as string;
    const numFmt = abstractNumFormats.get(abstractNumIdVal);
    if (!numFmt) continue;

    const start = abstractNumStarts.get(abstractNumIdVal);

    if (numFmt === "bullet") {
      listTypeMap.set(numId, { type: "bullet" });
    } else {
      listTypeMap.set(numId, {
        type: "ordered",
        ...(start !== undefined && { start }),
      });
    }
  }

  return listTypeMap;
}
