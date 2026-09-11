import {
  parseMarkdownTokens,
  renderMarkdownBlocks,
  type MarkdownMapper,
  type MdAlignment,
  type MdBlock,
  type MdInline,
  type MdMark,
} from "@docen/markdown";
import { LevelFormat } from "@office-open/docx";

import type { JSONContent } from "../core";
import {
  assignOrderedReferences,
  buildOrderedLevels,
  HTML_ORDERED_TEMP,
} from "../extensions/list-numbering";
import { HEADING_COMPILE_MAP, HEADING_PARSE_MAP } from "../extensions/paragraph";

type NodeAttrs = Record<string, any>;

/** A mark node as `JSONContent["marks"]` declares it (an inline shape, wider
 *  than JSONContent itself). */
type TiptapMark = NonNullable<JSONContent["marks"]>[number];

/** Per-reference per-level custom numbering start (reference → level → start),
 *  sourced from paragraphs or the abstract definitions. */
type StartLookup = Map<string, Map<number, number>>;

// ── Parse: Markdown IR → Tiptap JSON ──

const markToTiptap = (mark: MdMark): TiptapMark | null => {
  switch (mark.type) {
    case "bold":
      return { type: "bold" };
    case "italic":
      return { type: "italic" };
    case "strike":
      return { type: "strike" };
    case "code":
      return { type: "code" };
    case "link":
      return {
        type: "link",
        attrs: {
          href: mark.href,
          title: mark.title ?? null,
          target: mark.href.startsWith("#") ? null : "_blank",
        },
      };
  }
};

export function parseInline(inlines: MdInline[]): JSONContent[] {
  const out: JSONContent[] = [];
  for (const node of inlines) {
    if (node.type === "text") {
      if (!node.text) continue;
      const marks = (node.marks ?? []).map(markToTiptap).filter((m): m is TiptapMark => m != null);
      if (marks.length > 0) out.push({ type: "text", text: node.text, marks });
      else out.push({ type: "text", text: node.text });
    } else if (node.type === "image") {
      out.push({
        type: "image",
        attrs: { src: node.src, alt: node.alt ?? null, title: node.title ?? null },
      });
    } else {
      out.push({ type: "hardBreak" });
    }
  }
  return out;
}

function cellContent(blocks: MdBlock[], align: MdAlignment): JSONContent[] {
  const nodes = blocks.flatMap(parseBlockNodes);
  if (align) {
    // GFM column alignment lands on plain paragraphs only (headings/code keep
    // their own semantics).
    for (const node of nodes) {
      if (node.type === "paragraph" && !node.attrs?.heading && !node.attrs?.style) {
        node.attrs = { ...node.attrs, alignment: align };
      }
    }
  }
  return nodes.length > 0 ? nodes : [{ type: "paragraph" }];
}

export function parseBlockNodes(block: MdBlock): JSONContent[] {
  switch (block.type) {
    case "heading": {
      const depth = Math.min(9, Math.max(1, block.depth));
      return [
        {
          type: "paragraph",
          attrs: { heading: HEADING_COMPILE_MAP[depth] ?? "Heading1" },
          content: parseInline(block.children),
        },
      ];
    }
    case "paragraph":
      return [{ type: "paragraph", content: parseInline(block.children) }];
    // A code block is a paragraph styled "Code" (no dedicated OOXML element);
    // the text keeps its newlines — the compiler splits them into <w:br/>.
    // The fence's info string has no OOXML home either — it rides the
    // docen-only codeLanguage attr (renderDocx skips it).
    case "codeBlock":
      return [
        {
          type: "paragraph",
          attrs: { style: "Code", codeLanguage: block.lang ?? null },
          content: [{ type: "text", text: block.text }],
        },
      ];
    case "thematicBreak":
      return [{ type: "paragraph", attrs: { thematicBreak: true } }];
    // A footnote definition becomes a plain paragraph carrying its label —
    // DOCX has no importable footnote syntax, so the text survives verbatim.
    case "footnoteDefinition":
      return [
        {
          type: "paragraph",
          content: [{ type: "text", text: `[^${block.label}]: ` }, ...parseInline(block.children)],
        },
      ];
    // A quote paragraph carries Word's built-in "IntenseQuote" style instead
    // of a wrapper node.
    case "quote": {
      const out: JSONContent[] = [];
      for (const inner of block.blocks) {
        for (const node of parseBlockNodes(inner)) {
          out.push({ ...node, attrs: { ...node.attrs, style: "IntenseQuote" } });
        }
      }
      return out;
    }
    case "list": {
      // The schema has no list nodes — the tree flattens to paragraphs with
      // bullet/numbering attrs; nested lists recurse one level deeper. Ordered
      // items carry the html-ordered placeholder reference; parseMarkdown
      // rewrites each run to a fresh generated reference. A GFM task item has
      // no DOCX semantics — its checkbox survives as a text prefix, and a
      // custom start rides the first item's numbering per level (parseMarkdown
      // lifts these into the reference's abstractNum definition).
      const out: JSONContent[] = [];
      const walk = (list: Extract<MdBlock, { type: "list" }>, level: number): void => {
        let firstItem = true;
        for (const item of list.items) {
          let firstBlock = true;
          for (const inner of item.blocks) {
            if (inner.type === "list") {
              walk(inner, level + 1);
              continue;
            }
            const isStartItem = list.ordered && list.start !== undefined && firstItem && firstBlock;
            const listAttrs: NodeAttrs = list.ordered
              ? {
                  numbering: {
                    reference: HTML_ORDERED_TEMP,
                    level,
                    ...(isStartItem ? { start: list.start } : {}),
                  },
                }
              : { bullet: { level } };
            const nodes = parseBlockNodes(inner);
            if (item.checked !== undefined && nodes[0]?.type === "paragraph") {
              const prefix: JSONContent = { type: "text", text: item.checked ? "[x] " : "[ ] " };
              nodes[0] = { ...nodes[0], content: [prefix, ...(nodes[0].content ?? [])] };
            }
            for (const node of nodes) {
              out.push({ ...node, attrs: { ...listAttrs, ...node.attrs } });
            }
            firstBlock = false;
          }
          firstItem = false;
        }
      };
      walk(block, 0);
      return out;
    }
    case "table": {
      if (block.rows.length === 0) return [];
      return [
        {
          type: "table",
          content: block.rows.map((row, rowIndex) => ({
            type: "tableRow",
            // Word models header-ness as the ROW's tblHeader.
            ...(rowIndex === 0 && row.cells.some((cell) => cell.header)
              ? { attrs: { tableHeader: true } }
              : {}),
            content: row.cells.map((cell, col) => ({
              type: "tableCell",
              content: cellContent(cell.blocks, block.align[col] ?? null),
            })),
          })),
        },
      ];
    }
  }
}

// ── Serialize: Tiptap JSON → Markdown IR ──

const markToIr = (marks: JSONContent[] | undefined): MdMark[] => {
  const out: MdMark[] = [];
  for (const mark of marks ?? []) {
    switch (mark.type) {
      case "bold":
        out.push({ type: "bold" });
        break;
      case "italic":
        out.push({ type: "italic" });
        break;
      case "strike":
        out.push({ type: "strike" });
        break;
      case "code":
        out.push({ type: "code" });
        break;
      case "link":
        if (mark.attrs?.href) {
          const md: MdMark = { type: "link", href: mark.attrs.href };
          if (mark.attrs.title) md.title = mark.attrs.title;
          out.push(md);
        }
        break;
      // highlight/underline/subscript/superscript/textStyle/ruby/insertion/
      // deletion have no Markdown equivalent — the text survives, marks drop.
    }
  }
  return out;
};

function inlineOf(node: JSONContent): MdInline[] {
  const out: MdInline[] = [];
  const walk = (children: JSONContent[] | undefined): void => {
    for (const child of children ?? []) {
      if (child.type === "text") {
        const marks = markToIr(child.marks);
        out.push(
          marks.length > 0
            ? { type: "text", text: child.text ?? "", marks }
            : { type: "text", text: child.text ?? "" },
        );
      } else if (child.type === "image") {
        const attrs = (child.attrs ?? {}) as NodeAttrs;
        const md: MdInline = { type: "image", src: attrs.src ?? "" };
        if (attrs.alt) md.alt = attrs.alt;
        if (attrs.title) md.title = attrs.title;
        out.push(md);
      } else if (child.type === "hardBreak") {
        out.push({ type: "hardBreak" });
      } else {
        // Unknown inline container (inline passthrough, tab): keep its text.
        walk(child.content);
      }
    }
  };
  walk(node.content);
  return out;
}

/** Plain text of a paragraph with hard breaks as newlines (code blocks). */
function textOf(node: JSONContent): string {
  const parts: string[] = [];
  const walk = (children: JSONContent[] | undefined): void => {
    for (const child of children ?? []) {
      if (child.type === "text") parts.push(child.text ?? "");
      else if (child.type === "hardBreak") parts.push("\n");
      else walk(child.content);
    }
  };
  walk(node.content);
  return parts.join("");
}

/** A GFM footnote definition line (`[^label]: body`). */
const FOOTNOTE_DEF = /^\[\^([^\]\s]+)\]:\s*([\s\S]*)$/;

function paragraphToBlock(node: JSONContent): MdBlock {
  const attrs = (node.attrs ?? {}) as NodeAttrs;
  if (attrs.thematicBreak) return { type: "thematicBreak" };
  const depth = attrs.heading ? HEADING_PARSE_MAP[attrs.heading] : undefined;
  if (depth) return { type: "heading", depth, children: inlineOf(node) };
  if (attrs.style === "Code")
    return { type: "codeBlock", text: textOf(node), lang: attrs.codeLanguage ?? undefined };
  // A footnote paragraph round-trips as a definition instead of getting
  // bracket-escaped. Only mark-free text counts — a styled paragraph starting
  // with the same characters stays prose (escaping keeps its text intact).
  const m = FOOTNOTE_DEF.exec(textOf(node));
  if (m && node.content?.length && node.content.every((c) => c.type === "text")) {
    const prefix = `[^${m[1]}]: `;
    const children = inlineOf(node);
    const first = children[0];
    if (first?.type === "text") {
      if (first.text === prefix) children.shift();
      else if (first.text.startsWith(prefix))
        children[0] = { ...first, text: first.text.slice(prefix.length) };
    }
    return { type: "footnoteDefinition", label: m[1], children };
  }
  return { type: "paragraph", children: inlineOf(node) };
}

/** A paragraph's flat list attrs, or null when it isn't a list item. */
function listAttrsOf(
  node: JSONContent,
): { level: number; ordered: boolean; reference: string | null } | null {
  const attrs = (node.attrs ?? {}) as NodeAttrs;
  if (attrs.bullet) return { level: attrs.bullet.level ?? 0, ordered: false, reference: null };
  if (attrs.numbering) {
    return {
      level: attrs.numbering.level ?? 0,
      ordered: true,
      reference: attrs.numbering.reference ?? "",
    };
  }
  return null;
}

/** A GFM task checkbox as parseBlockNodes writes it (a standalone `[ ] ` run)
 *  or as a DOCX round-trip merges it (the prefix fused into the first run).
 *  Returns the checked state plus the paragraph stripped of the prefix, or
 *  null when the paragraph isn't a task item. */
function taskItemOf(para: JSONContent): { checked: boolean; body: JSONContent } | null {
  const first = para.content?.[0];
  if (first?.type !== "text") return null;
  for (const [prefix, checked] of [
    ["[ ] ", false],
    ["[x] ", true],
  ] as const) {
    if (first.text === prefix)
      return { checked, body: { ...para, content: para.content!.slice(1) } };
    if (first.text?.startsWith(prefix))
      return {
        checked,
        body: {
          ...para,
          content: [{ ...first, text: first.text.slice(prefix.length) }, ...para.content!.slice(1)],
        },
      };
  }
  return null;
}

/** A list's custom start (its own paragraph's, else the abstract
 *  definition's), or undefined when it numbers from 1. */
function listStartOf(para: JSONContent, defStarts: StartLookup): number | undefined {
  const attrs = listAttrsOf(para)!;
  const own = para.attrs?.numbering?.start;
  const fromDef = attrs.reference ? defStarts.get(attrs.reference)?.get(attrs.level) : undefined;
  const start = typeof own === "number" ? own : fromDef;
  return typeof start === "number" && start > 1 ? start : undefined;
}

/** Rebuild one nested list from a run of consecutive same-kind list
 *  paragraphs: level jumps open sub-lists under the previous item, level
 *  drops close them. Loose items collapse — flatness lost that when the tree
 *  was flattened. */
function listToBlock(group: JSONContent[], defStarts: StartLookup): MdBlock {
  const first = listAttrsOf(group[0])!;
  const rootStart = listStartOf(group[0], defStarts);
  const root: Extract<MdBlock, { type: "list" }> = {
    type: "list",
    ordered: first.ordered,
    ...(rootStart !== undefined ? { start: rootStart } : {}),
    items: [],
  };
  const stack: Extract<MdBlock, { type: "list" }>[] = [root];
  for (const para of group) {
    const { level } = listAttrsOf(para)!;
    while (stack.length > level + 1) stack.pop();
    // The checkbox rides as a text prefix (parseBlockNodes) — lift it back to
    // GFM task semantics so it exports unescaped.
    const task = taskItemOf(para);
    const item: Extract<MdBlock, { type: "list" }>["items"][number] = {
      blocks: [paragraphToBlock(task ? task.body : para)],
      ...(task ? { checked: task.checked } : {}),
    };
    if (stack.length === level + 1) {
      stack[stack.length - 1].items.push(item);
    } else {
      // Deeper level: a new sub-list hangs off the previous item.
      const parent = stack[stack.length - 1];
      const subStart = listStartOf(para, defStarts);
      const sub: Extract<MdBlock, { type: "list" }> = {
        type: "list",
        ordered: listAttrsOf(para)!.ordered,
        ...(subStart !== undefined ? { start: subStart } : {}),
        items: [item],
      };
      parent.items[parent.items.length - 1].blocks.push(sub);
      stack.push(sub);
    }
  }
  return root;
}

/** A maximal run of consecutive list paragraphs becomes list(s): adjacent
 *  items with a different kind (bullet vs ordered) or a different ordered
 *  reference split into separate lists, mirroring Markdown list semantics. */
function listRunsToBlocks(run: JSONContent[], defStarts: StartLookup): MdBlock[] {
  const out: MdBlock[] = [];
  let group: JSONContent[] = [];
  let key: string | null = null;
  const flush = (): void => {
    if (group.length > 0) {
      out.push(listToBlock(group, defStarts));
      group = [];
    }
  };
  for (const para of run) {
    const attrs = listAttrsOf(para)!;
    const k = attrs.ordered ? `n:${attrs.reference}` : "b";
    if (k !== key) {
      flush();
      key = k;
    }
    group.push(para);
  }
  flush();
  return out;
}

function tableToBlock(node: JSONContent): MdBlock | null {
  const rows = (node.content ?? []).filter((row) => row.type === "tableRow");
  if (rows.length === 0) return null;
  const align: MdAlignment[] = [];
  const irRows = rows.map((row, rowIndex) => {
    const header = rowIndex === 0 && !!(row.attrs as NodeAttrs)?.tableHeader;
    return {
      header,
      cells: (row.content ?? [])
        .filter((cell) => cell.type === "tableCell")
        .map((cell, col) => {
          if (rowIndex === 0) {
            const a = ((cell.content ?? [])[0]?.attrs ?? {}) as NodeAttrs;
            align[col] =
              a.alignment === "center" || a.alignment === "right" || a.alignment === "left"
                ? a.alignment
                : null;
          }
          return { header, blocks: (cell.content ?? []).flatMap(nodeToBlocks) };
        }),
    };
  });
  return { type: "table", align, rows: irRows };
}

function nodeToBlocks(node: JSONContent): MdBlock[] {
  switch (node.type) {
    case "paragraph":
      return [paragraphToBlock(node)];
    case "table": {
      const block = tableToBlock(node);
      return block ? [block] : [];
    }
    // Content-bearing containers (textbox / content controls / groups /
    // shapes): recurse so the text inside survives — Markdown has no box
    // semantics, only the content.
    case "textbox":
    case "sdtBlock":
    case "wpgGroup":
    case "wpsShape":
      return (node.content ?? []).flatMap(nodeToBlocks);
    default:
      // Page/section/column breaks, TOC fields, passthrough XML, charts —
      // no Markdown equivalent.
      return [];
  }
}

function toBlocks(doc: JSONContent): MdBlock[] {
  const out: MdBlock[] = [];
  const content = doc.content ?? [];
  // Custom numbering starts live in the abstract definitions for documents
  // parsed from DOCX (paragraphs carry them only for markdown imports).
  const defStarts: StartLookup = new Map();
  const defs = ((doc.attrs as NodeAttrs | undefined)?.numbering as NodeAttrs | undefined)
    ?.abstractNumberings as
    | { reference?: string; levels?: { level?: number; start?: number }[] }[]
    | undefined;
  for (const def of defs ?? []) {
    if (!def.reference) continue;
    for (const lvl of def.levels ?? []) {
      if (typeof lvl.start === "number" && lvl.start > 1) {
        if (!defStarts.has(def.reference)) defStarts.set(def.reference, new Map());
        if (!defStarts.get(def.reference)!.has(lvl.level ?? 0))
          defStarts.get(def.reference)!.set(lvl.level ?? 0, lvl.start);
      }
    }
  }
  for (let i = 0; i < content.length;) {
    const node = content[i];
    if (listAttrsOf(node)) {
      const run: JSONContent[] = [];
      while (i < content.length && listAttrsOf(content[i])) run.push(content[i++]);
      out.push(...listRunsToBlocks(run, defStarts));
      continue;
    }
    const attrs = (node.attrs ?? {}) as NodeAttrs;
    if (node.type === "paragraph" && attrs.style === "IntenseQuote") {
      const run: JSONContent[] = [];
      while (
        i < content.length &&
        content[i].type === "paragraph" &&
        (content[i].attrs as NodeAttrs)?.style === "IntenseQuote" &&
        !listAttrsOf(content[i])
      ) {
        run.push(content[i++]);
      }
      out.push({ type: "quote", blocks: run.map(paragraphToBlock) });
      continue;
    }
    out.push(...nodeToBlocks(node));
    i++;
  }
  // Footnote bodies live in documentExtras, outside the flow — append them at
  // the end so exporting keeps their content (Markdown's definition syntax
  // re-imports as plain footnote paragraphs).
  const extras = ((doc.attrs as NodeAttrs | undefined)?.documentExtras ?? {}) as NodeAttrs;
  const footnotes = (extras.footnotes ?? []) as { id?: number; children?: JSONContent[] }[];
  footnotes.forEach((note, index) => {
    const children = (note.children ?? []).flatMap((para) =>
      para.type === "paragraph" ? inlineOf(para) : [],
    );
    out.push({ type: "footnoteDefinition", label: String(note.id ?? index + 1), children });
  });
  return out;
}

// ── Mapper + bound exports ──

/** The DOCX schema's Markdown mapping — the reference `MarkdownMapper`. */
export const docxMarkdownMapper: MarkdownMapper<JSONContent> = {
  parseBlock: parseBlockNodes,
  parseInline,
  toBlocks,
};

/**
 * Parse Markdown string to Tiptap JSON.
 */
export function parseMarkdown(markdown: string): JSONContent {
  const content = parseMarkdownTokens(markdown).flatMap(parseBlockNodes);
  const doc: JSONContent = assignOrderedReferences({ type: "doc", content });
  // A custom list start rides the first item's paragraph (reference now
  // assigned) — lift the collected per-level starts into the reference's
  // abstractNum definition so compile registers them and Word numbers from
  // the imported value instead of restarting at 1.
  const levelStarts: StartLookup = new Map();
  for (const node of content) {
    const num = ((node.attrs ?? {}) as NodeAttrs).numbering as
      | { reference?: string; level?: number; start?: number }
      | undefined;
    if (typeof num?.start !== "number" || typeof num.reference !== "string") continue;
    if (!levelStarts.has(num.reference)) levelStarts.set(num.reference, new Map());
    const levels = levelStarts.get(num.reference)!;
    if (!levels.has(num.level ?? 0)) levels.set(num.level ?? 0, num.start);
  }
  if (levelStarts.size > 0) {
    doc.attrs = {
      numbering: {
        abstractNumberings: [...levelStarts].map(([reference, levels]) => ({
          reference,
          levels: buildOrderedLevels(LevelFormat.DECIMAL, levels),
        })),
      },
    };
  }
  return doc;
}

/**
 * Generate Markdown string from Tiptap JSON.
 */
export function generateMarkdown(doc: JSONContent): string {
  return renderMarkdownBlocks(toBlocks(doc));
}
