import type { MdBlock, MdInline, MdMark, MdTableRow } from "./types";

/** Characters that always change structure mid-line when literal. */
const ALWAYS_SPECIAL = /([\\`*_[\]<>])/g;
/** Characters that only change structure at the start of a line. */
const LINE_START_SPECIAL = /^([#+>~!+-])/;

function escapeText(text: string, inTable: boolean, startOfLine: boolean): string {
  let out = text.replace(ALWAYS_SPECIAL, "\\$1");
  if (inTable) out = out.replace(/\|/g, "\\|");
  if (startOfLine) {
    out = out.replace(LINE_START_SPECIAL, "\\$1");
    // "1. " at line start would lex as an ordered list item.
    out = out.replace(/^(\d+)([.)])(\s|$)/, "$1\\$2$3");
  }
  return out;
}

/** Nesting order of emphasis marks: the link wrapper outermost, then the
 *  delimiter pairs from widest to narrowest. The opening/closing algorithm
 *  walks this order so interleaved runs nest consistently. */
const MARK_ORDER = ["link", "bold", "italic", "strike"] as const;

const sortMarks = (marks: MdMark[]): MdMark[] =>
  [...marks].sort(
    (a, b) =>
      MARK_ORDER.indexOf(a.type as (typeof MARK_ORDER)[number]) -
      MARK_ORDER.indexOf(b.type as (typeof MARK_ORDER)[number]),
  );

const openDelim = (mark: MdMark): string =>
  mark.type === "link" ? "[" : mark.type === "bold" ? "**" : mark.type === "italic" ? "*" : "~~";

function closeDelim(mark: MdMark): string {
  if (mark.type === "link") {
    const title = mark.title ? ` "${mark.title.replace(/"/g, '\\"')}"` : "";
    return `](${/[<>\s()]/.test(mark.href) ? `<${mark.href}>` : mark.href}${title})`;
  }
  return openDelim(mark);
}

function renderCodeSpan(text: string, inTable: boolean): string {
  let content = text;
  if (inTable) content = content.replace(/\|/g, "\\|");
  const longest = Math.max(0, ...[...content.matchAll(/`+/g)].map((r) => r[0].length));
  const fence = "`".repeat(longest + 1);
  const padded = content.startsWith("`") || content.endsWith("`") ? ` ${content} ` : content;
  return `${fence}${padded}${fence}`;
}

/**
 * Render inline nodes. Marks are opened and closed across runs by comparing
 * each run's mark set with its neighbors' (lookahead), so split runs of one
 * emphasis group stay inside one pair of delimiters, and edge whitespace
 * always lands outside the delimiters (`**bold** ` never `**bold **`).
 */
function renderInline(inlines: MdInline[], inTable: boolean): string {
  const runs = inlines.filter((n) => n.type !== "hardBreak" || !inTable);
  let out = "";
  let startOfLine = true;
  let active: MdMark[] = [];

  const openNew = (marks: MdMark[]): void => {
    for (const mark of sortMarks(marks)) {
      if (
        !active.some(
          (m) =>
            m.type === mark.type &&
            (mark.type !== "link" || m.type !== "link" || m.href === mark.href),
        )
      ) {
        out += openDelim(mark);
        active.push(mark);
      }
    }
  };
  const closeGone = (next: MdMark[]): void => {
    while (active.length) {
      const top = active[active.length - 1];
      const kept = next.some(
        (m) =>
          m.type === top.type && (top.type !== "link" || m.type !== "link" || m.href === top.href),
      );
      if (!kept) {
        out += closeDelim(active.pop() as MdMark);
      } else break;
    }
  };

  for (let i = 0; i < runs.length; i++) {
    const node = runs[i];
    if (node.type === "hardBreak") {
      closeGone([]);
      active = [];
      out += "  \n";
      startOfLine = true;
      continue;
    }
    if (node.type === "image") {
      closeGone([]);
      active = [];
      const alt = node.alt ? escapeText(node.alt, inTable, false) : "";
      const src = /[\s()]/.test(node.src) ? `<${node.src}>` : node.src;
      const title = node.title ? ` "${node.title.replace(/"/g, '\\"')}"` : "";
      out += `![${alt}](${src}${title})`;
      startOfLine = false;
      continue;
    }

    const marks = node.marks ?? [];
    // Inline code wraps its content alone; the active group closes first.
    if (marks.some((m) => m.type === "code")) {
      closeGone([]);
      active = [];
      out += renderCodeSpan(node.text, inTable);
      startOfLine = false;
      continue;
    }

    const next =
      i + 1 < runs.length
        ? runs[i + 1].type === "text"
          ? ((runs[i + 1] as Extract<MdInline, { type: "text" }>).marks ?? [])
          : []
        : [];
    const target = sortMarks(marks.filter((m) => m.type !== "code"));
    openNew(target);
    const lead = node.text.match(/^\s*/)?.[0] ?? "";
    const trail = node.text.match(/\s*$/)?.[0] ?? "";
    const core = node.text.slice(lead.length, node.text.length - trail.length);
    if (lead) {
      out += escapeText(lead, inTable, startOfLine);
      startOfLine = false;
    }
    if (core) {
      out += escapeText(core, inTable, startOfLine && !lead);
      startOfLine = false;
    }
    // Close what the next run no longer shares before trailing whitespace so
    // the delimiters hug the content (`**bold** ` not `**bold **`).
    closeGone(
      target.filter((m) =>
        next.some(
          (n) => n.type === m.type && (m.type !== "link" || n.type !== "link" || n.href === m.href),
        ),
      ),
    );
    if (trail) out += escapeText(trail, inTable, false);
  }
  closeGone([]);
  return out;
}

function renderTableCell(cell: MdTableRow["cells"][number]): string {
  const content = cell.blocks
    .map((b) => renderBlockLines(b, { inTable: true }).join(" "))
    .join(" ");
  return ` ${content || " "} `;
}

function renderTable(block: Extract<MdBlock, { type: "table" }>): string[] {
  const { align, rows } = block;
  // Ragged rows (a split merged cell renders an empty continuation cell) pad
  // to the widest row so the GFM grid stays rectangular.
  const cols = Math.max(align.length, ...rows.map((row) => row.cells.length));
  const cellsOf = (row: MdTableRow) => {
    const cells = row.cells.map((cell) => renderTableCell(cell));
    while (cells.length < cols) cells.push("   ");
    return cells;
  };
  const lines = [`|${cellsOf(rows[0]).join("|")}|`];
  const divider = Array.from({ length: cols }, (_, i) =>
    align[i] === "center"
      ? " :-: "
      : align[i] === "right"
        ? " ---: "
        : align[i] === "left"
          ? " :--- "
          : " --- ",
  );
  lines.push(`|${divider.join("|")}|`);
  for (const row of rows.slice(1)) lines.push(`|${cellsOf(row).join("|")}|`);
  return lines;
}

function renderList(block: Extract<MdBlock, { type: "list" }>, level: number): string[] {
  const indent = " ".repeat(4 * level);
  const lines: string[] = [];
  let counter = block.start ?? 1;
  for (const item of block.items) {
    const marker = block.ordered ? `${counter++}.` : "-";
    // GFM task items keep their checkbox in the marker position.
    const task = item.checked === undefined ? "" : item.checked ? "[x] " : "[ ] ";
    const hang = " ".repeat(marker.length + 1 + task.length);
    item.blocks.forEach((sub, index) => {
      if (sub.type === "list") {
        lines.push(...renderList(sub, level + 1));
        return;
      }
      const body = renderBlockLines(sub);
      body.forEach((line, lineIndex) => {
        if (index === 0 && lineIndex === 0) {
          lines.push(`${indent}${marker} ${task}${line}`.trimEnd());
        } else {
          // Block continuation (code, nested paragraphs) re-indents under the
          // marker so the item keeps holding it.
          lines.push(`${indent}${hang}${line}`.trimEnd());
        }
      });
    });
  }
  return lines;
}

function renderBlockLines(block: MdBlock, opts: { inTable?: boolean } = {}): string[] {
  switch (block.type) {
    case "heading":
      return [
        `${"#".repeat(Math.min(6, Math.max(1, block.depth)))} ${renderInline(block.children, false)}`,
      ];
    case "paragraph":
      return [renderInline(block.children, !!opts.inTable)];
    case "codeBlock": {
      const longest = Math.max(0, ...[...block.text.matchAll(/`+/g)].map((r) => r[0].length));
      const fence = "`".repeat(Math.max(3, longest + 1));
      return [`${fence}${block.lang ?? ""}\n${block.text}\n${fence}`];
    }
    case "quote": {
      // Soft breaks inside a paragraph span lines — prefix every line so the
      // quote never relies on lazy continuation.
      const inner = block.blocks.map((b) => renderBlockLines(b).join("\n")).join("\n\n");
      return [
        inner
          .split("\n")
          .map((l) => (l === "" ? ">" : `> ${l}`))
          .join("\n"),
      ];
    }
    case "list":
      return renderList(block, 0);
    case "table":
      return renderTable(block);
    case "thematicBreak":
      return ["---"];
    case "footnoteDefinition":
      // The label renders verbatim — it IS the syntax; only the body escapes.
      return [`[^${block.label}]: ${renderInline(block.children, !!opts.inTable)}`.trimEnd()];
    default:
      return [];
  }
}

/** Render IR blocks back to a Markdown string. */
export function renderMarkdownBlocks(blocks: MdBlock[]): string {
  return blocks.map((b) => renderBlockLines(b).join("\n")).join("\n\n");
}
