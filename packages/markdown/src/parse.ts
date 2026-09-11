import { lexer, type Token, type Tokens } from "marked";

import type { MdAlignment, MdBlock, MdInline, MdListItem, MdMark, MdTableRow } from "./types";

const withMark = (inlines: MdInline[], mark: MdMark): MdInline[] =>
  inlines.map((n) => (n.type === "text" ? { ...n, marks: [...(n.marks ?? []), mark] } : n));

/** A GFM footnote definition line (`[^label]: body`). */
const FOOTNOTE_DEF = /^\[\^([^\]\s]+)\]:\s*(.*)$/;

/** Run marked's inline rules on a bare substring (a footnote body) by lexing
 *  it as a one-paragraph document. */
function parseMarkdownInline(text: string): MdInline[] {
  const [first] = lexer(text);
  if (first?.type === "paragraph") return parseInlineTokens((first as Tokens.Paragraph).tokens);
  return text ? [{ type: "text", text }] : [];
}

function parseInlineTokens(tokens: Token[] | undefined, text?: string): MdInline[] {
  if (!tokens || tokens.length === 0) {
    return text ? [{ type: "text", text }] : [];
  }
  const out: MdInline[] = [];
  for (const token of tokens) {
    switch (token.type) {
      case "text": {
        // A nested text token (list items, emphasis) carries its own inline
        // subtree — flatten it, the marks already sit on the children.
        const t = token as Tokens.Text;
        out.push(...parseInlineTokens(t.tokens ?? [], t.text));
        break;
      }
      case "escape":
        out.push({ type: "text", text: (token as Tokens.Escape).text });
        break;
      case "strong":
        out.push(...withMark(parseInlineTokens((token as Tokens.Strong).tokens), { type: "bold" }));
        break;
      case "em":
        out.push(...withMark(parseInlineTokens((token as Tokens.Em).tokens), { type: "italic" }));
        break;
      case "del":
        out.push(...withMark(parseInlineTokens((token as Tokens.Del).tokens), { type: "strike" }));
        break;
      case "codespan":
        out.push({
          type: "text",
          text: (token as Tokens.Codespan).text,
          marks: [{ type: "code" }],
        });
        break;
      case "link": {
        const link = token as Tokens.Link;
        out.push(
          ...withMark(parseInlineTokens(link.tokens, link.text), {
            type: "link",
            href: link.href,
            title: link.title || undefined,
          }),
        );
        break;
      }
      case "image": {
        const image = token as Tokens.Image;
        out.push({
          type: "image",
          src: image.href,
          alt: image.text || undefined,
          title: image.title || undefined,
        });
        break;
      }
      case "br":
        out.push({ type: "hardBreak" });
        break;
      case "html":
        // No HTML pipeline exists — keep the literal text so content survives.
        out.push({ type: "text", text: (token as Tokens.Tag).text });
        break;
      default:
        break;
    }
  }
  return out;
}

function parseListItems(items: Tokens.ListItem[]): MdListItem[] {
  return items.map((item) => ({
    // A loose item lexes one `paragraph` token per block — every one becomes
    // an entry block; nested lists recurse as-is.
    blocks: (item.tokens ?? [])
      .map(parseBlockToken)
      .filter((b): b is MdBlock | MdBlock[] => b != null)
      .flat(),
    ...(item.task ? { checked: !!item.checked } : {}),
  }));
}

function parseTable(token: Tokens.Table): MdBlock {
  const parseCell = (cell: Tokens.TableCell, header: boolean) => ({
    header: header || undefined,
    blocks: [
      { type: "paragraph", children: parseInlineTokens(cell.tokens ?? [], cell.text) },
    ] as MdBlock[],
  });
  const rows: MdTableRow[] = [
    { cells: token.header.map((cell) => parseCell(cell, true)) },
    ...token.rows.map((row) => ({ cells: row.map((cell) => parseCell(cell, false)) })),
  ];
  const align = token.align.map((a): MdAlignment =>
    a === "left" || a === "center" || a === "right" ? a : null,
  );
  return { type: "table", align, rows };
}

function parseBlockToken(token: Token): MdBlock | MdBlock[] | null {
  switch (token.type) {
    case "heading": {
      const h = token as Tokens.Heading;
      return { type: "heading", depth: h.depth, children: parseInlineTokens(h.tokens, h.text) };
    }
    case "paragraph": {
      const para = token as Tokens.Paragraph;
      // GFM footnote definitions (`[^1]: note`) lex as plain paragraphs —
      // lift each definition line out so it round-trips as its own block
      // instead of getting bracket-escaped on the way back.
      const lines = para.raw.split("\n");
      if (lines.length > 0 && lines.every((l) => FOOTNOTE_DEF.test(l))) {
        return lines.map((line) => {
          const m = FOOTNOTE_DEF.exec(line)!;
          return {
            type: "footnoteDefinition",
            label: m[1],
            // Re-parse the body so escapes/emphasis inside survive verbatim.
            children: parseMarkdownInline(m[2]),
          } satisfies MdBlock;
        });
      }
      return { type: "paragraph", children: parseInlineTokens(para.tokens) };
    }
    case "space":
      return null;
    case "hr":
      return { type: "thematicBreak" };
    case "code": {
      const code = token as Tokens.Code;
      return { type: "codeBlock", text: code.text, lang: code.lang || undefined };
    }
    case "blockquote": {
      const quote = token as Tokens.Blockquote;
      const blocks = (quote.tokens ?? [])
        .map(parseBlockToken)
        .filter((b): b is MdBlock | MdBlock[] => b != null)
        .flat();
      return { type: "quote", blocks };
    }
    case "list": {
      const list = token as Tokens.List;
      return {
        type: "list",
        ordered: !!list.ordered,
        start:
          list.ordered && typeof list.start === "number" && list.start !== 1
            ? list.start
            : undefined,
        items: parseListItems(list.items ?? []),
      };
    }
    case "table":
      return parseTable(token as Tokens.Table);
    case "text":
      // Bare top-level text lexes as a `text` token, not a paragraph.
      return {
        type: "paragraph",
        children: parseInlineTokens((token as Tokens.Text).tokens ?? [], token.raw),
      };
    case "html":
      // No HTML pipeline exists — keep the raw markup as literal text so the
      // content survives instead of being dropped.
      return {
        type: "paragraph",
        children: [{ type: "text", text: token.raw.trim() }],
      };
    case "def":
      // A footnote definition lexes as a link reference definition (its
      // `[label]: text` shape) — lift it into its own block instead of losing
      // the note. Plain link definitions carry no visible content of their
      // own (inline links already embed their href), so they stay dropped.
      return parseFootnoteDef(token.raw);
    default:
      // Unknown tokens have no IR — dropped by design.
      return null;
  }
}

/** Parse Markdown into the neutral IR block list. */
export function parseMarkdownTokens(markdown: string): MdBlock[] {
  return lexer(markdown)
    .map(parseBlockToken)
    .filter((b): b is MdBlock | MdBlock[] => b != null)
    .flat();
}

/** One footnote definition line, or null for any other definition text. */
function parseFootnoteDef(raw: string): MdBlock | null {
  const m = FOOTNOTE_DEF.exec(raw);
  if (!m) return null;
  return {
    type: "footnoteDefinition",
    label: m[1],
    // Re-parse the body so escapes/emphasis inside survive verbatim.
    children: parseMarkdownInline(m[2]),
  };
}
