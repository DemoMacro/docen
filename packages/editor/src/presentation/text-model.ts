import type { TableCellOptions } from "@docen/pptx";
import type {
  BulletOptions,
  ParagraphDescriptorOptions,
  StrikeStyle,
  TextBodyOptions,
  TextRunOptions,
  UnderlineStyle,
} from "@office-open/core/drawing";

import { bodyParagraphsOf, cellParagraphsOf } from "./commands";

/** The shape/cell source accepted by every in-place text operation. */
export type TextModelSource =
  | { kind: "shape"; shape: { textBody?: { paragraphs?: unknown } | undefined } }
  | { kind: "cell"; cell: TableCellOptions };

/** One contiguous styled piece of an editable paragraph. */
interface StyledPiece {
  text: string;
  style?: Omit<TextRunOptions, "text">;
  source?: unknown;
}

/** One editable paragraph: its OOXML properties plus styled text pieces. */
interface EditableLine {
  properties?: ParagraphDescriptorOptions["properties"];
  pieces: StyledPiece[];
}

/** A run style without its text payload — safe to clone onto insertions. */
type RunStyle = StyledPiece["style"];

/** The paragraph model behind one source, normalized in place. */
function paragraphsOf(source: TextModelSource): ParagraphDescriptorOptions[] {
  if (source.kind === "cell") return cellParagraphsOf(source.cell);
  return bodyParagraphsOf(source.shape.textBody as TextBodyOptions);
}

/** Split a paragraph's children into editable styled pieces. Non-text children
 *  (breaks and fields) ride through as zero-width source anchors instead of
 *  being dropped by a plain-text rewrite. */
function editableLines(paragraphs: ParagraphDescriptorOptions[]): EditableLine[] {
  return paragraphs.map((paragraph) => ({
    properties: paragraph.properties ? { ...paragraph.properties } : undefined,
    pieces: (paragraph.children ?? []).map((child): StyledPiece => {
      if (
        typeof child === "object" &&
        child !== null &&
        ("break" in child || ("type" in child && typeof child.type === "string"))
      ) {
        return { text: "", source: child };
      }
      if (typeof child === "object" && child !== null && "text" in child) {
        const candidate = child as { text?: unknown };
        if (typeof candidate.text === "string" && candidate.text.length > 0) {
          const { text: _text, ...style } = child as TextRunOptions;
          return { text: candidate.text, style };
        }
      }
      return { text: "", source: child };
    }),
  }));
}

/** Rebuild descriptor paragraphs, merging adjacent pieces that carry the same
 *  style so ordinary typing does not create one run per character. */
function paragraphsFromLines(lines: EditableLine[]): ParagraphDescriptorOptions[] {
  return lines.map((line) => {
    const children: NonNullable<ParagraphDescriptorOptions["children"]> = [];
    for (const piece of line.pieces) {
      if (piece.source != null) {
        children.push(piece.source as NonNullable<ParagraphDescriptorOptions["children"]>[number]);
        continue;
      }
      if (!piece.text) continue;
      const previous = children.at(-1);
      if (
        previous &&
        typeof previous === "object" &&
        "text" in previous &&
        typeof previous.text === "string" &&
        sameStyle(previous, piece.style)
      )
        previous.text += piece.text;
      else children.push({ ...piece.style, text: piece.text });
    }
    return {
      ...(line.properties ? { properties: { ...line.properties } } : {}),
      children,
    };
  });
}

function sameStyle(left: TextRunOptions, right: RunStyle): boolean {
  const { text: _left, ...leftStyle } = left;
  return JSON.stringify(leftStyle) === JSON.stringify(right ?? {});
}

function plainTextOfLines(lines: EditableLine[]): string {
  return lines.map((line) => line.pieces.reduce((text, piece) => text + piece.text, "")).join("\n");
}

function locate(lines: EditableLine[], position: number): { line: number; offset: number } {
  let remaining = Math.max(0, position);
  for (let line = 0; line < lines.length; line++) {
    const length = lines[line]!.pieces.reduce((sum, piece) => sum + piece.text.length, 0);
    if (remaining <= length) return { line, offset: remaining };
    remaining -= length;
  }
  const last = lines.at(-1);
  return {
    line: Math.max(0, lines.length - 1),
    offset: last ? last.pieces.reduce((sum, piece) => sum + piece.text.length, 0) : 0,
  };
}

function pieceStyleAt(lines: EditableLine[], position: number): RunStyle {
  const { line: lineIndex } = locate(lines, position);
  const line = lines[lineIndex];
  let consumed = 0;
  for (const piece of line?.pieces ?? []) {
    const end = consumed + piece.text.length;
    if (position > consumed && position <= end) return piece.style;
    if (position === consumed && piece.style) return piece.style;
    consumed = end;
  }
  const before = line?.pieces.at(-1)?.style;
  return before ?? line?.pieces[0]?.style;
}

function piecesBefore(line: EditableLine, offset: number): StyledPiece[] {
  const kept: StyledPiece[] = [];
  let consumed = 0;
  for (const piece of line.pieces) {
    if (piece.source != null && offset === consumed) {
      kept.push(piece);
      consumed += piece.text.length;
      continue;
    }
    const sliceEnd = Math.min(piece.text.length, Math.max(0, offset - consumed));
    const text = piece.text.slice(0, sliceEnd);
    if (text) kept.push({ ...piece, text });
    consumed += piece.text.length;
  }
  return kept;
}

function piecesAfter(line: EditableLine, offset: number): StyledPiece[] {
  const kept: StyledPiece[] = [];
  let consumed = 0;
  for (const piece of line.pieces) {
    const sliceStart = Math.min(piece.text.length, Math.max(0, offset - consumed));
    const text = piece.text.slice(sliceStart);
    if (text) kept.push({ ...piece, text });
    consumed += piece.text.length;
  }
  return kept;
}

function removeRange(lines: EditableLine[], start: number, end: number): EditableLine[] {
  const from = locate(lines, start);
  const to = locate(lines, end);
  const first = lines[from.line];
  const last = lines[to.line];
  if (!first || !last) return [...lines];

  const before = piecesBefore(first, from.offset);
  const after = piecesAfter(last, to.offset);
  const merged: EditableLine = { properties: first.properties, pieces: [...before, ...after] };
  if (from.line === to.line) {
    return lines.toSpliced(from.line, 1, merged);
  }
  return lines.toSpliced(from.line, to.line - from.line + 1, merged);
}

function insertText(
  lines: EditableLine[],
  position: number,
  text: string,
  inheritedStyle?: RunStyle,
): EditableLine[] {
  if (!text) return [...lines];
  const { line: lineIndex, offset } = locate(lines, position);
  const line = lines[lineIndex];
  if (!line) return [...lines];
  const before = piecesBefore(line, offset);
  const after = piecesAfter(line, offset);
  const style = before.at(-1)?.style ?? after[0]?.style ?? inheritedStyle;
  const segments = text.split("\n");
  const replacement: EditableLine[] = segments.map((segment, index) => ({
    properties: line.properties ? { ...line.properties } : undefined,
    pieces: [
      ...(index === 0 ? before : []),
      ...(segment ? [{ text: segment, style: style ? { ...style } : undefined }] : []),
      ...(index === segments.length - 1 ? after : []),
    ],
  }));
  return lines.toSpliced(lineIndex, 1, ...replacement);
}

/** The exact plain text the textarea edits. */
export function textModelOf(source: TextModelSource): string {
  return plainTextOfLines(editableLines(paragraphsOf(source)));
}

/** Apply one textarea edit without flattening styled runs. The caller supplies
 *  the previous text so the changed range can be isolated by common affixes. */
export function writeTextModel(source: TextModelSource, next: string, previous: string): void {
  let start = 0;
  const shared = Math.min(previous.length, next.length);
  while (start < shared && previous[start] === next[start]) start++;
  let previousEnd = previous.length;
  let nextEnd = next.length;
  while (
    previousEnd > start &&
    nextEnd > start &&
    previous[previousEnd - 1] === next[nextEnd - 1]
  ) {
    previousEnd--;
    nextEnd--;
  }
  if (start === previousEnd && start === nextEnd) return;

  const paragraphs = paragraphsOf(source);
  const lines = editableLines(paragraphs);
  const inheritedStyle = pieceStyleAt(lines, start);
  const removed = removeRange(lines, start, previousEnd);
  const updatedLines = insertText(removed, start, next.slice(start, nextEnd), inheritedStyle);
  const updated = paragraphsFromLines(updatedLines);
  paragraphs.length = 0;
  paragraphs.push(...updated);
}

/** Styles intersected by the selection; a collapsed caret selects the style
 *  typing will inherit at that insertion point. */
function selectedStyles(lines: EditableLine[], start: number, end: number): RunStyle[] {
  const styles: RunStyle[] = [];
  let consumed = 0;
  for (const line of lines) {
    for (const piece of line.pieces) {
      const pieceEnd = consumed + piece.text.length;
      const selected =
        end > start ? pieceEnd > start && consumed < end : start >= consumed && start <= pieceEnd;
      if (selected) styles.push(piece.style);
      consumed = pieceEnd;
    }
    consumed++; // The newline between paragraphs.
  }
  return styles;
}

function selectedPieces(lines: EditableLine[], start: number, end: number): StyledPiece[] {
  const pieces: StyledPiece[] = [];
  let consumed = 0;
  for (const line of lines) {
    for (const piece of line.pieces) {
      const pieceEnd = consumed + piece.text.length;
      const selected =
        end > start ? pieceEnd > start && consumed < end : start >= consumed && start <= pieceEnd;
      if (selected) pieces.push(piece);
      consumed = pieceEnd;
    }
    consumed++;
  }
  return pieces;
}

function selectedLines(lines: EditableLine[], start: number, end: number): EditableLine[] {
  const from = locate(lines, start);
  const to = locate(lines, end);
  return lines.slice(from.line, to.line + 1);
}

function splitLinesAt(lines: EditableLine[], position: number): void {
  let consumed = 0;
  for (const line of lines) {
    for (let index = 0; index < line.pieces.length; index++) {
      const piece = line.pieces[index]!;
      const end = consumed + piece.text.length;
      if (position > consumed && position < end) {
        const before = { ...piece, text: piece.text.slice(0, position - consumed) };
        const after = { ...piece, text: piece.text.slice(position - consumed) };
        line.pieces = line.pieces.toSpliced(index, 1, before, after);
        return;
      }
      consumed = end;
    }
    consumed++;
  }
}

/** Apply a ribbon text command to the textarea's selected range (or insertion
 *  point). Returns false when the command had no observable effect. */
export function formatTextModel(
  source: TextModelSource,
  start: number,
  end: number,
  name: string,
  value?: string,
): boolean {
  const paragraphs = paragraphsOf(source);
  const lines = editableLines(paragraphs);
  const from = Math.min(start, end);
  const to = Math.max(start, end);
  splitLinesAt(lines, from);
  splitLinesAt(lines, to);
  const styles = selectedStyles(lines, from, to);
  const targets = selectedLines(lines, from, to);
  if (!targets.length) return false;
  const pieces = selectedPieces(lines, from, to);
  const editStyles = (action: (styled: Record<string, unknown>) => void): void => {
    for (const piece of pieces) {
      const styled: Record<string, unknown> = { ...piece.style };
      action(styled);
      piece.style = styled;
    }
  };

  switch (name) {
    case "bold": {
      const on = styles.length > 0 && styles.every((style) => style?.bold === true);
      editStyles((styled) => (on ? delete styled.bold : (styled.bold = true)));
      break;
    }
    case "italic": {
      const on = styles.length > 0 && styles.every((style) => style?.italic === true);
      editStyles((styled) => (on ? delete styled.italic : (styled.italic = true)));
      break;
    }
    case "underline": {
      const on = styles.every(
        (style) => style?.underline !== undefined && style.underline !== "none",
      );
      editStyles((styled) =>
        on ? delete styled.underline : (styled.underline = "single" satisfies UnderlineStyle),
      );
      break;
    }
    case "strike": {
      const on = styles.every(
        (style) => style?.strike !== undefined && style.strike !== "noStrike",
      );
      editStyles((styled) =>
        on ? delete styled.strike : (styled.strike = "singleStrike" satisfies StrikeStyle),
      );
      break;
    }
    case "font-face": {
      if (!value) return false;
      editStyles((styled) => (styled.font = value));
      break;
    }
    case "font-size": {
      const size = Number(value);
      if (!Number.isFinite(size) || size <= 0) return false;
      editStyles((styled) => (styled.size = size));
      break;
    }
    default: {
      const alignment =
        name === "align-left"
          ? "left"
          : name === "align-center"
            ? "center"
            : name === "align-right"
              ? "right"
              : name === "justify"
                ? "justify"
                : undefined;
      const bullet =
        name === "list"
          ? ({ type: "char" } as BulletOptions)
          : name === "numbering"
            ? ({ type: "autoNum" } as BulletOptions)
            : undefined;
      const spacing = (
        { "1.0": 100, "1.5": 150, "2.0": 200, "2.5": 250, "3.0": 300 } as Record<string, number>
      )[value ?? ""];
      if (!alignment && !bullet && name !== "line-spacing") return false;
      for (const line of targets) {
        const properties: Record<string, unknown> = { ...line.properties };
        if (alignment) properties.alignment = alignment;
        if (bullet) {
          const existing = properties.bullet as BulletOptions | undefined;
          if (name === "list" || name === "numbering") {
            const on = targets.every((target) => {
              const value = target.properties?.bullet;
              return value != null && value.type !== "none" && value.type === existing?.type;
            });
            if (on) delete properties.bullet;
            else properties.bullet = { ...bullet };
          }
        }
        if (name === "line-spacing" && spacing != null) properties.lineSpacingPercent = spacing;
        line.properties = properties as ParagraphDescriptorOptions["properties"];
      }
    }
  }

  const updated = paragraphsFromLines(lines);
  if (JSON.stringify(updated) === JSON.stringify(paragraphs)) return false;
  paragraphs.length = 0;
  paragraphs.push(...updated);
  return true;
}
