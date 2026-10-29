import type { ChartTitleOptions } from "@docen/docx";
import type {
  ParagraphDescriptorOptions,
  TextBodyOptions,
  TextRunOptions,
} from "@office-open/core/drawing";

function paragraphTextOf(paragraph: ParagraphDescriptorOptions | string): string {
  if (typeof paragraph === "string") return paragraph;
  const children = paragraph.children ?? [];
  return children
    .map((child) => {
      if (typeof child === "string") return child;
      return "text" in child && typeof child.text === "string" ? child.text : "";
    })
    .join("");
}

function bodyTextOf(body: string | TextBodyOptions): string {
  if (typeof body === "string") return body;
  if (body.text !== undefined) return body.text;
  return (body.paragraphs ?? []).map(paragraphTextOf).join("");
}

function isTextRun(
  child: NonNullable<ParagraphDescriptorOptions["children"]>[number],
): child is TextRunOptions {
  return typeof child !== "string" && !("type" in child || "break" in child || "content" in child);
}

function bodyWithText(body: string | TextBodyOptions, text: string): string | TextBodyOptions {
  if (typeof body === "string") return text;

  const [first] = body.paragraphs ?? [];
  if (first === undefined) return typeof body === "string" ? text : { ...body, text };

  const firstParagraph: ParagraphDescriptorOptions = typeof first === "string" ? {} : first;
  const styledRun = (firstParagraph.children ?? []).find(isTextRun);
  const children: NonNullable<ParagraphDescriptorOptions["children"]> = styledRun
    ? [{ ...styledRun, text }]
    : [text];

  return {
    ...body,
    paragraphs: [
      {
        ...firstParagraph,
        children,
        text: undefined,
      },
    ],
  };
}

/** Flatten a chart title's rich body to the text shown by the data grid. */
export function chartTitleTextOf(title: string | ChartTitleOptions): string {
  return typeof title === "string" ? title : bodyTextOf(title.text ?? "");
}

/** Replace a chart title's visible text while keeping its layout and styling. */
export function chartTitleWithText(
  title: string | ChartTitleOptions,
  text: string,
): string | ChartTitleOptions {
  return typeof title === "string"
    ? text
    : { ...title, text: bodyWithText(title.text ?? "", text) };
}
