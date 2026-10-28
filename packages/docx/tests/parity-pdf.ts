// Executable Word-oracle comparison. A parity fixture may carry an
// `oracle.pdf` exported by the pinned Word/WPS build; after the same model is
// exported from docen, this tool rasterizes both PDFs and reports page-level
// pixel parity. Rendering engines differ in font hinting and antialiasing, so
// the gate compares a bounded changed-pixel ratio instead of exact bytes.

import { execFile } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import { promisify } from "node:util";

import { PNG } from "pngjs";

const run = promisify(execFile);

/** Default 144 DPI balances readable PDF text against CI runtime. */
export const DEFAULT_DPI = 144;
/** Per-RGB-channel tolerance; isolates antialiasing noise from real geometry. */
export const DEFAULT_CHANNEL_TOLERANCE = 16;
/** Default changed-pixel budget for a page. Tighten per feature with oracle data. */
export const DEFAULT_MAX_CHANGED_RATIO = 0.002;

export interface PdfParityOptions {
  dpi?: number;
  channelTolerance?: number;
  maxChangedRatio?: number;
  pdftoppm?: string;
}

export interface PdfParityPage {
  page: number;
  width: number;
  height: number;
  changedPixels: number;
  changedRatio: number;
  reason?: "dimension-mismatch";
}

export interface PdfParityReport {
  expected: string;
  actual: string;
  dpi: number;
  channelTolerance: number;
  maxChangedRatio: number;
  pageCount: number;
  changedPixels: number;
  changedRatio: number;
  pages: PdfParityPage[];
}

/** Compare two PDFs through Poppler and return deterministic page metrics. */
export async function comparePdfFiles(
  expectedPdf: string,
  actualPdf: string,
  options: PdfParityOptions = {},
): Promise<PdfParityReport> {
  const dpi = options.dpi ?? DEFAULT_DPI;
  const channelTolerance = options.channelTolerance ?? DEFAULT_CHANNEL_TOLERANCE;
  const maxChangedRatio = options.maxChangedRatio ?? DEFAULT_MAX_CHANGED_RATIO;
  const temporaryDirectory = mkdtempSync(join(tmpdir(), "docen-parity-"));

  try {
    const expectedPages = await renderPdf(
      expectedPdf,
      join(temporaryDirectory, "expected"),
      dpi,
      options.pdftoppm,
    );
    const actualPages = await renderPdf(
      actualPdf,
      join(temporaryDirectory, "actual"),
      dpi,
      options.pdftoppm,
    );
    const pageCount = Math.max(expectedPages.length, actualPages.length);
    const pages: PdfParityPage[] = [];

    for (let index = 0; index < pageCount; index++) {
      const expectedPath = expectedPages[index];
      const actualPath = actualPages[index];
      if (!expectedPath || !actualPath) {
        pages.push({
          page: index + 1,
          width: 0,
          height: 0,
          changedPixels: Number.MAX_SAFE_INTEGER,
          changedRatio: 1,
          reason: "dimension-mismatch",
        });
        continue;
      }
      pages.push(
        comparePng(
          readFileSync(expectedPath),
          readFileSync(actualPath),
          channelTolerance,
          index + 1,
        ),
      );
    }

    const changedPixels = pages.reduce((total, page) => total + page.changedPixels, 0);
    const totalPixels = pages.reduce((total, page) => total + page.width * page.height, 0);
    return {
      expected: basename(expectedPdf),
      actual: basename(actualPdf),
      dpi,
      channelTolerance,
      maxChangedRatio,
      pageCount,
      changedPixels,
      changedRatio: totalPixels === 0 ? 1 : changedPixels / totalPixels,
      pages,
    };
  } finally {
    rmSync(temporaryDirectory, { recursive: true, force: true });
  }
}

/** Compare decoded PNG buffers. Size drift is itself a hard geometry failure. */
export function comparePng(
  expectedPng: Uint8Array,
  actualPng: Uint8Array,
  channelTolerance = DEFAULT_CHANNEL_TOLERANCE,
  page = 1,
): PdfParityPage {
  const expected = PNG.sync.read(Buffer.from(expectedPng));
  const actual = PNG.sync.read(Buffer.from(actualPng));
  if (expected.width !== actual.width || expected.height !== actual.height) {
    return {
      page,
      width: actual.width,
      height: actual.height,
      changedPixels: Number.MAX_SAFE_INTEGER,
      changedRatio: 1,
      reason: "dimension-mismatch",
    };
  }

  let changedPixels = 0;
  for (let offset = 0; offset < expected.data.length; offset += 4) {
    const redDelta = Math.abs(expected.data[offset]! - actual.data[offset]!);
    const greenDelta = Math.abs(expected.data[offset + 1]! - actual.data[offset + 1]!);
    const blueDelta = Math.abs(expected.data[offset + 2]! - actual.data[offset + 2]!);
    if (
      redDelta > channelTolerance ||
      greenDelta > channelTolerance ||
      blueDelta > channelTolerance
    ) {
      changedPixels++;
    }
  }
  const totalPixels = expected.width * expected.height;
  return {
    page,
    width: expected.width,
    height: expected.height,
    changedPixels,
    changedRatio: totalPixels === 0 ? 1 : changedPixels / totalPixels,
  };
}

async function renderPdf(
  pdfPath: string,
  outputPrefix: string,
  dpi: number,
  executableOverride?: string,
): Promise<string[]> {
  const executable = executableOverride ?? process.env.PDFTOPPM ?? "pdftoppm";
  if (!existsSync(pdfPath)) throw new Error(`PDF not found: ${pdfPath}`);
  await run(executable, ["-png", "-r", String(dpi), pdfPath, outputPrefix], {
    windowsHide: true,
  });
  const directory = join(outputPrefix, "..");
  const prefix = basename(outputPrefix);
  const pages = readdirSync(directory)
    .filter((name) => name.startsWith(`${prefix}-`) && name.endsWith(".png"))
    .sort((left, right) => pageNumber(left, prefix) - pageNumber(right, prefix))
    .map((name) => join(directory, name));
  if (pages.length === 0) throw new Error(`pdftoppm produced no pages: ${pdfPath}`);
  return pages;
}

function pageNumber(filename: string, prefix: string): number {
  return Number(filename.slice(prefix.length + 1, -".png".length));
}

interface CommandLine {
  expectedPdf: string;
  actualPdf: string;
  options: PdfParityOptions;
  reportPath?: string;
}

function parseCommandLine(argv: readonly string[]): CommandLine | undefined {
  const positional: string[] = [];
  const options: PdfParityOptions = {};
  let reportPath: string | undefined;

  for (let index = 0; index < argv.length; index++) {
    const argument = argv[index];
    if (!argument) continue;
    if (argument === "--report") {
      reportPath = argv[index + 1];
      index++;
    } else if (argument === "--pdftoppm") {
      options.pdftoppm = argv[index + 1];
      index++;
    } else if (argument.startsWith("--dpi=")) {
      options.dpi = Number(argument.slice("--dpi=".length));
    } else if (argument.startsWith("--max-changed-ratio=")) {
      options.maxChangedRatio = Number(argument.slice("--max-changed-ratio=".length));
    } else if (argument.startsWith("--channel-tolerance=")) {
      options.channelTolerance = Number(argument.slice("--channel-tolerance=".length));
    } else {
      positional.push(argument);
    }
  }

  const [expectedPdf, actualPdf] = positional;
  if (!expectedPdf || !actualPdf) return undefined;
  return { expectedPdf, actualPdf, options, reportPath };
}

async function main(): Promise<number> {
  const command = parseCommandLine(process.argv.slice(2));
  if (!command) {
    process.stderr.write(
      "Usage: parity:pdf <word-oracle.pdf> <docen.pdf> [--report=parity.json] [--dpi=144] [--max-changed-ratio=0.002]\n",
    );
    return 2;
  }

  const report = await comparePdfFiles(command.expectedPdf, command.actualPdf, command.options);
  const serialized = `${JSON.stringify(report, null, 2)}\n`;
  if (command.reportPath) {
    writeFileSync(command.reportPath, serialized);
  } else {
    process.stdout.write(serialized);
  }

  const failingPages = report.pages.filter(
    (page) => page.reason === "dimension-mismatch" || page.changedRatio > report.maxChangedRatio,
  );
  if (failingPages.length > 0) {
    for (const page of failingPages) {
      process.stderr.write(
        `page ${page.page}: ${(page.changedRatio * 100).toFixed(3)}% changed${page.reason ? ` (${page.reason})` : ""}\n`,
      );
    }
    return 1;
  }
  process.stdout.write(
    `PDF parity passed: ${report.pageCount} pages, ${(report.changedRatio * 100).toFixed(3)}% changed\n`,
  );
  return 0;
}

if (process.argv[1]?.endsWith("parity-pdf.ts")) {
  main().then(
    (status) => (process.exitCode = status),
    (error: unknown) => {
      process.stderr.write(error instanceof Error ? `${error.message}\n` : `${String(error)}\n`);
      process.exitCode = 1;
    },
  );
}
