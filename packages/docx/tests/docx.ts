import { readFileSync, writeFileSync, existsSync, readdirSync, rmSync, mkdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { performance } from "node:perf_hooks";
import { fileURLToPath } from "node:url";

import { generateDocument } from "@office-open/docx";

import { parseDOCX, generateHTML, prepareDocument, compileDocument } from "../src";

const __dirname = dirname(fileURLToPath(import.meta.url));

const docxDir = join(__dirname, "docx"); // input: DOCX files produced by html.ts (forward test)
const tempDir = join(__dirname, ".temp");
const jsonDir = join(tempDir, "json");
const htmlDir = join(tempDir, "html");
const docxOutDir = join(tempDir, "docx");

// Create/clean output directories
for (const dir of [jsonDir, htmlDir, docxOutDir]) {
  if (existsSync(dir)) {
    for (const file of readdirSync(dir)) {
      rmSync(join(dir, file), { force: true });
    }
  } else {
    mkdirSync(dir, { recursive: true });
  }
}

const docxFiles = readdirSync(docxDir).filter((f) => f.endsWith(".docx"));

console.log(
  `⏳ Testing ${docxFiles.length} files: DOCX → JSON → HTML → (prepare → compile → generate)\n`,
);

interface StepTimings {
  parse: number;
  html: number;
  prepare: number;
  compile: number;
  generate: number;
}

interface TestResult {
  file: string;
  success: boolean;
  error?: string;
  timings?: StepTimings;
}

const results: TestResult[] = [];

for (const docxFile of docxFiles) {
  try {
    const baseName = docxFile.replace(".docx", "");
    const buffer = readFileSync(join(docxDir, docxFile));

    // Step 1: DOCX → Tiptap JSON (runtime model)
    let t0 = performance.now();
    const json = parseDOCX(buffer);
    const tParse = performance.now() - t0;
    writeFileSync(join(jsonDir, `${baseName}.json`), JSON.stringify(json, null, 2));

    // Step 2: Tiptap JSON → HTML
    t0 = performance.now();
    const html = generateHTML(json);
    const tHtml = performance.now() - t0;
    writeFileSync(join(htmlDir, `${baseName}.html`), html);

    // Step 3: Tiptap JSON → DOCX — split into prepare/compile/generate to
    // pinpoint the generateDOCX bottleneck. No `document` option is passed, so
    // applyDocumentOptions is a no-op (returns base) and is omitted.
    t0 = performance.now();
    await prepareDocument(json);
    const tPrepare = performance.now() - t0;

    t0 = performance.now();
    const docOpts = compileDocument(json);
    const tCompile = performance.now() - t0;

    t0 = performance.now();
    const outBuffer = await generateDocument(docOpts);
    const tGenerate = performance.now() - t0;
    writeFileSync(join(docxOutDir, `${baseName}.docx`), outBuffer);

    results.push({
      file: docxFile,
      success: true,
      timings: {
        parse: tParse,
        html: tHtml,
        prepare: tPrepare,
        compile: tCompile,
        generate: tGenerate,
      },
    });
    console.log(
      `  ✅ ${docxFile}  parse ${tParse.toFixed(1)} · html ${tHtml.toFixed(1)} · prep ${tPrepare.toFixed(1)} · comp ${tCompile.toFixed(1)} · gen ${tGenerate.toFixed(1)} (ms)`,
    );
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    results.push({ file: docxFile, success: false, error: msg });
    console.log(`  ❌ ${docxFile}: ${msg}`);
  }
}

// Summary
const passed = results.filter((r) => r.success).length;
const failed = results.length - passed;

console.log(`\n${"=".repeat(78)}`);
console.log(`📊 ${passed}/${results.length} passed`);

if (failed > 0) {
  console.log(`\n❌ Failed ${failed}:`);
  for (const r of results) {
    if (!r.success) console.log(`  ${r.file}: ${r.error}`);
  }
} else {
  console.log("🎉 All tests passed!");
}

// Per-file timing breakdown, slowest first.
const timed = results
  .filter((r): r is TestResult & { timings: StepTimings } => !!r.timings)
  .map((r) => ({
    file: r.file,
    ...r.timings,
    total:
      r.timings.parse + r.timings.html + r.timings.prepare + r.timings.compile + r.timings.generate,
  }))
  .sort((a, b) => b.total - a.total);

if (timed.length) {
  console.log(`\n⏱  Timings (ms):`);
  console.log(
    `  ${"file".padEnd(34)} ${"parse".padStart(7)} ${"html".padStart(7)} ${"prep".padStart(7)} ${"comp".padStart(7)} ${"gen".padStart(8)} ${"total".padStart(8)}`,
  );
  for (const t of timed) {
    console.log(
      `  ${t.file.slice(0, 34).padEnd(34)} ${t.parse.toFixed(1).padStart(7)} ${t.html.toFixed(1).padStart(7)} ${t.prepare.toFixed(1).padStart(7)} ${t.compile.toFixed(1).padStart(7)} ${t.generate.toFixed(1).padStart(8)} ${t.total.toFixed(1).padStart(8)}`,
    );
  }
  const sum = timed.reduce(
    (acc, t) => ({
      parse: acc.parse + t.parse,
      html: acc.html + t.html,
      prepare: acc.prepare + t.prepare,
      compile: acc.compile + t.compile,
      generate: acc.generate + t.generate,
      total: acc.total + t.total,
    }),
    { parse: 0, html: 0, prepare: 0, compile: 0, generate: 0, total: 0 },
  );
  console.log(
    `  ${"".padEnd(34)} ${"-".repeat(7)} ${"-".repeat(7)} ${"-".repeat(7)} ${"-".repeat(7)} ${"-".repeat(8)} ${"-".repeat(8)}`,
  );
  console.log(
    `  ${`total (${timed.length})`.padEnd(34)} ${sum.parse.toFixed(1).padStart(7)} ${sum.html.toFixed(1).padStart(7)} ${sum.prepare.toFixed(1).padStart(7)} ${sum.compile.toFixed(1).padStart(7)} ${sum.generate.toFixed(1).padStart(8)} ${sum.total.toFixed(1).padStart(8)}`,
  );
}

console.log("=".repeat(78));
