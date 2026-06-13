import { readFileSync, writeFileSync, existsSync, readdirSync, rmSync, mkdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

import { parseDOCX, generateHTML, generateDOCX } from "../src";

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

console.log(`⏳ Testing ${docxFiles.length} files: DOCX → JSON → HTML → DOCX\n`);

interface TestResult {
  file: string;
  success: boolean;
  error?: string;
}

const results: TestResult[] = [];

for (const docxFile of docxFiles) {
  try {
    const baseName = docxFile.replace(".docx", "");
    const buffer = readFileSync(join(docxDir, docxFile));

    // Step 1: DOCX → Tiptap JSON (runtime model)
    const json = parseDOCX(buffer);
    writeFileSync(join(jsonDir, `${baseName}.json`), JSON.stringify(json, null, 2));

    // Step 2: Tiptap JSON → HTML
    const html = generateHTML(json);
    writeFileSync(join(htmlDir, `${baseName}.html`), html);

    // Step 3: Tiptap JSON → DOCX (regenerate to verify the full pipeline)
    const outBuffer = await generateDOCX(json);
    writeFileSync(join(docxOutDir, `${baseName}.docx`), outBuffer);

    results.push({ file: docxFile, success: true });
    console.log(`  ✅ ${docxFile}`);
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    results.push({ file: docxFile, success: false, error: msg });
    console.log(`  ❌ ${docxFile}: ${msg}`);
  }
}

// Summary
const passed = results.filter((r) => r.success).length;
const failed = results.length - passed;

console.log(`\n${"=".repeat(50)}`);
console.log(`📊 ${passed}/${results.length} passed`);

if (failed > 0) {
  console.log(`\n❌ Failed ${failed}:`);
  for (const r of results) {
    if (!r.success) console.log(`  ${r.file}: ${r.error}`);
  }
} else {
  console.log("\n🎉 All tests passed!");
}

console.log("=".repeat(50));
