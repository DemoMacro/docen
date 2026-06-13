import { readFileSync, writeFileSync, existsSync, readdirSync, rmSync, mkdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

import { generateDOCX, parseHTML } from "../src";

const __dirname = dirname(fileURLToPath(import.meta.url));

const htmlDir = join(__dirname, "html");
const jsonDir = join(__dirname, "json");
const docxDir = join(__dirname, "docx");

// Create/clean output directories
for (const dir of [jsonDir, docxDir]) {
  if (existsSync(dir)) {
    for (const file of readdirSync(dir)) {
      rmSync(join(dir, file), { force: true });
    }
  } else {
    mkdirSync(dir, { recursive: true });
  }
}

const htmlFiles = readdirSync(htmlDir).filter((f) => f.endsWith(".html"));

console.log(`⏳ Testing ${htmlFiles.length} files: HTML → JSON → DOCX\n`);

interface TestResult {
  file: string;
  success: boolean;
  error?: string;
}

const results: TestResult[] = [];

for (const htmlFile of htmlFiles) {
  try {
    const html = readFileSync(join(htmlDir, htmlFile), "utf-8");
    const baseName = htmlFile.replace(".html", "");

    // Step 1: HTML → Tiptap JSON
    const json = parseHTML(html);
    writeFileSync(join(jsonDir, `${baseName}.json`), JSON.stringify(json, null, 2));

    // Step 2: JSON → DOCX (generateDOCX prepares + compiles internally)
    const buffer = await generateDOCX(json);
    writeFileSync(join(docxDir, `${baseName}.docx`), buffer);

    results.push({ file: htmlFile, success: true });
    console.log(`  ✅ ${htmlFile}`);
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    results.push({ file: htmlFile, success: false, error: msg });
    console.log(`  ❌ ${htmlFile}: ${msg}`);
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
