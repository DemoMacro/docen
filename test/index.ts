import { readFileSync, writeFileSync, existsSync, readdirSync, rmSync, mkdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { generateDOCX } from "../packages/export-docx/src";
import { parseDOCX } from "../packages/import-docx/src";
import { generateJSON, generateHTML } from "./html";
import { PageBreak } from "../packages/export-docx/src/docx";
import { unzipSync } from "fflate";
import { fromXml } from "xast-util-from-xml";
import { convertMillimetersToTwip } from "docx";

// Get current file directory
const __dirname = dirname(fileURLToPath(import.meta.url));

// Get all HTML files from html directory
const htmlDir = join(__dirname, "html");
const jsonDir = join(__dirname, "json");
const docxDir = join(__dirname, "docx");
const parsedJsonDir = join(__dirname, "json-parsed");
const parsedHtmlDir = join(__dirname, "html-parsed");
const parsedDocxDir = join(__dirname, "docx-parsed");
const docxStructureDir = join(__dirname, "docx-structure");

// Check and create directories if they don't exist
if (!existsSync(jsonDir)) {
  console.log("Creating json directory...");
  mkdirSync(jsonDir, { recursive: true });
} else {
  // Clear json directory
  console.log("Clearing json directory...");
  const jsonFiles = readdirSync(jsonDir);
  for (const file of jsonFiles) {
    rmSync(join(jsonDir, file), { force: true, recursive: true });
  }
}

if (!existsSync(docxDir)) {
  console.log("Creating docx directory...");
  mkdirSync(docxDir, { recursive: true });
} else {
  // Clear docx directory
  console.log("Clearing docx directory...");
  const docxFiles = readdirSync(docxDir);
  for (const file of docxFiles) {
    rmSync(join(docxDir, file), { force: true, recursive: true });
  }
}

if (!existsSync(parsedJsonDir)) {
  console.log("Creating json-parsed directory...");
  mkdirSync(parsedJsonDir, { recursive: true });
} else {
  // Clear json-parsed directory
  console.log("Clearing json-parsed directory...");
  const parsedJsonFiles = readdirSync(parsedJsonDir);
  for (const file of parsedJsonFiles) {
    rmSync(join(parsedJsonDir, file), { force: true, recursive: true });
  }
}

if (!existsSync(docxStructureDir)) {
  console.log("Creating docx-structure directory...");
  mkdirSync(docxStructureDir, { recursive: true });
} else {
  // Clear docx-structure directory
  console.log("Clearing docx-structure directory...");
  const structureFiles = readdirSync(docxStructureDir);
  for (const file of structureFiles) {
    rmSync(join(docxStructureDir, file), { force: true, recursive: true });
  }
}

if (!existsSync(parsedHtmlDir)) {
  console.log("Creating html-parsed directory...");
  mkdirSync(parsedHtmlDir, { recursive: true });
} else {
  // Clear html-parsed directory
  console.log("Clearing html-parsed directory...");
  const parsedHtmlFiles = readdirSync(parsedHtmlDir);
  for (const file of parsedHtmlFiles) {
    rmSync(join(parsedHtmlDir, file), { force: true, recursive: true });
  }
}

if (!existsSync(parsedDocxDir)) {
  console.log("Creating docx-parsed directory...");
  mkdirSync(parsedDocxDir, { recursive: true });
} else {
  // Clear docx-parsed directory
  console.log("Clearing docx-parsed directory...");
  const parsedDocxFiles = readdirSync(parsedDocxDir);
  for (const file of parsedDocxFiles) {
    rmSync(join(parsedDocxDir, file), { force: true, recursive: true });
  }
}

// Read all HTML files
const htmlFiles = readdirSync(htmlDir).filter((file) => file.endsWith(".html"));

console.log(`Found ${htmlFiles.length} HTML files to convert`);

// Convert each HTML file to JSON
htmlFiles.forEach((htmlFile) => {
  try {
    console.log(`Converting ${htmlFile}...`);

    const htmlPath = join(htmlDir, htmlFile);
    const jsonFile = htmlFile.replace(".html", ".json");
    const jsonPath = join(jsonDir, jsonFile);

    const html = readFileSync(htmlPath, "utf-8");
    const json = generateJSON(html);

    writeFileSync(jsonPath, JSON.stringify(json, null, 2));

    console.log(`✓ Converted ${htmlFile} to ${jsonFile}`);
  } catch (error) {
    console.error(`✗ Error converting ${htmlFile}:`, error);
  }
});

console.log("\nConverting JSON to DOCX...");

// Read all JSON files and convert to DOCX
const jsonFiles = readdirSync(jsonDir).filter((file) => file.endsWith(".json"));

// Simple JSON comparison function
function compareJSON(original: any, parsed: any, path = ""): string[] {
  const differences: string[] = [];

  if (original === parsed) return differences;

  if (typeof original !== typeof parsed) {
    differences.push(`${path}: Type mismatch (${typeof original} vs ${typeof parsed})`);
    return differences;
  }

  if (typeof original !== "object" || original === null || parsed === null) {
    if (original !== parsed) {
      differences.push(
        `${path}: Value mismatch (${JSON.stringify(original)} vs ${JSON.stringify(parsed)})`,
      );
    }
    return differences;
  }

  // Handle arrays
  if (Array.isArray(original) && Array.isArray(parsed)) {
    if (original.length !== parsed.length) {
      differences.push(`${path}: Array length mismatch (${original.length} vs ${parsed.length})`);
      return differences;
    }

    // For marks array, compare as sets (ignore order)
    if (path.includes("marks")) {
      const matched = Array.from({ length: parsed.length }, () => false);
      for (let i = 0; i < original.length; i++) {
        let found = false;
        for (let j = 0; j < parsed.length; j++) {
          if (!matched[j] && compareJSON(original[i], parsed[j], "").length === 0) {
            matched[j] = true;
            found = true;
            break;
          }
        }
        if (!found) {
          differences.push(`${path}[${i}]: Element not found in parsed array`);
        }
      }
      return differences;
    }

    // For other arrays, compare in order
    for (let i = 0; i < original.length; i++) {
      differences.push(...compareJSON(original[i], parsed[i], `${path}[${i}]`));
    }
    return differences;
  }

  if (Array.isArray(original) || Array.isArray(parsed)) {
    differences.push(`${path}: Type mismatch (array vs non-array)`);
    return differences;
  }

  // Special handling for link marks: only compare href, ignore unsupported attributes (target, rel, class)
  if (original.type === "link" && parsed.type === "link" && original.attrs && parsed.attrs) {
    // Only compare href attribute for links
    if (original.attrs.href !== parsed.attrs.href) {
      differences.push(
        `${path}.attrs.href: Value mismatch (${JSON.stringify(original.attrs.href)} vs ${JSON.stringify(parsed.attrs.href)})`,
      );
    }
    return differences;
  }

  // Special handling for table cell types: tableHeader and tableCell are equivalent (DOCX format limitation)
  if (
    original.type &&
    parsed.type &&
    ((original.type === "tableHeader" && parsed.type === "tableCell") ||
      (original.type === "tableCell" && parsed.type === "tableHeader"))
  ) {
    // Skip the type field comparison for tableHeader/tableCell, but compare other fields
    const originalCopy = { ...original };
    const parsedCopy = { ...parsed };
    delete originalCopy.type;
    delete parsedCopy.type;

    const originalKeys = Object.keys(originalCopy);
    const parsedKeys = Object.keys(parsedCopy);

    // Check that all required keys from original are present in parsed
    for (const key of originalKeys) {
      if (!parsedKeys.includes(key)) {
        differences.push(`${path}.${key}: Missing in parsed`);
      }
    }

    // Compare common keys (excluding type)
    for (const key of originalKeys) {
      if (parsedKeys.includes(key)) {
        differences.push(
          ...compareJSON(originalCopy[key], parsedCopy[key], path ? `${path}.${key}` : key),
        );
      }
    }

    // Check for extra keys (excluding type which we already handled)
    for (const key of parsedKeys) {
      if (!originalKeys.includes(key) && key !== "attrs") {
        differences.push(`${path}.${key}: Extra in parsed`);
      }
    }

    return differences;
  }

  const originalKeys = Object.keys(original);
  const parsedKeys = Object.keys(parsed);

  // Special handling for attrs: only check required keys exist, ignore extra keys
  const isAttrsObject = path.endsWith(".attrs");
  if (isAttrsObject) {
    // Check that all required keys from original are present in parsed
    for (const key of originalKeys) {
      if (!parsedKeys.includes(key)) {
        differences.push(`${path}.${key}: Missing in parsed`);
      }
    }
    // Extra keys in attrs are acceptable (DOCX format limitations)
    // Compare only common keys
    for (const key of originalKeys) {
      if (parsedKeys.includes(key)) {
        differences.push(...compareJSON(original[key], parsed[key], path ? `${path}.${key}` : key));
      }
    }
    return differences;
  }

  // For non-attrs objects, check both missing and extra keys
  // Check for missing keys
  for (const key of originalKeys) {
    if (!parsedKeys.includes(key)) {
      differences.push(`${path}.${key}: Missing in parsed`);
    }
  }

  // Check for extra keys
  for (const key of parsedKeys) {
    if (!originalKeys.includes(key)) {
      // Extra 'attrs' keys are acceptable (DOCX format limitations)
      // DOCX may add additional formatting attributes
      if (key !== "attrs") {
        differences.push(`${path}.${key}: Extra in parsed`);
      }
    }
  }

  // Compare common keys
  for (const key of originalKeys) {
    if (parsedKeys.includes(key)) {
      differences.push(...compareJSON(original[key], parsed[key], path ? `${path}.${key}` : key));
    }
  }

  return differences;
}

// Files to skip comparison (not implemented or not required)
const skipComparisonFiles = new Set([
  "blockquote.json", // DOCX doesn't have semantic blockquote structure
  "code.json", // code mark cannot be properly imported (DOCX limitation)
  "code-block-lowlight.json",
  "code-block-with-language.json",
  "details.json",
  "image.json", // DOCX format limitations
  "mathematics.json",
  "table-cell.json", // contains code marks that cannot be properly imported
  "text-style.json", // color names converted to hex (red → #FF0000) is expected behavior
  "text.json", // contains code marks that cannot be properly imported
]);

// Process files sequentially
void (async () => {
  for (const jsonFile of jsonFiles) {
    try {
      console.log(`\n--- Processing ${jsonFile} ---`);

      const jsonPath = join(jsonDir, jsonFile);
      const docxFile = jsonFile.replace(".json", ".docx");
      const docxPath = join(docxDir, docxFile);
      const parsedJsonFile = jsonFile;
      const parsedJsonPath = join(parsedJsonDir, parsedJsonFile);

      const originalJSON = JSON.parse(readFileSync(jsonPath, "utf-8"));

      // Generate DOCX
      console.log(`  → Generating ${docxFile}...`);
      const docxBuffer = await generateDOCX(originalJSON, {
        title: docxFile.replace(".docx", ""),
        outputType: "nodebuffer",
        sections: [
          {
            properties: {
              page: {
                size: {
                  width: convertMillimetersToTwip(210),
                  height: convertMillimetersToTwip(297),
                },
                margin: {
                  top: convertMillimetersToTwip(20),
                  right: convertMillimetersToTwip(20),
                  bottom: convertMillimetersToTwip(20),
                  left: convertMillimetersToTwip(20),
                },
              },
            },
            children: [],
          },
        ],
        table: {
          run: {
            width: {
              size: 100,
              type: "pct",
            },
            alignment: "center",
            layout: "autofit",
          },
          cell: {
            paragraph: {
              alignment: "center",
            },
            run: {
              verticalAlign: "center",
            },
          },
        },
        styles: {
          default: {
            document: {
              paragraph: {
                spacing: {
                  line: 480,
                },
              },
              run: {
                size: 28,
              },
            },
          },
        },
        horizontalRule: {
          paragraph: {
            border: undefined,
            children: [new PageBreak()],
          },
        },
      });

      writeFileSync(docxPath, docxBuffer);
      console.log(`  ✓ Generated ${docxFile}`);

      // Save DOCX XML structure for debugging
      console.log(`  → Saving DOCX XML structure...`);
      const files = unzipSync(docxBuffer);
      const documentXml = new TextDecoder().decode(files["word/document.xml"]);
      const documentXast = fromXml(documentXml);
      const xastJsonPath = join(docxStructureDir, jsonFile);
      writeFileSync(xastJsonPath, JSON.stringify(documentXast, null, 2));
      console.log(`  ✓ Saved DOCX structure`);

      // Parse DOCX back to JSON
      console.log(`  → Parsing ${docxFile} back to JSON...`);
      const parsedJSON = await parseDOCX(docxBuffer);

      writeFileSync(parsedJsonPath, JSON.stringify(parsedJSON, null, 2));

      // Convert parsed JSON back to HTML
      console.log(`  → Converting parsed JSON back to HTML...`);
      const parsedHtmlFile = jsonFile.replace(".json", ".html");
      const parsedHtmlPath = join(parsedHtmlDir, parsedHtmlFile);
      const parsedHTML = generateHTML(parsedJSON);
      writeFileSync(parsedHtmlPath, parsedHTML);
      console.log(`  ✓ Generated ${parsedHtmlFile}`);

      // Convert parsed JSON back to DOCX
      console.log(`  → Converting parsed JSON back to DOCX...`);
      const parsedDocxFile = jsonFile.replace(".json", ".docx");
      const parsedDocxPath = join(parsedDocxDir, parsedDocxFile);
      const parsedDocxBuffer = await generateDOCX(parsedJSON, {
        title: parsedDocxFile.replace(".docx", ""),
        outputType: "nodebuffer",
        table: {
          run: {
            width: {
              size: 100,
              type: "pct",
            },
            alignment: "center",
            layout: "autofit",
          },
          cell: {
            paragraph: {
              alignment: "center",
            },
            run: {
              verticalAlign: "center",
            },
          },
        },
        styles: {
          default: {
            document: {
              paragraph: {
                spacing: {
                  line: 480,
                },
              },
              run: {
                size: 28,
              },
            },
          },
        },
        horizontalRule: {
          paragraph: {
            border: undefined,
            children: [new PageBreak()],
          },
        },
      });
      writeFileSync(parsedDocxPath, parsedDocxBuffer);
      console.log(`  ✓ Generated ${parsedDocxFile}`);

      // Compare JSONs (skip for files that are not implemented)
      if (skipComparisonFiles.has(jsonFile)) {
        console.log(`  → Skipping comparison (not implemented)`);
      } else {
        console.log(`  → Comparing JSONs...`);
        const differences = compareJSON(originalJSON, parsedJSON);

        if (differences.length === 0) {
          console.log(`  ✓ ✓ ✓ Perfect match! JSONs are identical.`);
        } else {
          console.log(`  ✗ ✗ ✗ Found ${differences.length} differences:`);
          differences.slice(0, 10).forEach((diff) => console.log(`    - ${diff}`));
          if (differences.length > 10) {
            console.log(`    ... and ${differences.length - 10} more differences`);
          }
        }
      }
    } catch (error) {
      console.error(`  ✗ Error processing ${jsonFile}:`, error);
    }
  }

  console.log("\n=== Conversion complete! ===");
})();
