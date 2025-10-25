import {
  readFileSync,
  writeFileSync,
  existsSync,
  readdirSync,
  rmSync,
  mkdirSync,
} from "node:fs";
import { join } from "node:path";
import { generateJSON, generateDOCX } from "../packages/export-docx/src";

// Get all HTML files from html directory
const htmlDir = "html";
const jsonDir = "json";
const docxDir = "docx";

// Check and create directories if they don't exist
if (!existsSync(jsonDir)) {
  console.log("Creating json directory...");
  mkdirSync(jsonDir, { recursive: true });
} else {
  // Clear json directory
  console.log("Clearing json directory...");
  const jsonFiles = readdirSync(jsonDir);
  for (const file of jsonFiles) {
    rmSync(join(jsonDir, file), { force: true });
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
    rmSync(join(docxDir, file), { force: true });
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

jsonFiles.forEach(async (jsonFile) => {
  try {
    console.log(`Converting ${jsonFile} to DOCX...`);

    const jsonPath = join(jsonDir, jsonFile);
    const docxFile = jsonFile.replace(".json", ".docx");
    const docxPath = join(docxDir, docxFile);

    const jsonContent = JSON.parse(readFileSync(jsonPath, "utf-8"));
    const docxBuffer = await generateDOCX(jsonContent, {
      title: docxFile.replace(".docx", ""),
      outputType: "nodebuffer",
      table: {
        properties: {
          width: {
            size: 100,
            type: "pct", // Percentage width
          },
          alignment: "center", // Center align tables
          layout: "autofit", // Fixed layout for auto-width behavior
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
    });

    writeFileSync(docxPath, docxBuffer);

    console.log(`✓ Converted ${jsonFile} to ${docxFile}`);
  } catch (error) {
    console.error(`✗ Error converting ${jsonFile} to DOCX:`, error);
  }
});

console.log("\nConversion complete!");
