import { readFileSync, writeFileSync, existsSync, readdirSync, mkdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseDOCX, generateDOCX } from "docen";
import { convertMillimetersToTwip } from "docx";

// Get current file directory
const __dirname = dirname(fileURLToPath(import.meta.url));

const tmpDir = join(__dirname, ".cache");
const inputDir = join(tmpDir, "input");
const jsonDir = join(tmpDir, "json");
const outputDir = join(tmpDir, "output");

// Check and create directories if they don't exist
for (const dir of [tmpDir, inputDir, jsonDir, outputDir]) {
  if (!existsSync(dir)) {
    mkdirSync(dir, { recursive: true });
  }
}

// Check if input directory exists
if (!existsSync(inputDir)) {
  console.log("⚠️  Input directory not found, creating it...");
  mkdirSync(inputDir, { recursive: true });
  console.log("📁 Please put DOCX files in the .cache/input directory and run again.");
  process.exit(0);
}

// Get all DOCX files from input directory
const docxFiles = readdirSync(inputDir).filter((file) => file.endsWith(".docx"));

if (docxFiles.length === 0) {
  console.log("⚠️  No DOCX files found in .cache/input directory.");
  console.log("📁 Please put DOCX files in the .cache/input directory and run again.");
  process.exit(0);
}

console.log(`⏳ Testing ${docxFiles.length} DOCX file(s)...`);

// Process each DOCX file
void (async () => {
  for (const docxFile of docxFiles) {
    try {
      const docxPath = join(inputDir, docxFile);
      const outputFile = docxFile.replace(".docx", "-converted.docx");
      const outputPath = join(outputDir, outputFile);
      const jsonFile = docxFile.replace(".docx", ".json");
      const jsonPath = join(jsonDir, jsonFile);

      console.log(`📄 Processing: ${docxFile}`);

      // Read DOCX file
      const docxBuffer = readFileSync(docxPath);

      // Parse DOCX to JSON
      const json = await parseDOCX(docxBuffer, {
        image: {
          canvasImport: () => import("@napi-rs/canvas"),
          enableImageCrop: true,
        },
        ignoreEmptyParagraphs: false,
      });

      // Save parsed JSON
      writeFileSync(jsonPath, JSON.stringify(json, null, 2), "utf-8");
      console.log(`  💾 Saved JSON: ${jsonFile}`);

      // Generate DOCX from JSON
      const convertedDocxBuffer = await generateDOCX(json, {
        title: outputFile.replace(".docx", ""),
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
      });

      // Write converted DOCX
      writeFileSync(outputPath, convertedDocxBuffer);

      console.log(`✅ Successfully converted: ${docxFile} → ${outputFile}`);
    } catch (error) {
      console.error(`❌ Error processing ${docxFile}:`, error);
    }
  }

  console.log("\n🎉 All files processed!");
  console.log(`📁 Input directory: ${inputDir}`);
  console.log(`📁 JSON directory: ${jsonDir}`);
  console.log(`📁 Output directory: ${outputDir}`);
})();
