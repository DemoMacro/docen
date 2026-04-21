import { readFileSync, writeFileSync, existsSync, readdirSync, mkdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseDOCX, generateDOCX, generateHTML } from "docen";

// Get current file directory
const __dirname = dirname(fileURLToPath(import.meta.url));

const tmpDir = join(__dirname, ".cache");
const inputDir = join(tmpDir, "input");
const jsonDir = join(tmpDir, "json");
const htmlDir = join(tmpDir, "html");
const outputDir = join(tmpDir, "output");

// Check and create directories if they don't exist
for (const dir of [tmpDir, inputDir, jsonDir, htmlDir, outputDir]) {
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
          crop: true,
        },
        paragraph: {
          ignoreEmpty: false,
        },
      });

      // Save parsed JSON
      writeFileSync(jsonPath, JSON.stringify(json, null, 2), "utf-8");
      console.log(`  💾 Saved JSON: ${jsonFile}`);

      // Generate HTML from JSON
      const htmlFile = docxFile.replace(".docx", ".html");
      const htmlPath = join(htmlDir, htmlFile);
      const htmlContent = generateHTML(json);
      writeFileSync(htmlPath, htmlContent, "utf-8");
      console.log(`  💾 Saved HTML: ${htmlFile}`);

      // Generate DOCX from JSON
      const convertedDocxBuffer = await generateDOCX(json, {
        title: outputFile.replace(".docx", ""),
        outputType: "nodebuffer",
        sections: [
          {
            properties: {
              page: {
                margin: {
                  top: "2.54cm",
                  bottom: "2.54cm",
                  left: "3.18cm",
                  right: "3.18cm",
                  header: "1.5cm",
                  footer: "1.75cm",
                },
              },
              grid: {
                type: "lines",
              },
            },
            children: [],
          },
        ],
        styles: {
          default: {
            document: {
              paragraph: {
                alignment: "left",
                spacing: {
                  line: 480,
                },
                indent: {
                  firstLine: "1cm",
                },
              },
              run: {
                font: "宋体",
                size: 28,
              },
            },
            heading1: {
              paragraph: {
                alignment: "center",
                spacing: {
                  line: 480,
                },
                indent: {
                  firstLine: 0,
                },
                keepNext: true,
              },
              run: {
                bold: true,
                size: 36,
              },
            },
            heading2: {
              paragraph: {
                alignment: "center",
                spacing: {
                  line: 480,
                },
                indent: {
                  firstLine: 0,
                },
                keepNext: true,
              },
              run: {
                bold: true,
                size: 32,
              },
            },
            heading3: {
              paragraph: {
                alignment: "center",
                spacing: {
                  line: 480,
                },
                indent: {
                  firstLine: 0,
                },
                keepNext: true,
              },
              run: {
                bold: true,
                size: 30,
              },
            },
            heading4: {
              paragraph: {
                alignment: "center",
                spacing: {
                  line: 480,
                },
                indent: {
                  firstLine: 0,
                },
                keepNext: true,
              },
              run: {
                bold: true,
                size: 28,
              },
            },
            heading5: {
              paragraph: {
                alignment: "left",
                spacing: {
                  line: 480,
                },
                indent: {
                  firstLine: "1cm",
                },
              },
              run: {
                bold: true,
                size: 28,
              },
            },
            heading6: {
              paragraph: {
                alignment: "left",
                spacing: {
                  line: 480,
                },
                indent: {
                  firstLine: "1cm",
                },
              },
              run: {
                bold: true,
                size: 28,
              },
            },
          },
        },
        image: {
          style: {
            id: "Image",
            name: "Image",
            uiPriority: 1,
            semiHidden: false,
            unhideWhenUsed: true,
            quickFormat: true,
            paragraph: {
              alignment: "center",
              spacing: {
                line: 360,
              },
              indent: {
                firstLine: 0,
              },
            },
          },
        },
        table: {
          style: {
            id: "Table",
            name: "Table",
            uiPriority: 2,
            semiHidden: false,
            unhideWhenUsed: true,
            quickFormat: true,
            paragraph: {
              alignment: "center",
              keepNext: false,
              spacing: {
                line: 240, // Single line spacing (1.0)
              },
              indent: {
                firstLine: 0,
              },
            },
            run: {
              size: 28,
            },
          },
          run: {
            width: {
              size: 100,
              type: "pct", // Percentage width
            },
            alignment: "center", // Center align tables
            layout: "autofit", // Fixed layout for auto-width behavior
          },
          cell: {
            run: {
              verticalAlign: "center", // Vertical center alignment
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
  console.log(`📁 HTML directory: ${htmlDir}`);
  console.log(`📁 Output directory: ${outputDir}`);
})();
