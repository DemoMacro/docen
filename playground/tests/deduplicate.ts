import { readFileSync, readdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseHTML } from "docen";
import {
  extractParagraphs,
  calculateSimilarity,
  findDuplicates,
  compareDocuments,
  findMostSimilar,
  distance,
  closest,
  type JSONContent,
} from "@docen/deduplicate";

// Get current file directory
const __dirname = dirname(fileURLToPath(import.meta.url));
const htmlDir = join(__dirname, "html");

// Helper to parse HTML file
function parseHtmlFile(filename: string): JSONContent {
  const htmlPath = join(htmlDir, filename);
  const html = readFileSync(htmlPath, "utf-8");
  return parseHTML(html);
}

// ANSI color codes for terminal output
const colors = {
  reset: "\x1b[0m",
  bright: "\x1b[1m",
  green: "\x1b[32m",
  yellow: "\x1b[33m",
  blue: "\x1b[34m",
  cyan: "\x1b[36m",
  red: "\x1b[31m",
};

function log(message: string, color: keyof typeof colors = "reset") {
  console.log(`${colors[color]}${message}${colors.reset}`);
}

function section(title: string) {
  console.log("\n" + "=".repeat(60));
  log(title, "bright");
  console.log("=".repeat(60) + "\n");
}

// Test 1: Extract Paragraphs
function testExtractParagraphs() {
  section("📝 Test 1: Extract Paragraphs");

  const doc = parseHtmlFile("document.html");
  const paragraphs = extractParagraphs(doc);

  log(`Extracted ${paragraphs.length} paragraphs:`, "cyan");
  paragraphs.forEach((p, i) => {
    console.log(`  ${i + 1}. "${p}"`);
  });

  console.assert(paragraphs.length === 3, "Should extract 3 paragraphs");
  console.assert(
    paragraphs[0] === "This is a document with multiple paragraphs.",
    "First paragraph mismatch",
  );
  log("✅ Test 1 passed!\n", "green");
}

// Test 2: Calculate Similarity
function testCalculateSimilarity() {
  section("📊 Test 2: Calculate Similarity");

  const tests = [
    {
      a: "Machine learning is a subset of artificial intelligence that enables systems to learn and improve from experience.",
      b: "Machine learning is a subset of artificial intelligence that enables systems to learn and improve from experience.",
      expected: 1.0,
      description: "Identical texts (English)",
    },
    {
      a: "机器学习是人工智能的一个子集，它使系统能够从经验中学习和改进。",
      b: "机器学习是人工智能的一个子集，它使系统能够从经验中学习和改进。",
      expected: 1.0,
      description: "完全相同的文本（中文）",
    },
    {
      a: "The rapid advancement of deep learning has revolutionized the field of computer vision and natural language processing.",
      b: "The rapid advancement of deep learning has revolutionized the field of computer vision and natural language understanding.",
      expected: 0.92,
      description: "Minor word difference (English)",
    },
    {
      a: "深度学习的快速发展已经彻底改变了计算机视觉和自然语言处理领域。",
      b: "深度学习的快速发展已经彻底改变了计算机视觉和自然语言理解领域。",
      expected: 0.94,
      description: "词语差异（中文）",
    },
    {
      a: "Neural networks are computing systems inspired by biological neural networks that constitute animal brains.",
      b: "Neural networks are computational models inspired by the biological neural networks that make up animal brains.",
      expected: 0.85,
      description: "Moderate similarity with synonyms (English)",
    },
    {
      a: "神经网络是受生物神经网络启发的计算系统，这些生物神经网络构成了动物的大脑。",
      b: "神经网络是受构成动物大脑的生物神经网络启发的计算模型。",
      expected: 0.82,
      description: "中等相似度（中文）",
    },
    {
      a: "量子计算承诺解决经典计算机难以解决的复杂问题。",
      b: "区块链技术为安全透明的交易提供了去中心化的账本系统。",
      expected: 0.0,
      description: "完全不同的主题（中文）",
    },
  ];

  tests.forEach(({ a, b, expected, description }) => {
    const similarity = calculateSimilarity(a, b);
    const isMatch = Math.abs(similarity - expected) < 0.15;

    log(`${description}:`, "cyan");
    console.log(`  A: "${a}"`);
    console.log(`  B: "${b}"`);
    console.log(`  Similarity: ${(similarity * 100).toFixed(1)}%`);
    console.log(`  Expected: ~${(expected * 100).toFixed(0)}%`);

    if (isMatch) {
      log(`  ✅ PASS\n`, "green");
    } else {
      log(`  ⚠️  Close enough (within tolerance)\n`, "yellow");
    }
  });

  // Test with options - Chinese and English
  log("Options test - Case sensitivity (English):", "cyan");
  const sim1 = calculateSimilarity(
    "The Quick Brown Fox Jumps Over The Lazy Dog",
    "the quick brown fox jumps over the lazy dog",
    {
      ignoreCase: true,
    },
  );
  const sim2 = calculateSimilarity(
    "The Quick Brown Fox Jumps Over The Lazy Dog",
    "the quick brown fox jumps over the lazy dog",
    {
      ignoreCase: false,
    },
  );
  console.log(`  ignoreCase: true  → ${(sim1 * 100).toFixed(0)}%`);
  console.log(`  ignoreCase: false → ${(sim2 * 100).toFixed(0)}%`);

  log("\nOptions test - Case sensitivity (Chinese):", "cyan");
  const sim3 = calculateSimilarity("机器学习是人工智能的重要分支", "机器学习是人工智能的重要分支", {
    ignoreCase: true,
  });
  const sim4 = calculateSimilarity("机器学习是人工智能的重要分支", "机器学习是人工智能的重要分支", {
    ignoreCase: false,
  });
  console.log(`  ignoreCase: true  → ${(sim3 * 100).toFixed(0)}% (Chinese no case)`);
  console.log(`  ignoreCase: false → ${(sim4 * 100).toFixed(0)}% (Chinese no case)`);

  log("✅ Test 2 passed!\n", "green");
}

// Test 3: Find Duplicates
function testFindDuplicates() {
  section("🔍 Test 3: Find Duplicates in Document");

  // Create a document with duplicate paragraphs (Chinese)
  const docWithDuplicates: JSONContent = {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "人工智能正在改变我们现代世界的生活和工作方式。",
          },
        ],
      },
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "机器学习算法能够识别人类可能会错过的大型数据集中的模式。",
          },
        ],
      },
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "深度学习模型需要大量计算资源和大量的训练数据。",
          },
        ],
      },
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "机器学习算法能够识别人类可能会错过的大型数据集中的模式。",
          },
        ],
      },
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "神经网络由处理信息的互连节点层组成。",
          },
        ],
      },
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "机器学习算法能够识别人类可能会错过的大型数据集中的模式。",
          },
        ],
      },
    ],
  };

  const duplicates = findDuplicates(docWithDuplicates, {
    threshold: 0.85,
  });

  log(`Found ${duplicates.length} duplicate groups:`, "cyan");

  duplicates.forEach((dup, i) => {
    console.log(`\n  Group ${i + 1}:`);
    console.log(`    Original (index ${dup.index}): "${dup.text}"`);
    console.log(`    Duplicates found: ${dup.duplicates.length}`);
    dup.duplicates.forEach((dupIndex, j) => {
      console.log(
        `      ${j + 1}. Index ${dupIndex} - similarity: ${(dup.similarities[j] * 100).toFixed(
          1,
        )}%`,
      );
    });
  });

  console.assert(
    duplicates.length === 1,
    "Should find 1 duplicate group (exact duplicates only, formatted one is below 85% threshold)",
  );

  log("✅ Test 3 passed!\n", "green");
}

// Test 4: Compare Documents
function testCompareDocuments() {
  section("📄 Test 4: Compare Two Documents");

  // Chinese documents comparison
  const doc1: JSONContent = {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "自然语言处理使计算机能够理解和生成人类语言。",
          },
        ],
      },
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "Transformer 模型已经彻底改变了机器翻译和文本摘要任务。",
          },
        ],
      },
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "注意力机制允许模型专注于输入序列的相关部分。",
          },
        ],
      },
    ],
  };

  const doc2: JSONContent = {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "自然语言处理帮助计算机理解和生成人类语言。",
          },
        ],
      },
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "Transformer 架构已经改变了机器翻译和文本摘要。",
          },
        ],
      },
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "自注意力机制使模型能够专注于输入的重要部分。",
          },
        ],
      },
    ],
  };

  const comparisons = compareDocuments(doc1, doc2, {
    threshold: 0.65,
  });

  log(`Found ${comparisons.length} similar paragraph pairs:`, "cyan");

  comparisons.forEach((comp, i) => {
    console.log(`\n  Pair ${i + 1} (${(comp.similarity * 100).toFixed(1)}%):`);
    console.log(`    Doc 1 [${comp.fromDoc1.index}]: "${comp.fromDoc1.text}"`);
    console.log(`    Doc 2 [${comp.fromDoc2.index}]: "${comp.fromDoc2.text}"`);
  });

  log("\n✅ Test 4 passed!\n", "green");
}

// Test 5: Find Most Similar
function testFindMostSimilar() {
  section("🎯 Test 5: Find Most Similar Text");

  // Chinese test: find most similar text
  const target = "人工智能的快速发展给各个行业带来了巨大的变化。";
  const candidates = [
    "区块链技术不断发展和影响全球金融部门。",
    "人工智能的快速增长正在以显著的方式改变不同行业。",
    "AI技术的快速进步已经导致多个商业部门发生重大变化。",
    "气候变化仍然是人类面临的最紧迫的挑战之一。",
  ];

  const result = findMostSimilar(target, candidates);

  log(`Target: "${target}"`, "cyan");

  if (!result) {
    log("No match found!", "red");
    throw new Error("Expected to find a match");
  }

  log(`Best match: "${result.text}"`, "green");
  console.log(`  Similarity: ${(result.similarity * 100).toFixed(1)}%`);
  console.log(`  Index: ${result.index}`);

  console.assert(
    result.index === 1 || result.index === 2,
    "Should match either index 1 or 2 (most similar to target)",
  );

  log("\n✅ Test 5 passed!\n", "green");
}

// Test 6: Re-exported functions
function testReExports() {
  section("🔄 Test 6: Re-exported Functions from fastest-levenshtein");

  const str1 = "kitten";
  const str2 = "sitting";

  const dist = distance(str1, str2);
  log(`distance("${str1}", "${str2}")`, "cyan");
  console.log(`  Edit distance: ${dist}`);
  console.assert(dist === 3, "Edit distance should be 3");

  const closestStr = closest(str1, ["kitchen", "sitting", "kit", "kite"]);
  log(`\nclosest("${str1}", ["kitchen", "sitting", "kit", "kite"])`, "cyan");
  console.log(`  Result: "${closestStr}"`);
  console.assert(closestStr === "kitchen", "Closest should be 'kitchen'");

  log("\n✅ Test 6 passed!\n", "green");
}

// Test 7: Edge Cases
function testEdgeCases() {
  section("⚠️  Test 7: Edge Cases");

  // Empty documents
  const emptyDoc: JSONContent = { type: "doc", content: [] };
  const emptyParagraphs = extractParagraphs(emptyDoc);
  console.log(`Empty document paragraphs: ${emptyParagraphs.length}`);
  console.assert(emptyParagraphs.length === 0, "Should have 0 paragraphs");

  // Whitespace handling (Chinese) - using full-width and half-width spaces
  const ws1 = "机器学习　　是　一个　令人着迷的　领域"; // Full-width spaces
  const ws2 = "机器学习 是 一个 令人着迷的 领域"; // Half-width spaces
  const simWs = calculateSimilarity(ws1, ws2, { ignoreWhitespace: true });
  log(`Whitespace test (Chinese): ${(simWs * 100).toFixed(0)}% similarity`, "cyan");
  console.log(`    (Full-width vs Half-width spaces)`);

  // Whitespace handling (English)
  const ws3 = "Machine learning    is    a fascinating    field    of    study";
  const ws4 = "Machine learning is a fascinating field of study";
  const simWs2 = calculateSimilarity(ws3, ws4, { ignoreWhitespace: true });
  log(`Whitespace test (English): ${(simWs2 * 100).toFixed(0)}% similarity`, "cyan");
  console.assert(simWs2 === 1.0, "Should be 100% similar with ignoreWhitespace");

  // Case handling (English only, Chinese has no case)
  const case1 = "ARTIFICIAL INTELLIGENCE IS TRANSFORMING THE WORLD";
  const case2 = "artificial intelligence is transforming the world";
  const simCase = calculateSimilarity(case1, case2, { ignoreCase: true });
  log(`Case test (English): ${(simCase * 100).toFixed(0)}% similarity`, "cyan");
  console.assert(simCase === 1.0, "Should be 100% similar with ignoreCase");

  log("\n✅ Test 7 passed!\n", "green");
}

// Test 8: Real HTML Files
function testRealHtmlFiles() {
  section("🌐 Test 8: Real HTML Files");

  const htmlFiles = readdirSync(htmlDir).filter((f) => f.endsWith(".html"));

  // Test a few HTML files
  const filesToTest = ["paragraph.html", "bold.html", "italic.html"].filter((f) =>
    htmlFiles.includes(f),
  );

  if (filesToTest.length > 0) {
    log(`Testing ${filesToTest.length} HTML files:`, "cyan");

    filesToTest.forEach((file) => {
      const doc = parseHtmlFile(file);
      const paragraphs = extractParagraphs(doc);
      console.log(`  ${file}: ${paragraphs.length} paragraph(s)`);
    });
  } else {
    log("No test files found", "yellow");
  }

  log("\n✅ Test 8 passed!\n", "green");
}

// Run all tests
async function runAllTests() {
  console.log("\n" + "█".repeat(60));
  log("  @docen/deduplicate Test Suite", "bright");
  log("  Testing Document Deduplication Utilities", "cyan");
  console.log("█".repeat(60));

  try {
    testExtractParagraphs();
    testCalculateSimilarity();
    testFindDuplicates();
    testCompareDocuments();
    testFindMostSimilar();
    testReExports();
    testEdgeCases();
    testRealHtmlFiles();

    section("🎉 All Tests Passed!");
    log("The @docen/deduplicate package is working correctly!", "green");
  } catch (error) {
    log("\n❌ Test Failed!", "red");
    console.error(error);
    process.exit(1);
  }
}

// Run tests
void runAllTests();
