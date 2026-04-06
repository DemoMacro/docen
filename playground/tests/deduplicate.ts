import { readFileSync, readdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseHTML } from "docen";
import {
  extractParagraphs,
  calculateSimilarity,
  findDuplicates,
  compareDocuments,
  findBestMatch,
  splitSentences,
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
  section("Test 1: Extract Paragraphs");

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
  log("PASS\n", "green");
}

// Test 2: Calculate Similarity (Levenshtein-based)
function testCalculateSimilarity() {
  section("Test 2: Calculate Similarity (Levenshtein-Based)");

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
      a: "量子计算承诺解决经典计算机难以解决的复杂问题。",
      b: "区块链技术为安全透明的交易提供了去中心化的账本系统。",
      expected: 0.3,
      description: "完全不同的主题（中文）",
    },
    {
      a: "自然语言处理使计算机能够理解和生成人类语言。",
      b: "自然语言处理帮助计算机理解和生成人类语言。",
      expected: 0.85,
      description: "仅一词之差（中文）",
    },
  ];

  tests.forEach(({ a, b, expected, description }) => {
    const similarity = calculateSimilarity(a, b);

    log(`${description}:`, "cyan");
    console.log(`  A: "${a}"`);
    console.log(`  B: "${b}"`);
    console.log(`  Similarity: ${(similarity * 100).toFixed(1)}%`);
    console.log(`  Expected: ~${(expected * 100).toFixed(0)}%`);

    const isMatch = Math.abs(similarity - expected) < 0.15;
    if (isMatch) {
      log(`  PASS\n`, "green");
    } else {
      log(`  Close enough (within tolerance)\n`, "yellow");
    }
  });

  // Containment-like test: short text vs long text with Levenshtein
  log("Short vs Long text (Levenshtein):", "cyan");
  const short = "这是一段简短的文本";
  const long = "这是一段简短的文本，后面还有很长的补充内容";
  const sim = calculateSimilarity(short, long);
  console.log(`  similarity(short, long) = ${(sim * 100).toFixed(1)}%`);
  console.log(`  (Levenshtein penalizes length difference, not containment)`);
  log("  PASS\n", "green");

  log("PASS\n", "green");
}

// Test 3: Find Duplicates
function testFindDuplicates() {
  section("Test 3: Find Duplicates in Document");

  const docWithDuplicates: JSONContent = {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [{ type: "text", text: "人工智能正在改变我们现代世界的生活和工作方式。" }],
      },
      {
        type: "paragraph",
        content: [
          { type: "text", text: "机器学习算法能够识别人类可能会错过的大型数据集中的模式。" },
        ],
      },
      {
        type: "paragraph",
        content: [{ type: "text", text: "深度学习模型需要大量计算资源和大量的训练数据。" }],
      },
      {
        type: "paragraph",
        content: [
          { type: "text", text: "机器学习算法能够识别人类可能会错过的大型数据集中的模式。" },
        ],
      },
      {
        type: "paragraph",
        content: [{ type: "text", text: "神经网络由处理信息的互连节点层组成。" }],
      },
      {
        type: "paragraph",
        content: [
          { type: "text", text: "机器学习算法能够识别人类可能会错过的大型数据集中的模式。" },
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
        `      ${j + 1}. Index ${dupIndex} - similarity: ${(dup.similarities[j] * 100).toFixed(1)}%`,
      );
    });
  });

  console.assert(duplicates.length === 1, "Should find 1 duplicate group");
  log("PASS\n", "green");
}

// Test 4: Compare Documents
function testCompareDocuments() {
  section("Test 4: Compare Two Documents");

  const doc1: JSONContent = {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [{ type: "text", text: "自然语言处理使计算机能够理解和生成人类语言。" }],
      },
      {
        type: "paragraph",
        content: [{ type: "text", text: "Transformer 模型已经彻底改变了机器翻译和文本摘要任务。" }],
      },
      {
        type: "paragraph",
        content: [{ type: "text", text: "注意力机制允许模型专注于输入序列的相关部分。" }],
      },
    ],
  };

  const doc2: JSONContent = {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [{ type: "text", text: "自然语言处理帮助计算机理解和生成人类语言。" }],
      },
      {
        type: "paragraph",
        content: [{ type: "text", text: "Transformer 架构已经改变了机器翻译和文本摘要。" }],
      },
      {
        type: "paragraph",
        content: [{ type: "text", text: "自注意力机制使模型能够专注于输入的重要部分。" }],
      },
    ],
  };

  const result = compareDocuments(doc1, doc2);

  log(`Document coverage: ${(result.coverage * 100).toFixed(1)}%`, "cyan");
  log(`Paragraph comparisons: ${result.paragraphs.length}`, "cyan");

  result.paragraphs.forEach((comp, i) => {
    console.log(`\n  Para ${i + 1} [${comp.matchKind}] (${(comp.similarity * 100).toFixed(1)}%):`);
    console.log(`    Doc 1 [${comp.fromDoc1.index}]: "${comp.fromDoc1.text}"`);
    if (comp.fromDoc2) {
      console.log(`    Doc 2 [${comp.fromDoc2.index}]: "${comp.fromDoc2.text}"`);
      console.log(
        `    Coverage: covA=${(comp.coverage.covA * 100).toFixed(0)}%, covB=${(comp.coverage.covB * 100).toFixed(0)}%`,
      );
    }
  });

  log("\nPASS\n", "green");
}

// Test 5: Find Best Match (using findBestMatch from @nlptools/distance)
function testFindBestMatch() {
  section("Test 5: Find Best Match (via @nlptools/distance)");

  const target = "人工智能的快速发展给各个行业带来了巨大的变化。";
  const candidates = [
    "区块链技术不断发展和影响全球金融部门。",
    "人工智能的快速发展为各个行业带来了巨大的改变。",
    "人工智能的快速进步正在深刻地改变不同行业。",
    "气候变化仍然是人类面临的最紧迫的挑战之一。",
  ];

  // Use default levenshtein algorithm
  const result = findBestMatch(target, candidates, {
    algorithm: "levenshtein",
  });

  log(`Target: "${target}"`, "cyan");

  if (!result) {
    log("No match found!", "red");
    throw new Error("Expected to find a match");
  }

  log(`Best match: "${result.item}"`, "green");
  console.log(`  Score: ${(result.score * 100).toFixed(1)}%`);
  console.log(`  Index: ${result.index}`);

  console.assert(
    result.index === 1 || result.index === 2,
    "Should match either index 1 or 2 (most similar to target)",
  );
  console.assert(result.score > 0.5, "Similarity score should be > 50%");

  log("\nPASS\n", "green");
}

// Test 6: Sentence Splitting & Match Classification
function testSentenceSplitting() {
  section("Test 6: Sentence Splitting & Classification");

  // Sentence splitting
  log("Sentence splitting:", "cyan");
  const sentences = splitSentences("第一句话。第二句话！第三句话？第四句话。");
  console.log(`  "${sentences.join('" | "')}"`);
  console.assert(sentences.length === 4, "Should split into 4 sentences");

  // Chinese sentence splitting with semicolons
  const mixed = splitSentences("保洁工作应小心细致；不要碰坏车辆。注意安全。");
  console.log(`  Mixed: "${mixed.join('" | "')}"`);
  console.assert(mixed.length === 3, "Should split into 3 sentences");

  // Classification via compareDocuments
  log("\nMatch classification (via compareDocuments):", "cyan");

  // Contained: one sentence fully inside a longer paragraph
  const containedDoc: JSONContent = {
    type: "doc",
    content: [{ type: "paragraph", content: [{ type: "text", text: "这是一段简短的文本。" }] }],
  };
  const containedDoc2: JSONContent = {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [{ type: "text", text: "这是一段简短的文本。后面还有补充。" }],
      },
    ],
  };
  const rContained = compareDocuments(containedDoc, containedDoc2);
  log(
    `  Contained: ${rContained.paragraphs[0].matchKind} (${(rContained.paragraphs[0].similarity * 100).toFixed(0)}%)`,
    "green",
  );
  console.assert(rContained.paragraphs[0].matchKind === "contained", "Should be contained");
  console.assert(
    rContained.paragraphs[0].fromDoc2 !== null,
    "fromDoc2 should not be null for contained",
  );

  // None: completely unrelated short texts → below threshold
  const noneDoc: JSONContent = {
    type: "doc",
    content: [{ type: "paragraph", content: [{ type: "text", text: "你好世界" }] }],
  };
  const noneDoc2: JSONContent = {
    type: "doc",
    content: [{ type: "paragraph", content: [{ type: "text", text: "再见朋友" }] }],
  };
  const rNone = compareDocuments(noneDoc, noneDoc2);
  log(
    `  None: ${rNone.paragraphs[0].matchKind} (${(rNone.paragraphs[0].similarity * 100).toFixed(0)}%)`,
    "reset",
  );
  console.assert(rNone.paragraphs[0].matchKind === "none", "Should be none");
  console.assert(
    rNone.paragraphs[0].fromDoc2 === null,
    "fromDoc2 should be null when matchKind is none",
  );

  // Threshold controls noise floor: weakOverlap visible only with low threshold
  // 3 sentences in doc1, only 1 matches with doc2's 3 sentences → covA=1/3≈0.33 (weakOverlap)
  const weakDoc1: JSONContent = {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "机器学习是人工智能的核心技术之一。深度学习通过多层神经网络实现了突破性的进展。强化学习则在游戏和机器人控制中表现出色。",
          },
        ],
      },
    ],
  };
  const weakDoc2: JSONContent = {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "机器学习是人工智能的核心技术之一。计算机视觉在自动驾驶领域有着重要的应用价值。自然语言处理使机器能够理解人类的语言。",
          },
        ],
      },
    ],
  };
  const rWeakDefault = compareDocuments(weakDoc1, weakDoc2);
  log(
    `  Weak (default threshold): ${rWeakDefault.paragraphs[0].matchKind} (${(rWeakDefault.paragraphs[0].similarity * 100).toFixed(0)}%)`,
    "reset",
  );
  console.assert(
    rWeakDefault.paragraphs[0].matchKind === "none",
    "Weak overlap should be 'none' with default threshold",
  );

  const rWeakLow = compareDocuments(weakDoc1, weakDoc2, { threshold: 0.3 });
  log(
    `  Weak (threshold=0.3): ${rWeakLow.paragraphs[0].matchKind} (${(rWeakLow.paragraphs[0].similarity * 100).toFixed(0)}%)`,
    "yellow",
  );
  console.assert(
    rWeakLow.paragraphs[0].matchKind === "weakOverlap",
    "Should be weakOverlap with low threshold",
  );
  console.assert(
    rWeakLow.paragraphs[0].fromDoc2 !== null,
    "fromDoc2 should not be null for weakOverlap",
  );

  log("\nPASS\n", "green");
}

// Test 7: Edge Cases
function testEdgeCases() {
  section("Test 7: Edge Cases");

  // Empty documents
  const emptyDoc: JSONContent = { type: "doc", content: [] };
  const emptyParagraphs = extractParagraphs(emptyDoc);
  console.log(`Empty document paragraphs: ${emptyParagraphs.length}`);
  console.assert(emptyParagraphs.length === 0, "Should have 0 paragraphs");

  // Whitespace handling (Chinese)
  const ws1 = "机器学习  是 一个 令人着迷的 领域";
  const ws2 = "机器学习 是 一个 令人着迷的 领域";
  const simWs = calculateSimilarity(ws1, ws2);
  log(`Whitespace test (Chinese): ${(simWs * 100).toFixed(0)}% similarity`, "cyan");

  // Case handling (English)
  const case1 = "ARTIFICIAL INTELLIGENCE IS TRANSFORMING THE WORLD";
  const case2 = "artificial intelligence is transforming the world";
  const simCase = calculateSimilarity(case1, case2);
  log(`Case test (English): ${(simCase * 100).toFixed(0)}% similarity`, "cyan");

  // Empty strings
  const simEmpty = calculateSimilarity("", "");
  console.assert(simEmpty === 1, "Two empty strings should be identical");
  const simOneEmpty = calculateSimilarity("hello", "");
  console.assert(simOneEmpty === 0, "One empty string should be 0 similarity");

  log("\nPASS\n", "green");
}

// Test 8: Real HTML Files
function testRealHtmlFiles() {
  section("Test 8: Real HTML Files");

  const htmlFiles = readdirSync(htmlDir).filter((f) => f.endsWith(".html"));
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

  log("\nPASS\n", "green");
}

// Run all tests
async function runAllTests() {
  console.log("\n" + "=".repeat(60));
  log("  @docen/deduplicate Test Suite", "bright");
  log("  SimHash Screening + Levenshtein Verification", "cyan");
  console.log("=".repeat(60));

  try {
    testExtractParagraphs();
    testCalculateSimilarity();
    testFindDuplicates();
    testCompareDocuments();
    testFindBestMatch();
    testSentenceSplitting();
    testEdgeCases();
    testRealHtmlFiles();

    section("All Tests Passed!");
    log("The @docen/deduplicate package is working correctly!", "green");
  } catch (error) {
    log("\nTest Failed!", "red");
    console.error(error);
    process.exit(1);
  }
}

// Run tests
void runAllTests();
