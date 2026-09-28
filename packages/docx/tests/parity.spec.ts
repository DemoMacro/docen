// The fixture walker + harness contract tests. Every directory under
// tests/fixtures is projected and compared against its own layout.golden.json;
// fixtures without a golden are skipped with a console note (regenerate with
// UPDATE_GOLDENS=1). The second describe covers the harness itself (missing
// golden skips, mismatch diff, both model branches) against temporary fixtures
// under tests/.temp/parity-spec.

import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { generateDOCXSync } from "../src";
import {
  DIFF_LIMIT,
  DOCX_FILENAME,
  GOLDEN_FILENAME,
  MODEL_FILENAME,
  checkFixture,
  diffLayout,
  discoverFixtures,
  formatDiff,
  updateGoldensRequested,
} from "./parity";

const update = updateGoldensRequested();
const fixtures = discoverFixtures();

describe("layout parity fixtures", () => {
  it("ships at least three fixtures with a model", () => {
    expect(fixtures.length).toBeGreaterThanOrEqual(3);
    for (const fixture of fixtures) {
      expect(fixture.modelPath ?? fixture.docxPath, `${fixture.name}: model present`).toBeDefined();
    }
  });

  for (const fixture of fixtures) {
    it(`${fixture.name} matches ${GOLDEN_FILENAME}`, async () => {
      const outcome = await checkFixture(fixture, { update });
      if (outcome.status === "missing-golden") {
        console.log(
          `⏭ ${fixture.name}: no ${GOLDEN_FILENAME} — skipping (UPDATE_GOLDENS=1 regenerates)`,
        );
        return;
      }
      if (outcome.status === "updated") {
        console.log(
          `✍ ${fixture.name}: wrote ${GOLDEN_FILENAME} (run pnpm exec vp check --fix to format)`,
        );
        return;
      }
      if (outcome.status === "mismatch") throw new Error(outcome.diff);
      expect(outcome.status).toBe("match");
    });
  }
});

// ── harness contract ──

const tempRoot = join(dirname(fileURLToPath(import.meta.url)), ".temp", "parity-spec");
const MINIMAL_MODEL = JSON.stringify({
  sections: [{ children: [{ paragraph: { text: "hello" } }] }],
});
const TIPTAP_MODEL = JSON.stringify({
  type: "doc",
  content: [{ type: "paragraph", content: [{ type: "text", text: "hi" }] }],
});

function writeFixture(name: string, files: Record<string, string>): void {
  const dir = join(tempRoot, name);
  mkdirSync(dir, { recursive: true });
  for (const [file, contents] of Object.entries(files)) {
    writeFileSync(join(dir, file), contents);
  }
}

function fixtureOf(name: string) {
  const fixture = discoverFixtures(tempRoot).find((entry) => entry.name === name);
  if (!fixture) throw new Error(`temp fixture ${name} not found`);
  return fixture;
}

describe.sequential("parity harness", () => {
  it("skips a fixture without a golden instead of failing", async () => {
    writeFixture("missing-golden", { [MODEL_FILENAME]: MINIMAL_MODEL });
    const outcome = await checkFixture(fixtureOf("missing-golden"));
    expect(outcome.status).toBe("missing-golden");
  });

  it("rejects a fixture carrying both model inputs", async () => {
    writeFixture("both-inputs", {
      [MODEL_FILENAME]: MINIMAL_MODEL,
      [DOCX_FILENAME]: "not-a-docx",
    });
    await expect(checkFixture(fixtureOf("both-inputs"))).rejects.toThrow(/keep exactly one/);
  });

  it("UPDATE_GOLDENS writes the golden, then the fixture matches", async () => {
    writeFixture("update-golden", { [MODEL_FILENAME]: MINIMAL_MODEL });
    const fixture = fixtureOf("update-golden");
    expect(await checkFixture(fixture, { update: true })).toMatchObject({ status: "updated" });
    expect(existsSync(fixture.goldenPath)).toBe(true);
    expect(await checkFixture(fixture)).toMatchObject({ status: "match" });
  });

  it("treats a formatter-rewrapped golden as a match and does not rewrite it", async () => {
    writeFixture("format-drift", { [MODEL_FILENAME]: MINIMAL_MODEL });
    const fixture = fixtureOf("format-drift");
    await checkFixture(fixture, { update: true });
    // Re-wrap the JSON like the repo formatter does (whitespace only).
    const rewrapped = readFileSync(fixture.goldenPath, "utf8").replace(/\s+/g, " ");
    writeFileSync(fixture.goldenPath, rewrapped);
    expect(await checkFixture(fixture)).toMatchObject({ status: "match" });
    // UPDATE_GOLDENS stays a no-op: the golden is not semantically stale.
    expect(await checkFixture(fixture, { update: true })).toMatchObject({ status: "match" });
    expect(readFileSync(fixture.goldenPath, "utf8")).toBe(rewrapped);
  });

  it("reports a mismatch as a readable path/value diff", async () => {
    writeFixture("mismatch", { [MODEL_FILENAME]: MINIMAL_MODEL });
    const fixture = fixtureOf("mismatch");
    await checkFixture(fixture, { update: true });
    const golden = JSON.parse(readFileSync(fixture.goldenPath, "utf8")) as {
      sections: { blocks: { inline: { text: string }[] }[] }[];
    };
    golden.sections[0]!.blocks[0]!.inline[0]!.text = "changed";
    writeFileSync(fixture.goldenPath, `${JSON.stringify(golden, null, 2)}\n`);

    const outcome = await checkFixture(fixture);
    expect(outcome.status).toBe("mismatch");
    if (outcome.status !== "mismatch") return;
    expect(outcome.diff).toContain("sections[0].blocks[0].inline[0].text");
    expect(outcome.diff).toContain('golden: "changed"');
    expect(outcome.diff).toContain('actual: "hello"');
    expect(outcome.mismatches).toHaveLength(1);
  });

  it("projects a Tiptap model through compileDocument", async () => {
    writeFixture("tiptap-model", { [MODEL_FILENAME]: TIPTAP_MODEL });
    const fixture = fixtureOf("tiptap-model");
    await checkFixture(fixture, { update: true });
    const golden = JSON.parse(readFileSync(fixture.goldenPath, "utf8")) as {
      sections: { blocks: { inline: { kind: string; text: string }[] }[] }[];
    };
    expect(golden.sections[0]!.blocks[0]!.inline[0]).toMatchObject({ kind: "text", text: "hi" });
  });

  it("projects an input.docx fixture through parseDOCX", async () => {
    const dir = join(tempRoot, "docx-input");
    mkdirSync(dir, { recursive: true });
    const bytes = generateDOCXSync({
      type: "doc",
      content: [{ type: "paragraph", content: [{ type: "text", text: "docx" }] }],
    });
    writeFileSync(join(dir, DOCX_FILENAME), bytes);
    const fixture = fixtureOf("docx-input");
    await checkFixture(fixture, { update: true });
    const golden = JSON.parse(readFileSync(fixture.goldenPath, "utf8")) as {
      sections: { blocks: { inline: { kind: string; text: string }[] }[] }[];
    };
    expect(golden.sections[0]!.blocks[0]!.inline[0]).toMatchObject({ kind: "text", text: "docx" });
  });

  it("marks the diff as truncated at the cap", () => {
    const size = DIFF_LIMIT + 20;
    const golden = Object.fromEntries(Array.from({ length: size }, (_, i) => [`k${i}`, 1]));
    const actual = Object.fromEntries(Array.from({ length: size }, (_, i) => [`k${i}`, 2]));
    const mismatches = diffLayout(golden, actual);
    expect(mismatches).toHaveLength(DIFF_LIMIT);
    const text = formatDiff("cap", mismatches);
    expect(text).toContain(`${DIFF_LIMIT}+ field(s) differ`);
    expect(text).toContain(`truncated at ${DIFF_LIMIT} fields`);
  });

  it("cleans up the temporary fixtures", () => {
    rmSync(tempRoot, { recursive: true, force: true });
    expect(existsSync(tempRoot)).toBe(false);
  });
});
