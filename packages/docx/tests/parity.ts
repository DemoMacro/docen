// Fixture-driven layout parity harness — test infrastructure for the
// Word-parity program. A fixture directory holds either `model.json`
// (Tiptap JSON or DocumentOptions), or a binary `input.docx`; the harness
// projects it through the production chain
//   Tiptap JSON --compileDocument--> DocumentOptions --projectDocumentOptions--> LayoutDoc
// and deep-compares the deterministic serialization against `layout.golden.json`.
//
// Regenerate goldens with `UPDATE_GOLDENS=1 pnpm exec vp test run parity`
// followed by `pnpm exec vp check --fix` (the repo formatter re-wraps short
// arrays; the harness compares structurally, so an unformatted golden still
// passes tests but fails `vp check`).
// Maintainer/oracle workflow: packages/docx/tests/fixtures/README.md.

import { existsSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import type { DocumentOptions } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";

import { compileDocument, parseDOCX } from "../src";
import { projectDocument } from "../src/layout";

/** The fixtures root (this file lives in packages/docx/tests). */
export const FIXTURES_ROOT = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
export const MODEL_FILENAME = "model.json";
export const DOCX_FILENAME = "input.docx";
export const GOLDEN_FILENAME = "layout.golden.json";
export const ORACLE_FILENAME = "oracle.pdf";
export const META_FILENAME = "meta.json";
/** Diff entries reported per mismatch before the list is truncated. */
export const DIFF_LIMIT = 40;

/** One fixture directory. `modelPath`/`docxPath` are set only when the file
 *  exists; a fixture with neither fails its check with a clear message. */
export interface ParityFixture {
  name: string;
  dir: string;
  goldenPath: string;
  modelPath?: string;
  docxPath?: string;
  /** The maintainer-provided PDF export (WPS/Word) — never read by the
   *  harness; present so reports can list which fixtures have been oracled. */
  oraclePath?: string;
  metaPath?: string;
}

export type FixtureOutcome =
  | { status: "match"; fixture: ParityFixture }
  | { status: "missing-golden"; fixture: ParityFixture }
  | { status: "updated"; fixture: ParityFixture }
  | { status: "mismatch"; fixture: ParityFixture; mismatches: ParityMismatch[]; diff: string };

export interface ParityMismatch {
  /** JSON path inside the projection, e.g. `sections[0].blocks[1].inline[0].text`. */
  path: string;
  golden: unknown;
  actual: unknown;
}

/** Walk every fixture directory under `root` (sorted by name). */
export function discoverFixtures(root: string = FIXTURES_ROOT): ParityFixture[] {
  const names = readdirSync(root, { withFileTypes: true })
    .filter((entry) => entry.isDirectory() && !entry.name.startsWith("."))
    .map((entry) => entry.name)
    .sort();
  return names.map((name) => {
    const dir = join(root, name);
    const optional = (file: string): string | undefined => {
      const path = join(dir, file);
      return existsSync(path) ? path : undefined;
    };
    return {
      name,
      dir,
      goldenPath: join(dir, GOLDEN_FILENAME),
      modelPath: optional(MODEL_FILENAME),
      docxPath: optional(DOCX_FILENAME),
      oraclePath: optional(ORACLE_FILENAME),
      metaPath: optional(META_FILENAME),
    };
  });
}

/** True for a Tiptap document (`{ type: "doc", … }`); every other model.json
 *  is treated as DocumentOptions. */
export function isTiptapDocument(model: unknown): model is JSONContent {
  return (
    typeof model === "object" && model !== null && (model as { type?: unknown }).type === "doc"
  );
}

/** The fixture's model in either direction of the converter bridge. A fixture
 *  carrying both model inputs is malformed — reject it instead of guessing. */
export async function loadFixtureModel(
  fixture: ParityFixture,
): Promise<JSONContent | DocumentOptions> {
  if (fixture.modelPath && fixture.docxPath) {
    throw new Error(
      `${fixture.name}: has both ${MODEL_FILENAME} and ${DOCX_FILENAME} — keep exactly one`,
    );
  }
  if (fixture.modelPath) {
    return JSON.parse(readFileSync(fixture.modelPath, "utf8")) as JSONContent | DocumentOptions;
  }
  if (fixture.docxPath) {
    return (await parseDOCX(readFileSync(fixture.docxPath))) as JSONContent;
  }
  throw new Error(`${fixture.name}: no ${MODEL_FILENAME} or ${DOCX_FILENAME} in the fixture`);
}

/** Project a fixture to the LayoutDoc the golden captures. */
export async function projectFixture(fixture: ParityFixture): Promise<unknown> {
  const model = await loadFixtureModel(fixture);
  const docOpts = isTiptapDocument(model) ? compileDocument(model) : model;
  return projectDocument(docOpts);
}

/** Deterministic, JSON-safe projection serialization: object keys and Map
 *  entries are sorted, Maps become plain objects, bytes become base64. */
export function serializeLayout(value: unknown): string {
  return `${JSON.stringify(toStable(value), null, 2)}\n`;
}

/** Compare a fixture's projection with its golden. `update: true` (driven by
 *  `UPDATE_GOLDENS=1`) rewrites the golden instead of failing a mismatch.
 *  Comparison is structural: a golden the repo formatter re-wrapped is a
 *  match, not a layout change. */
export async function checkFixture(
  fixture: ParityFixture,
  options: { update?: boolean } = {},
): Promise<FixtureOutcome> {
  const serialized = serializeLayout(await projectFixture(fixture));
  const goldenText = existsSync(fixture.goldenPath)
    ? readFileSync(fixture.goldenPath, "utf8")
    : undefined;

  if (goldenText !== undefined) {
    if (goldenText === serialized) return { status: "match", fixture };
    let golden: unknown;
    try {
      golden = JSON.parse(goldenText);
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      if (options.update) {
        writeFileSync(fixture.goldenPath, serialized);
        return { status: "updated", fixture };
      }
      const mismatches: ParityMismatch[] = [
        { path: "$", golden: `unparsable JSON: ${reason}`, actual: "see layout.golden.json" },
      ];
      return {
        status: "mismatch",
        fixture,
        mismatches,
        diff: formatDiff(fixture.name, mismatches),
      };
    }
    const mismatches = diffLayout(golden, JSON.parse(serialized));
    if (mismatches.length === 0) return { status: "match", fixture };
    if (options.update) {
      writeFileSync(fixture.goldenPath, serialized);
      return { status: "updated", fixture };
    }
    return { status: "mismatch", fixture, mismatches, diff: formatDiff(fixture.name, mismatches) };
  }

  if (options.update) {
    writeFileSync(fixture.goldenPath, serialized);
    return { status: "updated", fixture };
  }
  return { status: "missing-golden", fixture };
}

/** True when the environment asks for golden regeneration. */
export function updateGoldensRequested(env: NodeJS.ProcessEnv = process.env): boolean {
  return env.UPDATE_GOLDENS === "1";
}

/** Structural diff over normalized values; capped at {@link DIFF_LIMIT} so a
 *  large mismatch stays readable. */
export function diffLayout(golden: unknown, actual: unknown, limit = DIFF_LIMIT): ParityMismatch[] {
  const mismatches: ParityMismatch[] = [];
  walk(toStable(golden), toStable(actual), "", mismatches, limit);
  return mismatches;
}

/** One mismatch report per differing field, with a JSON path. The message
 *  marks the list as truncated when it hit the diff cap. */
export function formatDiff(fixtureName: string, mismatches: ParityMismatch[]): string {
  const truncated = mismatches.length >= DIFF_LIMIT;
  const lines = [
    `✗ ${fixtureName}: ${mismatches.length}${truncated ? "+" : ""} field(s) differ from ${GOLDEN_FILENAME}`,
  ];
  for (const mismatch of mismatches) {
    lines.push(
      `  at ${mismatch.path}`,
      `    golden: ${display(mismatch.golden)}`,
      `    actual: ${display(mismatch.actual)}`,
    );
  }
  if (truncated) lines.push(`  (list truncated at ${DIFF_LIMIT} fields)`);
  lines.push(`  (regenerate with UPDATE_GOLDENS=1 if the new layout is intended)`);
  return lines.join("\n");
}

function walk(
  golden: unknown,
  actual: unknown,
  path: string,
  out: ParityMismatch[],
  limit: number,
): void {
  if (out.length >= limit || Object.is(golden, actual)) return;
  if (Array.isArray(golden) && Array.isArray(actual)) {
    for (let i = 0; i < Math.max(golden.length, actual.length) && out.length < limit; i++) {
      walk(golden[i], actual[i], `${path}[${i}]`, out, limit);
    }
    return;
  }
  if (isRecord(golden) && isRecord(actual)) {
    const keys = [...new Set([...Object.keys(golden), ...Object.keys(actual)])].sort();
    for (const key of keys) {
      walk(golden[key], actual[key], path ? `${path}.${key}` : key, out, limit);
      if (out.length >= limit) return;
    }
    return;
  }
  out.push({ path: path || "$", golden, actual });
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function toStable(value: unknown): unknown {
  if (value instanceof Map) {
    const entries = [...value.entries()].sort(([a], [b]) => compareKeys(a, b));
    return Object.fromEntries(entries.map(([key, entry]) => [String(key), toStable(entry)]));
  }
  if (value instanceof Uint8Array) return { $bytes: Buffer.from(value).toString("base64") };
  if (Array.isArray(value)) return value.map((entry) => toStable(entry));
  if (isRecord(value)) {
    const stable: Record<string, unknown> = {};
    for (const key of Object.keys(value).sort()) {
      const entry = toStable(value[key]);
      if (entry !== undefined) stable[key] = entry;
    }
    return stable;
  }
  return value;
}

function compareKeys(a: unknown, b: unknown): number {
  if (typeof a === "number" && typeof b === "number") return a - b;
  const left = String(a);
  const right = String(b);
  return left < right ? -1 : left > right ? 1 : 0;
}

function display(value: unknown): string {
  if (value === undefined) return "undefined";
  const text = JSON.stringify(value) ?? "<unserializable>";
  return text.length > 200 ? `${text.slice(0, 200)}…` : text;
}
