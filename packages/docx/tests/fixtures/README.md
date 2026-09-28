# Layout parity fixtures

Each subdirectory is one parity fixture. The harness
(`packages/docx/tests/parity.ts`, spec `packages/docx/tests/parity.spec.ts`)
projects the fixture through the production chain and compares the result with
the fixture's golden:

```
model.json (Tiptap JSON or DocumentOptions) ─┐
                                             ├─> compileDocument ─┐
input.docx ──parseDOCX──> Tiptap JSON ───────┘                    ├─> projectDocumentOptions ─> LayoutDoc
                                                                  ┘
LayoutDoc ──serialize deterministically──> layout.golden.json
```

## Fixture layout

```
fixtures/<name>/
  model.json           either Tiptap JSON (`{ "type": "doc", … }`) or
                       DocumentOptions (`{ "sections": [ … ] }`); preferred
  input.docx           alternative binary input (parsed with parseDOCX)
  layout.golden.json   the projected LayoutDoc, deterministic JSON
  oracle.pdf           maintainer-provided PDF export (never read by the harness)
  meta.json            provenance of oracle.pdf (see below)
```

A fixture must contain exactly one of `model.json` / `input.docx`; the harness
rejects a directory with both. Prefer `model.json`: it is human-editable,
diffable, and binary-free. Use `input.docx` only when the fixture must prove
the parse path itself.

`layout.golden.json` is generated, never hand-edited:

```bash
# regenerate every golden (a matching golden is left untouched)
UPDATE_GOLDENS=1 pnpm exec vp test run parity
# then normalize formatting — UPDATE_GOLDENS writes 2-space JSON, the repo
# formatter re-wraps short arrays
pnpm exec vp check --fix

# or a single fixture
UPDATE_GOLDENS=1 pnpm exec vp test run parity -t "paragraph-heading-runs"

# verify without regenerating
pnpm exec vp test run parity
```

> `UPDATE_GOLDENS=1` writes deterministic 2-space JSON, but the repo formatter
> re-wraps short arrays (`"columnWidthsPx": [160, 160, 160]`), so a freshly
> regenerated golden can fail `vp check` even though the harness compares
> structurally and passes. Always run `pnpm exec vp check --fix` after
> regenerating; `UPDATE_GOLDENS=1` itself leaves semantically matching goldens
> untouched.

A fixture without a golden is **skipped, not failed** — the spec logs
`⏭ <name>: no layout.golden.json` and continues. Run with `UPDATE_GOLDENS=1`
to generate it. Review golden diffs like code: a changed golden is a behavior
change and must be justified against the oracle PDF.

## Oracle workflow (Word/WPS parity)

The harness runs headless; the oracle is a **manual** maintainer step. For a
fixture that must match a real Word processor:

1. Open the fixture in the pinned WPS Office build (record the exact version;
   do not upgrade mid-comparison — Word/WPS change pagination between builds).
2. Export the whole document to PDF with `File → Export → PDF`, no scaling,
   no "shrink to fit", no extra watermark.
3. Save the export next to the fixture as `oracle.pdf`.
4. Save the provenance as `meta.json`:

```json
{
  "oracle": {
    "application": "WPS Office",
    "version": "12.1.0.<build>",
    "platform": "Windows 11 24H2",
    "exportedAt": "2026-09-16",
    "exportedBy": "<maintainer>",
    "notes": "File → Export → PDF; no scaling; SimSun/Times New Roman installed"
  }
}
```

Missing `oracle.pdf`/`meta.json` is fine — the fixture still guards the
projection golden, and there is simply nothing to compare page-by-page yet.
`oracle.pdf` is never opened by the automated suite; page-by-page comparison
(rendered pages vs. PDF pages) is the follow-up step of the parity program.

## Adding a fixture

1. Create `fixtures/<name>/` with a minimized `model.json` that reproduces the
   layout case (one behavior per fixture; keep it small enough to review).
2. Generate the golden: `UPDATE_GOLDENS=1 pnpm exec vp test run parity`.
3. Inspect `layout.golden.json`: check page/flow geometry, spacing, indents,
   run styles, table spans. If it looks wrong, fix production code, not the
   golden.
4. If a Word/WPS oracle exists, add `oracle.pdf` + `meta.json` per above.
5. Record the feature in a parity checklist (see
   `docs/parity/checklist-template.md`).

## Current fixtures

| Fixture                  | Covers                                                                                                                                                                            |
| ------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `paragraph-heading-runs` | Style cascade (Normal → Heading1), exact/multiple line rules, indent forms, justification/centering, bold/italic/underline/strike/superscript/color/size/highlight/font overrides |
| `table-merged-cells`     | Horizontal span (`columnSpan`), vertical merge expansion (`verticalMerge` → `rowspan`), header row, `cantSplit`, row height, table borders/margins/alignment                      |
| `footnote-page-field`    | Footnote references + definitions with first-reference ordinals, `PAGE`/`NUMPAGES` dynamic atoms, `CREATEDATE` cached field, footer furniture slot                                |
