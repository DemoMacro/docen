# Word-parity program

docen's fidelity target is pixel parity with Word/WPS page by page. This
directory documents the parity workflow; the executable part lives in the
fixture-driven harness at `packages/docx/tests/parity.ts` (spec:
`packages/docx/tests/parity.spec.ts`, fixtures:
`packages/docx/tests/fixtures/`).

## The parity loop

1. **Capture** — reduce a real document (or a reported mismatch) to a minimal
   fixture under `packages/docx/tests/fixtures/<name>/`. A fixture holds a
   `model.json` (Tiptap JSON or DocumentOptions) or an `input.docx`.
2. **Project** — the harness runs the same chain the editor runs:
   `compileDocument` → `projectDocumentOptions`, then serializes the LayoutDoc
   deterministically.
3. **Golden** — the serialized projection is stored as
   `layout.golden.json` and guards every later change:

   ```bash
   UPDATE_GOLDENS=1 pnpm exec vp test run parity   # write/refresh goldens
   pnpm exec vp check --fix                         # normalize golden formatting
   pnpm exec vp test run parity                     # verify
   ```

   A changed golden is a behavior change: review the diff, and only regenerate
   when the new geometry is intended (and, for oracled fixtures, matches the
   PDF). `UPDATE_GOLDENS=1` writes 2-space JSON while the repo formatter
   re-wraps short arrays — the harness compares structurally, but `vp check`
   needs the `--fix` pass.

4. **Oracle** — for the fixtures that must match a real Word processor, export
   the document to `oracle.pdf` from the pinned WPS/Word build and record the
   build in `meta.json`. See
   `packages/docx/tests/fixtures/README.md` for the exact steps and the
   `meta.json` template.
5. **Compare pages** — page-by-page PDF comparison is a manual/integration step
   of the program (the harness is headless); the projection golden catches the
   regressions automatically in the meantime.

## Commands

```bash
pnpm exec vp test run parity                  # fixture walker + harness contract tests
UPDATE_GOLDENS=1 pnpm exec vp test run parity # regenerate goldens (idempotent)
pnpm exec vp check --fix                      # normalize golden formatting after a write
pnpm exec vp test bench                       # projection benchmark (medium document)
```

Fixtures without a golden are skipped with a `⏭` note — run with
`UPDATE_GOLDENS=1` to create them.

## What the golden covers

The golden is the **projection** (LayoutDoc input): resolved style cascade,
unit conversions, page flow box, table geometry/spans, footnote definitions,
field atoms, page furniture. Pagination and painting are downstream engine
stages with their own suites (`packages/layout/src/flow/flow.spec.ts`,
`packages/core`); a projection golden that is green while pagination drifts is
possible, so parity features should add flow/paint assertions where relevant.

## Adding a parity feature

Use `checklist-template.md` as the review artifact for each feature. The
minimum bar for "parity" on a feature is:

- [ ] at least one fixture covering the writing-relevant shape
- [ ] `UPDATE_GOLDENS=1` run and the golden diff reviewed
- [ ] oracle PDF + `meta.json` when a Word/WPS build is available
- [ ] visible behavior differences called out in the feature checklist
