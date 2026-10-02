# Parity checklist: <feature>

Copy this file to `docs/parity/checklists/<feature>.md` (create the directory
on first use) and fill it in while implementing the feature. One checklist per
writing-relevant feature; link the fixture(s) and any oracle.

## Scope

- Feature: <what Word behavior this covers, e.g. "table row split across pages">
- Fixtures: `packages/docx/tests/fixtures/<name>/`
- Related code: `<module(s) under packages/docx/src/layout/project or packages/layout>`

## Corpus

- [ ] `model.json` fixture minimized from a real document or reported mismatch
- [ ] `input.docx` fixture only if the parse path itself is under test
- [ ] Fixture covers the interesting edge (merged cells, split row, field,
      footnote, grid pitch, …), not just the happy path
- [ ] `layout.golden.json` generated with `UPDATE_GOLDENS=1` and reviewed

## Oracle

- [ ] `oracle.pdf` exported from the pinned WPS/Word build
- [ ] `meta.json` records application + exact version + platform + date
- [ ] Page-by-page spot check done: rendered pages vs. PDF pages
- [ ] Differences are upstream bugs or accepted deviations (list them below)

If no oracle exists yet, state why and leave the boxes unchecked — the golden
still guards the projection.

## Verification

- [ ] `pnpm exec vp test run parity` green
- [ ] `pnpm exec vp test run` introduces no new failures (baseline at the time
      of writing: 9 pre-existing failures in
      `packages/editor/src/document/canvas/caret-map.spec.ts`)
- [ ] `pnpm exec vp check` green (0 errors; run `--fix` after
      `UPDATE_GOLDENS=1`)
- [ ] Golden diff contains only the intended geometry changes
- [ ] `pnpm exec vp test bench` shows no projection regression (when touching
      hot paths)

## Notes

### Accepted deviations from the oracle

| Page | Difference | Reason / follow-up |
| ---- | ---------- | ------------------ |
|      |            |                    |

### Follow-ups

- <issues discovered that belong to another feature/PR>
