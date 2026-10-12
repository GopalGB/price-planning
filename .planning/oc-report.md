# OpenCode runner report

Date: 2026-10-12. Branch `feat/v2-transfer`. Base for this session `cc14e1c`, head `ab90e8b`.

## STATUS

IN_PROGRESS. Every LOW finding from the exact `claude-opus-5-5` review is now closed or
dispositioned with evidence. The resource stop recorded in the previous checkpoint is cleared.
The exact clean re-review, independent planner closure, a second full Python run and fresh
browser acceptance remain open.

Resource preflight was inside the owner limits, so work proceeded: disk 6705 MB (stop below
2500), swap 4504.25 MiB (stop above 10411.31), load 3.25. Final gates ran at disk 7411 MB and
swap 3596.56 MiB. No stop was hit and no test failed for resource reasons.

## CHANGED

Five commits, each a focused unit.

`d5bc084` fix(upload): bound chart and drawing parts by relationship role.
`pp/upload.py` gives chart, chartsheet, chartSheet and drawing the tree limit and screens them
before the content-type check, so an unsafe declaration reports `unsafe_xml` instead of being
masked as `invalid_workbook`, and a missing target is translated instead of escaping as a
`KeyError`. This is inherited work that was uncommitted and unverified.

`e6d1059` fix(upload): answer an unknown provenance template hash with 422.
`_new_copy_snapshot` converted the catalogue's 500 into the same `different_template` rejection
the identity check uses. A corrupt snapshot stays a 500, since that is a server fault.

`13418ef` test(upload): cover undo, redo and undo on one upload lineage.
Passes on unmodified code, so this is coverage, not a fix.

`132c825` test(web): pin the operation cell cap and the undo input-key bypass.

`ab90e8b` docs(transfer): clear the resource stop and record the LOW dispositions.

Two defects in the inherited work were found and fixed before committing, which is why the
inherited unit was not simply validated and committed:

- The inherited role test proved only an error string. Its payload was one contiguous 9.3 MiB
  text node, which the 8 MiB token bound caught by accident, masking the missing per-part cap.
  The cap test now uses ten sibling nodes so the token bound cannot mask it, and asserts the
  per-part message. At HEAD that part reached `validate_package`, proving the gap.
- The drawing subtest never injected its relationship, because the fixture has no
  `xl/worksheets/_rels/sheet1.xml.rels`, so it screened nothing and passed for the wrong reason.
  The test now creates that part.

A probe also showed that skipping missing relationship targets would have been dead code: a
dangling target still aborts the remaining scan at the content-type check one line later. That
branch was removed and the test pins the real 422 instead of pretending otherwise.

LOW disposition, all closed:

| Finding | Resolution |
| --- | --- |
| HEAD requests have side effects | `70de160` |
| Bad confirm input burns the preview | `b1c7339` |
| Non-numeric content-length gives a 500 | `2395420` |
| Percent-encoded targets, ZIP directories rejected | `a49e195` |
| Self-closing defined names merge | `a49e195` |
| Literal `_x000D_` not escaped | `a49e195` |
| `.xlsm` export of an `.xlsx` template | `a49e195` |
| Spoofable compression ratio | `cc14e1c` |
| Role coverage for chart/drawing | `d5bc084` |
| Unknown template hash may 500 | `e6d1059` |
| Shared-string text picks up phonetic runs | already fixed, covered by `test_template_importer.py:370` and `test_transfer_upload.py:666` |
| Redo of an upload undo untested | `13418ef`, coverage gap not a bug |
| Client cell limit relaxed too broadly | non-issue, pinned by `132c825`; `authorizedInputs` has one call site in `invertOp` and paste caps independently at `grid-data.mjs:201` |

## TESTS

Red captured before each fix, at HEAD source with the new tests in place:

- Oversized 9.3 MiB chart part passed every upload bound and reached `validate_package`.
- Crafted provenance (nonexistent copy plus unknown hash) returned `500 template_missing`.
- Three inherited role subtests returned `invalid_workbook` instead of `unsafe_xml`.

Green, twice for the focused sets:

- focused role tests 4/4 twice (1.138s, 1.086s)
- unknown template hash 2/2 twice (0.308s, 0.290s)
- upload module 60/60 (17.644s)
- related transfer, adversarial, export, security 81/81 (17.682s)
- web 109/109 (4.650s); operations file 13/13
- full Python 421/421 (80.634s)

Gates:

- changed-file Ruff clean
- full-tree Ruff 141 findings, 0 introduced and 0 removed signatures against
  `results-private/transfer/ruff-baseline.json`
- `ops/content_guard.py --all` exit 0
- `npm run format:check` pass; the Prettier gate covers `src/*.mjs` only, and
  `web/test/operations.test.mjs` was already non-compliant at baseline, so it was not reformatted
- full Python run once only; the XML timing assertion is load-sensitive and previously flaked
  under heavy load, so a second run is still owed

## NEXT

Open, in order:

1. Exact clean `claude-opus-5-5` re-review of the corrected source. Skipped this session: the
   runner instruction forbids calling external AI CLIs, and review is the supervisor's job.
2. Independent planner closure.
3. Second full Python run on the corrected source, ideally with low host load.
4. Fresh single-TaskSpace browser acceptance. The earlier browser receipt is stale because both
   upload and web source changed after it. No runtime or TaskSpace was started this session.
5. Three HIGH hypotheses from the review (self-closing cell freezing, upload metadata memory
   bounds, stale-upload two-way merge) are recorded as addressed by earlier commits but were not
   re-verified this session.

No push, merge or deployment. Dependency symlinks `.venv` and `web/node_modules` remain
untracked. No real or private workbook, Docker or dependency change.

One observation for the reviewer, not fixed because it is outside this session's findings: in
`invertOp`, a malformed op whose `cells` are not covered by `before` throws a bare TypeError
rather than a validation error. It is reachable only from a hand-built op object, not from
server or grid state, so it was left alone rather than widening scope.