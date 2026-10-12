# OpenCode runner report

Date: 2026-10-12. Branch `fix/bandit-hardening`. Base `d2a45ca`, head `8394df1`, three commits.

## STATUS

DONE for the task I picked up: CI's security job runs `bandit -r app ops -ll`, and it was
exiting 1 on 18 MEDIUM findings. It now exits 0 with no suppression beyond one documented
`nosec`. The underlying XML defect was real, not a lint nit: entity expansion was live.

Superseded: this report replaces the earlier `feat/v2-transfer` content at this path.

## CHANGED

Three commits, each one finding class, each verified before commit.

`51e301e` fix(security): parse workbook parts with defusedxml in app and ops.
Closed 11 B314 findings across `app/workbook.py`, `ops/find_inputs.py` and
`ops/fixtures/dashboard.py`. These were the substantive ones. `defusedxml` was already a
locked, vendored dependency from `d2a45ca`; this extends its use to the parsers CI scans.
The fixture module keeps stdlib `ElementTree` for building and serializing, because
`defusedxml` does not re-export `Element`/`SubElement`; only its parse entry point moved.

`76690a1` fix(security): name the all-interfaces bind constant, closing 6 B104 findings.
Bandit B104 fires on the literal `0.0.0.0` wherever it appears, including the two
comparisons that *reject* it as a public origin, so one deliberate deployment choice
produced six findings. The literal now lives in one named `ALL_INTERFACES` constant
carrying the justification and a scoped `nosec B104`. Behaviour unchanged.

`8394df1` fix(security): write lock and table-wipe SQL as complete statements, closing B608.
Both findings concatenated fixed literals, so nothing was injectable, but bandit reads a
`BinOp` and cannot see that. `app/db.py` now selects one of two whole statements instead of
appending a row-lock suffix, and the import wipe lists four complete statements instead of
building them from table names. No suppression was added here.

### Red evidence, captured before each fix

The XML defect was reproduced, not assumed. A 200 kB literal nested through six entity
levels expanded 64x (to 12.8 MB) in process memory before expat's amplification ceiling
refused it. A tighter bomb, sitting under that ceiling, parsed silently.

New tests in `tests/test_p0_regressions.py` (class `XmlEntityRefusal`) pin refusal at both
entry points. Both failed before the change:
- `find_inputs._unlocked_style_ids` accepted an entity-bearing styles part.
- `import_workbook` parsed the whole workbook and printed its sheet before the assertion.

After the change both raise `EntitiesForbidden`/`DTDForbidden`.

## TESTS

- `bandit -r app ops -ll`: exit 1 before, exit 0 after, 0 MEDIUM and 0 HIGH remaining.
- Full Python: 754 of 754 twice (187.967 s, 185.953 s). 752 before this branch's tests.
- `ops/content_guard.py --all`: exit 0.
- Ruff, whole tree, `app ops tests`: 698 findings now and 698 at the `d2a45ca` baseline,
  0 introduced and 0 removed, compared per file and rule in a scratch worktree.
- Focused: `tests.test_core`, `tests.test_demo`, `tests.test_p0_regressions` 36/36;
  `tests.test_runtime`, `tests.test_launchers_v2` 13/13.

Not run: web suite (no `web/node_modules` in this worktree; no `web/` file was touched),
browser acceptance, Office/Excel, real workbook, semgrep, pip-audit, gitleaks. No external
model review was attempted; the runner instruction forbids it and review is the supervisor's.

## Scope I checked and deliberately did not change

`pp/` is not in the CI bandit scan and still carries 39 findings, mostly B314. I checked
whether that is a live hole rather than assuming either way, by driving real entity bombs
through `_validate_archive` on the synthetic fixture.

It is already defended. `_validate_xml_members` walks every XML member with an expat parser
whose `StartDoctypeDeclHandler` is `_unsafe_xml`, and that screen runs before any
`ET.fromstring`/`iterparse` in the upload path. Bombing `[Content_Types].xml` and
`docProps/custom.xml` returned `422 unsafe_xml`; `xl/workbook.xml` and `xl/styles.xml`
returned `422 invalid_workbook`. `pp/upload.py` and `pp/export/transfer.py` were already
converted in `d2a45ca`, and `_read_workbook` is only reachable after `_validate_archive`.

So the remaining `pp/` findings are not an unmitigated vulnerability. Widening the scan to
`pp` would add 39 failures without closing a reachable hole, so I left the CI scope alone and
am flagging it rather than silently changing the gate.

One caveat on my own probe, since it nearly gave a false answer: my first version rewrote
only parts that already existed in the fixture, so two bombs were never written and the
screen looked like it had missed them. The numbers above come from the corrected probe that
adds the part. Recorded because the same trap would hide a real bug.

## NEXT

Open, in order:

1. Supervisor review of the three commits. No external AI CLI was called this session, as
   instructed.
2. Decide whether to widen `bandit -r app ops` to include `pp`. On the evidence above it
   would fail on defended code; the honest options are to leave the scope as is, or to
   suppress per finding class with a written reason.
3. Run the security job's other gates, which I did not run: semgrep (p/python, p/javascript,
   p/owasp-top-ten, p/secrets), `pip-audit --require-hashes -r requirements.lock`, gitleaks
   over history including the three new commits.
4. Web suite, if `web/` is ever touched. No web file changed here.
5. `backend-security.md` still records the pending final repair scan and reviewer assessment.
   That file was not updated because it describes the `feat/v2-backend` track, and I did not
   want to rewrite another track's record; the reviewer may want it reconciled.

No push, merge, PR or deployment. `.venv` and `web/node_modules` remain untracked. No
dependency change: `defusedxml==0.7.1` was already pinned and vendored before this work.
No real or private workbook was read.

## Incident to report

While taking the ruff baseline I ran `git stash -q` in a worktree with no uncommitted
changes. It had nothing to stash, so the following `git stash pop -q` popped
`stash@{0}`, which belonged to another branch (`feat/v2-dashboard`, "dashboard wip docs+dist
before merge"), and conflicted on `.planning/STATE.md` and two `static/dist` files.

No work was lost. The three commits were already committed, and the conflicting files are
build artifacts and docs. `git reset --hard HEAD` restored this worktree and the stash
remains at `stash@{0}` with all seven of its files. I verified the stash content afterward.

Reported rather than hidden: the stash belongs to someone else's branch, so whoever owns it
should confirm they still want it. The lesson is to use a scratch worktree for baselines,
which is what I switched to and what produced the verified 0-introduced number above.

Most important thing you may be missing: CI scans `app` and `ops`, but the v2 application
lives in `pp/`. The gate that went red today was the smallest part of the codebase, so a
green `bandit` line is not evidence about the code people actually run.