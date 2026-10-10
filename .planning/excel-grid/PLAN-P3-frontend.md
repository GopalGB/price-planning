# Track B execution plan

Approved scope: user's 2026-10-10 Track B instructions, go-token approved, proceed.
Goal: synthetic-backed P3 frontend matching SPEC 3.5 and section 5, with API.md as the immutable contract.
Ownership: web/, static/dist/, web/test/, frontend track evidence and STATE handoff. No backend edits, no push, no real workbook access.

- [ ] Unit 1: Test-first operation state, atomic validation, rollback, version-safe undo/redo, row cache and number formats. Node unit tests green.
- [ ] Unit 2: Test-first in-memory synthetic API with sessions, permissions, categories, copies, ops, history and WebSocket calculation. HTTP/WS tests green.
- [ ] Unit 3: Auth, Copies home, category membership, Share/Fork dialogs and visual shell. UI contract checks green.
- [ ] Unit 4: Univer 1.0.4 adapter, input-only operation interception, lazy 256-row blocks with eviction, live patches, presence, history, selection summary and export link. Adapter/state tests and browser checks green.
- [ ] Unit 5: Deterministic build and BUILD.json, security/dependency checks, independent exact-model review, Ego Lite acceptance and screenshots, Python regression suite, content guard, track report and commits.

Validation: write a failing test before each logical unit, run it red, implement, run green. Required final commands are the user's exact Python suite, content guard and branch log. Browser uses one TaskSpace and closes with finish({keep: []}).
Out of scope: backend changes, deployment, real workbook parity, office LAN acceptance, paid Strix runs, pushes and PR publication.
Checkpoint: beginning Unit 1. Existing worktree branch feat/v2-frontend. Available disk below 5 GiB, no agent fan-out. SPEC D7 identifies gpt-6.1-sol as executor; the active executor matches this explicit project choice.
