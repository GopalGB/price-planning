**Changes required.** Review dated 2026-10-10. These are specification defects and missing proof gates, not claims that the proposed libraries cannot work.

1. **HIGH - Strix containment is not established** ([L189](</Users/gopalmacbook/Desktop/Max HQ/price-planning/.planning/excel-grid/SPEC.md:189>)). Localhost and a throwaway directory do not enforce isolation. The [SDK documents platform-dependent restrictions](https://openai.github.io/openai-agents-python/sandbox/clients/), without network isolation.  
   **Fix:** “The custom backend remains unimplemented. Name its supported execution host and enforce filesystem/network isolation with filtered credentials. Until containment tests pass, report `STRIX_ISOLATION_BLOCKED`. Budget limits and approval are not containment.”

2. **HIGH - Engine commitment precedes evidence** ([L162](</Users/gopalmacbook/Desktop/Max HQ/price-planning/.planning/excel-grid/SPEC.md:162>)). The spike measured loading; parity, evaluation memory and edit latency remain unmeasured.  
   **Fix:** “Formualizer remains provisional. P1 blocks P2 until two Windows real-workbook replays pass baseline/post-edit parity, latency and memory targets. Define ordinary, broad and worst-fan-out edits in a fixed benchmark manifest. A fallback decision alone does not pass P1.”

3. **HIGH - Static sharing can freeze incorrect values** ([L77](</Users/gopalmacbook/Desktop/Max HQ/price-planning/.planning/excel-grid/SPEC.md:77>)). Static reference arguments can address mutable data.  
   **Fix:** “Share only formulas proven pure and transitively independent of all mutable inputs/sources. Unknown dependencies remain dynamic. Precompute reference targets, not mutable values. Version source overlays with copies/forks. Demonstrate engine sharing support or budget independent instances.”

4. **HIGH - Univer paging and permanent calculation suppression are unproved** ([L107](</Users/gopalmacbook/Desktop/Max HQ/price-planning/.planning/excel-grid/SPEC.md:107>)). [`NO_CALCULATION` concerns initialization](https://docs.univer.ai/guides/sheets/features/core/formula).  
   **Fix:** “Before P2, pin OSS packages and prove bounded row loading/eviction, zero client recalculation and server-value preservation through paste/undo. Test 4088×312 and 100000×17 sheets, including offscreen copy/find. Glide must pass identical merges, freezes, styles and editing gates before substitution.”

5. **HIGH - Validation can race with commit** ([L93](</Users/gopalmacbook/Desktop/Max HQ/price-planning/.planning/excel-grid/SPEC.md:93>)). [WAL serializes writers](https://www.sqlite.org/wal.html); it does not fix pre-transaction checks.  
   **Fix:** “Deduplicate authenticated operations before version validation. Recheck permissions and versions inside `BEGIN IMMEDIATE`; reject the entire stale batch. Enforce unique copy/operation IDs bound to actor/payload, return stored acknowledgements on identical retries, and configure bounded busy handling.”

6. **HIGH - Publication and reconnect semantics are incomplete** ([L98](</Users/gopalmacbook/Desktop/Max HQ/price-planning/.planning/excel-grid/SPEC.md:98>)). Revision tags alone do not prevent stale overwrite or missed broadcasts.  
   **Fix:** “Serialize each copy’s engine against an exact reconstructed revision. Publish complete, atomic result batches monotonically. Reconnect through ordered replay or a revision-consistent snapshot. Remote patches must not generate outgoing edits or local undo entries.”

7. **HIGH - Private-copy access is underspecified** ([L85](</Users/gopalmacbook/Desktop/Max HQ/price-planning/.planning/excel-grid/SPEC.md:85>)). Only writes explicitly require authorization.  
   **Fix:** “Define category-to-copy role inheritance and named-share eligibility. Authorize listings, snapshots, history, exports, forks and WebSocket subscriptions; revoke live access on membership changes. Fork atomically at one revision, retain inherited provenance, and record subsequent editors separately.”

8. **HIGH - Windows/LAN operability is assumed** ([L152](</Users/gopalmacbook/Desktop/Max HQ/price-planning/.planning/excel-grid/SPEC.md:152>)). [Uvicorn requires a compatible loop and WebSocket implementation](https://www.uvicorn.org/settings/).  
   **Fix:** “Pin supported CPython/Windows versions and complete hashed wheel dependencies, including a WebSocket backend; use asyncio. P0 tests offline/no-admin installation and remote-LAN connectivity. Name firewall/certificate prerequisites. Production uses HTTPS/WSS and Secure cookies. Keep calculation off the ASGI loop and verify worker isolation.”

9. **HIGH - Committed results can leak workbook data** ([L146](</Users/gopalmacbook/Desktop/Max HQ/price-planning/.planning/excel-grid/SPEC.md:146>)). Ignore rules and local hooks do not constrain every cloud workflow.  
   **Fix:** “CI enforces artifact/content allowlists. Committable benchmark reports contain approved aggregate metrics, versions and hashes only; mismatched values, formulas, business labels and raw logs remain local. Cloud workers receive synthetic data only.”

10. **MEDIUM - Backup acceptance lacks WAL-safe semantics** ([L165](</Users/gopalmacbook/Desktop/Max HQ/price-planning/.planning/excel-grid/SPEC.md:165>)).  
    **Fix:** “Use SQLite’s online backup API, never copy only the live database file. Include template/schema hashes. Restore into an empty directory and verify integrity, foreign keys, revisions and history while testing concurrent writes.”

11. **MEDIUM - Parity and synthetic CI lack an independent oracle** ([L148](</Users/gopalmacbook/Desktop/Max HQ/price-planning/.planning/excel-grid/SPEC.md:148>)). The spike comparator conflates blanks/empty strings and errors/text.  
    **Fix:** “Use typed Excel-generated baseline/post-edit goldens distinguishing blanks, empty strings, zero, text, errors and spill geometry. Generate deterministic semantic and scale fixtures. Require real-workbook acceptance after engine/importer changes.”

12. **MEDIUM - CI-created bundles leave commit provenance ambiguous** ([L151](</Users/gopalmacbook/Desktop/Max HQ/price-planning/.planning/excel-grid/SPEC.md:151>)).  
    **Fix:** “Commit generated bundles on the PR branch. CI deterministically rebuilds and compares without pushing; every generated commit reruns checks. Record source/bundle hashes and dependency licenses. Explicitly allowlist verified vendored binaries against the size rule.”

13. **MEDIUM - Repository handoff is ambiguous** ([L142](</Users/gopalmacbook/Desktop/Max HQ/price-planning/.planning/excel-grid/SPEC.md:142>)). The current Git root is Max HQ.  
    **Fix:** “P0 records the standalone checkout root, origin and migrated files. Mac/office/cloud handoffs record branch, commit, tests and next task in STATE.md. Explicitly load AGENTS.md in every harness. Deploy tested commits with backup, migration and rollback procedures.”

14. **LOW - Compatibility phase conflicts** ([L75](</Users/gopalmacbook/Desktop/Max HQ/price-planning/.planning/excel-grid/SPEC.md:75>)).  
    **Fix:** “Compatibility shims and semantic parity are required before P1 Engine exits.”

```text
STATUS: DONE_WITH_CONCERNS
CHANGED: none
TESTS: document/source cross-check; runtime tests not run
GATE: approved read-only review
LESSON: none
NEXT: amend SPEC
DONE: ranked defects and exact fixes
VERIFIED: five documents, primary sources, unchanged review files
LEFT: implementation feasibility gates
```