# Lessons

## FRONTEND-TRACK-SCOPE
hits: 1
signal: Explicit frontend track boundaries and verification requirements reiterated.
correct: Work only in this worktree. Never inspect private workbook paths or external Office files. Own web/, static/dist/, frontend tests and track handoff only. Preserve API.md, use synthetic fixtures, test first, commit without trailers, never push, and report unverified gates explicitly.

## INTERFACE-ADDRESS-BOUNDARY
hits: 1
signal: A prior interface parser treated broadcast and gateway tokens as local server addresses.
correct: Derive numeric LAN origins only from anchored local interface address records or a UDP socket's confirmed local source when enumeration yields no usable address. DNS answers are route hints, never trusted numeric origins. Test Darwin, Linux, and Windows records containing misleading broadcast, gateway, and DNS addresses.

## SPILL-UNDO-INPUT-LINEAGE
hits: 1
signal: A compensating undo used the current computed spill display as redo input and restricted backend restoration to sparse input-mask membership.
correct: Keep original input values as the undo/redo lineage independently of computed display state, and authorize restoration with immutable snapshot.is_input semantics while retaining locked behavior for formula-only spill followers. Test the full clear, spill, undo, redo cycle plus explicit and implicit original inputs before claiming the defect fixed.

## BROWSER-FINAL-GATE
hits: 1
signal: The browser TaskSpace was finished before review corrections were complete, deleting the space needed for the final source check.
correct: Retain the one TaskSpace until the final code and independent review gates are settled, then verify that revision and finish once. If the finished space no longer exists, name the browser skill's replacement-space approval requirement; preserve old proof as historical evidence and never claim it verifies later code.

## NATIVE-LOAD-KEEP-RESIDENT
hits: 1
signal: Owner redirected the adapter bottleneck work to measured native Formualizer loading.
correct: Load the template natively once per active copy, retain snapshot caches for untouched display, apply only input overlays, and evaluate affected visible cells with bounded transitive dependents. Keep memory budget and idle LRU with connected-editor protection. Never rebuild formulas per cell at load or reuse a single mutable engine across copies. Use synthetic fixtures here; controller owns real-workbook parity.

## SHARED-QUALIFIER-IS-A-CELL
hits: 1
signal: Native load evaluated 47 of 150 formula cells to #REF!, all shared-formula followers.
correct: Formualizer 0.11.1 expands a shared follower by translating the master with the follower offset, and that translation shifts a cross-sheet qualifier whose sheet name parses as a cell address, so S01!$A$1 becomes T1!$A$1. Excel never translates the sheet part. Reproduce minimally, do not patch the pinned library, and expand shared followers in the package before load. Check any streaming XML rewrite for cells that straddle a read boundary: they are copied through unexamined unless the flush stops at the first cell that cannot be completed in the window.

## CACHE-MUST-NOT-ANSWER-AN-EDIT
hits: 1
signal: Visible-only evaluation published nothing, because resolving a stale cell during the publish capture made every cell look unchanged.
correct: Clear the stale marks for the cells about to be evaluated before the capture pass, and restore them if evaluation raises. A cell downstream of an edit must never be answered from the value Excel cached, because that cache still describes the pre-edit state; evaluate it when its block opens or when it is read.
