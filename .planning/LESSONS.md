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
