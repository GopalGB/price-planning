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
