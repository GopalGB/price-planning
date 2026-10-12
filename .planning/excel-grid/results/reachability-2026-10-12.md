# Cell-level reachability on the real template (counts only, 2026-10-12)

Script: static reference parse of every formula (cells, ranges, whole rows and
columns, defined names), inputs = unlocked cells on protected sheets plus
unlocked row/column defaults, propagation with 2D prefix sums. 153 s, 714 MB peak.

| Measure | Count |
|---|---|
| Formula cells | 1,894,479 |
| Single-cell links / range links / unique ranges | 8,932,081 / 4,004,612 / 46,768 |
| Defined-name uses | 940,700 |
| OFFSET/INDIRECT formulas / volatile | 2 / 4 |
| Input cells (explicit + defaults) | 241,097 |
| Formulas any input can reach | 1,400,337 (73.9%), same in the strict and cautious runs |
| Cone of one input edit, 20 samples | 6 reach nothing; median 215,630; max 563,754 |

Conclusion: pruning to input-reachable formulas removes only 26%, not enough
for the 8 GB Mac or the 16 GB office cap. One edit can reach up to 30% of the
workbook, largely through whole-column ranges and names.
