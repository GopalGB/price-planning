"""Lines up an uploaded sheet's rows with the template's when rows were added, inserted or deleted.

Rows are compared by what they hold: where the formulas sit, and every value that was not edited (item
numbers, labels). Rows that hold nothing carry no evidence and are skipped. The result maps each row of
the upload onto a template row, or marks it as a row the template has no place for.
"""

from __future__ import annotations

from array import array
from bisect import bisect_right
from dataclasses import dataclass, field

FORMULA = ("formula",)  # a tuple never equals a cell value, so it cannot collide with data
MAX_SHIFT = 400  # rows added or removed on one sheet beyond which matching is not attempted
MIN_MATCHED = 0.5  # share of the template's filled rows that must match when rows moved
FULL_MATCH = 100.0
SAME_ROW = 0.5  # a weak preference for "nothing moved" when a row offers no evidence either way
SKIP = 1.0  # cost of leaving a row unpaired: identical rows (blank plan lines) keep one offset
MASK = (1 << 64) - 1


def formula_column_hash(col: int) -> int:
    """Stable across processes and restarts (Python's hash() is salted), summed per row."""
    return ((col + 1) * 0x9E3779B97F4A7C15) & MASK


@dataclass
class RowAlignment:
    mapping: dict = field(default_factory=dict)  # upload row -> template row, for rows the app applies
    inserted: list = field(default_factory=list)  # upload rows with data and no template row to go to
    deleted: list = field(default_factory=list)  # template rows with data that the upload no longer has
    moved: bool = False
    failure: str | None = None

    def place(self, row):
        """Template row for any upload row: matched rows as matched; rows with no content (blank or
        empty text) keep the offset of the nearest matched row above, unless that row is spoken for."""
        if row in self.mapping:
            return self.mapping[row]
        if row in self.inserted:
            return None
        if not hasattr(self, "_anchors"):
            self._anchors = sorted(self.mapping.items())
            self._taken = set(self.mapping.values())
        position = bisect_right(self._anchors, (row, float("inf")))
        above = self._anchors[position - 1] if position else (-1, -1)
        target = above[1] + (row - above[0])
        below = self._anchors[position] if position < len(self._anchors) else None
        if target < 0 or target in self._taken or (below is not None and target >= below[1]):
            return None
        return target


def align_rows(template: dict, upload: dict) -> RowAlignment:
    """template, upload: {row: {col: value or FORMULA}} holding filled rows only, 0-based."""
    t_rows, u_rows = sorted(template), sorted(upload)
    lo = min(0, len(u_rows) - len(t_rows)) - 3
    hi = max(0, len(u_rows) - len(t_rows)) + 3
    if hi - lo > 2 * MAX_SHIFT:
        return RowAlignment(failure=f"more than {MAX_SHIFT} rows were added or removed")
    pairs = _banded_pairs(t_rows, u_rows, template, upload, lo, hi)
    shift = _column_shift(pairs, template, upload)
    if shift:
        side = "inserted" if shift > 0 else "deleted"
        return RowAlignment(failure=f"columns look {side}: cells sit {abs(shift)} column(s) "
                                    f"{'right' if shift > 0 else 'left'} of the template's")
    result = _fill_gaps(pairs, t_rows, u_rows, template)
    filled = len(t_rows)
    evidence = sum(1 for u, t in result.mapping.items() if t in template and _weight(
        template[t], upload[u], t == u) >= 1)
    if result.moved and filled and evidence < MIN_MATCHED * filled:
        result.failure = f"only {evidence} of {filled} filled rows match the template"
    return result


def _weight(t_cells, u_cells, same_row):
    if t_cells == u_cells:
        return FULL_MATCH
    t_formulas = {col for col, value in t_cells.items() if value is FORMULA}
    u_formulas = {col for col, value in u_cells.items() if value is FORMULA}
    # Formulas in the same columns are evidence; a formula typed over or replaced by a value only
    # removes that evidence, it does not make the row a different row
    matches = sum(1 for col, value in u_cells.items()
                  if value is not FORMULA and t_cells.get(col, FORMULA) == value)
    return matches + (2 if t_formulas and t_formulas == u_formulas else 0) + (SAME_ROW if same_row else 0)


def _column_shift(pairs, template, upload, reach=3):
    """The offset a template column's contents moved by, when most of its rows agree it is not zero:
    a column was inserted or deleted before it."""
    votes, filled = {}, {}
    for u, t in pairs:
        t_cells, u_cells = template[t], upload[u]
        for col, value in t_cells.items():
            filled[col] = filled.get(col, 0) + 1
            for k in range(-reach, reach + 1):
                if col + k >= 0 and u_cells.get(col + k) == value:
                    votes.setdefault(col, {}).setdefault(k, 0)
                    votes[col][k] += 1
    for col in sorted(votes):
        counts = votes[col]
        k = max(counts, key=lambda key: (counts[key], -abs(key)))
        if k and filled[col] >= 2 and counts[k] >= filled[col] / 2 and counts[k] > counts.get(0, 0):
            return k
    return 0


def _banded_pairs(t_rows, u_rows, template, upload, lo, hi):
    """Highest-weight order-keeping pairing (an alignment, like a diff) within a band of row shifts."""
    n, m, width = len(t_rows), len(u_rows), hi - lo + 1
    if not n or not m:
        return []
    score = [array("d", [0.0]) * width for _ in range(n + 1)]
    step = [bytearray(width) for _ in range(n + 1)]  # 0 skip template row, 1 skip upload row, 2 pair

    def get(i, j):
        k = j - i - lo
        return score[i][k] if 0 <= k < width and 0 <= j <= m else None

    for i in range(n + 1):
        for k in range(width):
            j = i + lo + k
            if j < 0 or j > m or (i == 0 and j == 0):
                continue
            best, move = None, 0
            up = get(i - 1, j) if i else None
            if up is not None:
                best, move = up - SKIP, 0
            left = get(i, j - 1) if j else None
            if left is not None and (best is None or left - SKIP > best):
                best, move = left - SKIP, 1
            if i and j:
                diagonal = get(i - 1, j - 1)
                if diagonal is not None:
                    weight = _weight(template[t_rows[i - 1]], upload[u_rows[j - 1]],
                                     t_rows[i - 1] == u_rows[j - 1])
                    if best is None or diagonal + weight >= best:
                        best, move = diagonal + weight, 2
            score[i][k], step[i][k] = best if best is not None else float("-inf"), move
    i, j, pairs = n, m, []
    if not 0 <= m - n - lo < width:
        return []
    while i > 0 or j > 0:
        move = step[i][j - i - lo]
        if move == 2:
            pairs.append((u_rows[j - 1], t_rows[i - 1]))
            i, j = i - 1, j - 1
        elif move == 1:
            j -= 1
        else:
            i -= 1
    pairs.reverse()
    return pairs


def _fill_gaps(pairs, t_rows, u_rows, template):
    """Unpaired upload rows keep the offset of the pair above them while the template has an empty row
    there (a filled-in blank row, or rows added at the end); otherwise they are rows the template lacks."""
    result = RowAlignment(mapping=dict(pairs))
    taken = set(result.mapping.values())
    anchors = sorted(pairs)
    position = 0
    for u in u_rows:
        if u in result.mapping:
            continue
        while position < len(anchors) and anchors[position][0] < u:
            position += 1
        above = anchors[position - 1] if position else (-1, -1)
        below = anchors[position] if position < len(anchors) else None
        target = above[1] + (u - above[0])
        if (target >= 0 and (below is None or target < below[1]) and target not in taken
                and target not in template):
            result.mapping[u] = target
            taken.add(target)
        else:
            result.inserted.append(u)
    result.moved = bool(result.inserted) or any(u != t for u, t in result.mapping.items())
    if result.moved:
        result.deleted = [t for t in t_rows if t not in taken]
    return result
