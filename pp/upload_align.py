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
MAX_CELLS = 4_000_000  # alignment table size (rows x band) beyond which matching is not attempted
FULL_MATCH = 100.0
# Weights: one unchanged value is worth more than the two skipped rows an insert or delete costs, so
# rows pair by content; a row with no unchanged value is worth less, so it never pairs across an
# insert just to save the skips; identical rows (blank plan lines) keep one offset since moving it
# gains nothing and costs skips
SAME_ROW = 0.25  # a weak preference for "nothing moved" when a row offers no evidence either way
SKIP = 0.4
FORMULAS_ALIGN = 0.5  # same formula columns, counted only next to an unchanged value
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
    unsure: bool = False  # rows were only added, and no unchanged value proves they went at the end

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


def align_rows(template: dict, upload: dict, upload_end: int | None = None) -> RowAlignment:
    """template, upload: {row: {col: value or FORMULA}} holding filled rows only, 0-based. upload_end:
    one past the file's last row holding any cell, empty or not; a template row whose place falls at or
    past it was deleted, one whose place holds an empty row was cleared.

    A mapping is only trusted when the rows prove it: every run of rows that keeps one offset must
    hold a value that was not edited and that its neighbours do not share (an item number, a label),
    and the matching must not change when ties are broken the other way (identical rows). Anything
    less is reported as a failure, never guessed."""
    t_rows, u_rows = sorted(template), sorted(upload)
    lo = min(0, len(u_rows) - len(t_rows)) - 3
    hi = max(0, len(u_rows) - len(t_rows)) + 3
    if hi - lo > 2 * MAX_SHIFT:
        return RowAlignment(failure=f"more than {MAX_SHIFT} rows were added or removed")
    if (len(t_rows) + 1) * (hi - lo + 1) > MAX_CELLS:
        return RowAlignment(failure="the sheet is too large to match rows that moved")
    pairs = _banded_pairs(t_rows, u_rows, template, upload, lo, hi)
    shift = _column_shift(pairs, template, upload)
    if shift:
        side = "inserted" if shift > 0 else "deleted"
        return RowAlignment(failure=f"columns look {side}: cells sit {abs(shift)} column(s) "
                                    f"{'right' if shift > 0 else 'left'} of the template's")
    end = max(upload, default=-1) + 1 if upload_end is None else upload_end
    result = _fill_gaps(pairs, t_rows, u_rows, template, end)
    if not result.moved:
        result.unsure = len(result.mapping) > len(pairs) and not any(
            _anchored(t, u, template, upload) for u, t in pairs)
        return result
    other = _banded_pairs(t_rows, u_rows, template, upload, lo, hi, prefer_skip=True)
    if other != pairs:
        rows = sorted({u for u, _ in set(pairs) ^ set(other)})
        return RowAlignment(failure=f"rows {rows[0] + 1} to {rows[-1] + 1} are identical, so it cannot "
                                    "tell which of them were added or removed")
    loose = _loose_pair(pairs, template, upload)
    if loose is not None:
        return RowAlignment(failure=f"row {loose[0] + 1} of your file shares no unchanged value with row "
                                    f"{loose[1] + 1} of the template, so the rows cannot be matched safely")
    unproven = _unproven_run(pairs, template, upload)
    if unproven is not None:
        return RowAlignment(failure=f"rows {unproven[0] + 1} to {unproven[1] + 1} have no unchanged "
                                    "value (an item number or label) to match them by")
    return result


def _loose_pair(pairs, template, upload):
    """A pair with no unchanged value of its own is kept only when the nearest proven pairs above and
    below it keep its offset: boxed in like that it cannot have slid. Returns the first that is not."""
    supported = [_supported(template[t], upload[u]) for u, t in pairs]
    offsets = [u - t for u, t in pairs]
    following_proven, following = [None] * len(pairs), None
    for index in range(len(pairs) - 1, -1, -1):
        following_proven[index] = following
        if supported[index]:
            following = index
    last = None
    for index, (pair, ok) in enumerate(zip(pairs, supported)):
        if ok:
            last = index
            continue
        following = following_proven[index]
        if last is None and following is None:
            return pair
        if any(k is not None and offsets[k] != offsets[index] for k in (last, following)):
            return pair
    return None


def _supported(t_cells, u_cells):
    """The pair is the same row by content: identical, or one non-formula value unchanged."""
    return t_cells == u_cells or any(value is not FORMULA and u_cells.get(col) == value
                                     for col, value in t_cells.items())


def _anchored(t, u, template, upload):
    """An unedited value that its neighbouring template rows do not share pins the row in place."""
    t_cells, u_cells = template[t], upload[u]
    above, below = template.get(t - 1, {}), template.get(t + 1, {})
    return any(value is not FORMULA and u_cells.get(col) == value
               and above.get(col) != value and below.get(col) != value
               for col, value in t_cells.items())


def _unproven_run(pairs, template, upload):
    """The first run of pairs sharing one offset that no anchored pair proves, as upload rows."""
    start, offset, proven = None, None, False
    for u, t in pairs + [(None, None)]:
        if u is None or u - t != offset:
            if start is not None and not proven:
                return start, previous
            if u is None:
                return None
            start, offset, proven = u, u - t, False
        proven = proven or _anchored(t, u, template, upload)
        previous = u
    return None


def _weight(t_cells, u_cells, same_row):
    if t_cells == u_cells:
        return FULL_MATCH
    t_formulas = {col for col, value in t_cells.items() if value is FORMULA}
    u_formulas = {col for col, value in u_cells.items() if value is FORMULA}
    # Formulas in the same columns are weak evidence (worth no more than a skipped row); a formula typed
    # over or replaced by a value only removes that evidence, it does not make the row a different row
    matches = sum(1 for col, value in u_cells.items()
                  if value is not FORMULA and t_cells.get(col, FORMULA) == value)
    aligned = FORMULAS_ALIGN if matches and t_formulas and t_formulas == u_formulas else 0
    return matches + aligned + (SAME_ROW if same_row else 0)


def _column_shift(pairs, template, upload, reach=3):
    """The offset a template column moved by, when most of its rows agree it is not zero: a column was
    inserted or deleted before it. Which columns are filled survives edits; values that did not change
    confirm it. A value its own column also holds at the shifted place is no evidence either way."""
    occupied, valued, filled = {}, {}, {}
    for u, t in pairs:
        t_cells, u_cells = template[t], upload[u]
        for col, value in t_cells.items():
            filled[col] = filled.get(col, 0) + 1
            vacated = col not in u_cells
            for k in range(-reach, reach + 1):
                if col + k < 0 or col + k not in u_cells:
                    continue
                # A cell the template fills is empty and one it leaves empty is filled: edits keep that
                if k == 0 or (vacated and col + k not in t_cells):
                    occupied.setdefault(col, {}).setdefault(k, 0)
                    occupied[col][k] += 1
                if u_cells[col + k] == value and (k == 0 or t_cells.get(col + k) != value):
                    valued.setdefault(col, {}).setdefault(k, 0)
                    valued[col][k] += 1
    for votes in (occupied, valued):
        for col in sorted(votes):
            counts = votes[col]
            k = max(counts, key=lambda key: (counts[key], -abs(key)))
            if k and filled[col] >= 2 and counts[k] >= filled[col] / 2 and counts[k] > counts.get(0, 0):
                return k
    return 0


def _banded_pairs(t_rows, u_rows, template, upload, lo, hi, prefer_skip=False):
    """Highest-weight order-keeping pairing (an alignment, like a diff) within a band of row shifts.
    prefer_skip breaks ties the other way, so a second run shows whether the answer was a coin toss."""
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
            if left is not None and (best is None or left - SKIP > best
                                     or (prefer_skip and left - SKIP == best)):
                best, move = left - SKIP, 1
            if i and j:
                diagonal = get(i - 1, j - 1)
                if diagonal is not None:
                    weight = _weight(template[t_rows[i - 1]], upload[u_rows[j - 1]],
                                     t_rows[i - 1] == u_rows[j - 1])
                    if best is None or diagonal + weight > best or (
                            not prefer_skip and diagonal + weight == best):
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


def _gone(u, filled_upload, end):
    """No upload row is left at this place: none fits, it holds another row's content, or it lies past
    the end of the file."""
    return u == -1 or u in filled_upload or u >= end


def _upload_row_at(t, anchors, anchor_rows):
    """The upload row that keeps the offset of the nearest pair above template row t, or -1 when the
    offsets leave no room for one (the row really went). Pairs keep order, so they sort by both rows."""
    position = bisect_right(anchor_rows, t)
    above = anchors[position - 1] if position else (-1, -1)
    below = anchors[position] if position < len(anchors) else None
    u = above[0] + (t - above[1])
    return u if u >= 0 and (below is None or u < below[0]) else -1


def _fill_gaps(pairs, t_rows, u_rows, template, end):
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
    filled_upload = set(u_rows)
    # A template row is deleted only when no upload row is left at its place: an upload row that sits
    # there empty means the row was cleared, which is an ordinary edit
    anchor_rows = [t for _, t in anchors]
    result.deleted = [t for t in t_rows if t not in taken
                      and _gone(_upload_row_at(t, anchors, anchor_rows), filled_upload, end)]
    result.moved = (bool(result.inserted) or bool(result.deleted)
                    or any(u != t for u, t in result.mapping.items()))
    return result
