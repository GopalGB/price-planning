"""Row matching for uploads whose rows moved: it either matches rows exactly or refuses."""

from __future__ import annotations

import random
import unittest

from pp.upload_align import FORMULA, align_rows


def _table(items, prices):
    return {row: {0: item, 1: price, 2: FORMULA} for row, (item, price) in enumerate(zip(items, prices))}


class AlignRowsTests(unittest.TestCase):
    def test_inserted_row_is_named_and_rows_below_shift_back(self):
        template = _table(["S-0", "S-1", "S-2", "S-3"], [10, 20, 30, 40])
        upload = _table(["S-0", "NEW", "S-1", "S-2", "S-3"], [11, 99, 21, 31, 41])

        result = align_rows(template, upload)

        self.assertIsNone(result.failure)
        self.assertEqual(result.mapping, {0: 0, 2: 1, 3: 2, 4: 3})
        self.assertEqual((result.inserted, result.deleted), ([1], []))

    def test_rows_added_at_the_end_keep_their_place(self):
        template = _table(["S-0", "S-1"], [10, 20])
        upload = _table(["S-0", "S-1", "S-2"], [10, 25, 30])

        result = align_rows(template, upload)

        self.assertEqual((result.failure, result.moved, result.mapping), (None, False, {0: 0, 1: 1, 2: 2}))

    def test_last_row_deleted_is_reported_not_cleared(self):
        template = _table(["S-0", "S-1", "S-2"], [10, 20, 30])
        upload = _table(["S-0", "S-1"], [10, 20])

        result = align_rows(template, upload, upload_end=2)

        self.assertEqual((result.failure, result.deleted), (None, [2]))

    def test_row_emptied_in_place_is_a_clear_not_a_delete(self):
        template = _table(["S-0", "S-1", "S-2"], [10, 20, 30])
        upload = {0: template[0], 2: template[2]}

        result = align_rows(template, upload, upload_end=3)

        self.assertEqual((result.failure, result.moved, result.deleted), (None, False, []))

    def test_rows_with_nothing_unchanged_are_refused_when_rows_moved(self):
        # No item column and every price edited: a deleted row cannot be located
        template = {row: {1: 10 * row, 2: FORMULA} for row in range(8)}
        upload = {row: {1: 10 * row + 1, 2: FORMULA} for row in range(7)}

        result = align_rows(template, upload, upload_end=7)

        self.assertIsNotNone(result.failure)
        self.assertEqual(result.mapping, {})

    def test_identical_rows_around_a_deletion_are_refused(self):
        pad = {0: "Pad", 1: 5, 2: FORMULA}
        template = {0: {0: "Hdr"}, 1: pad, 2: pad, 3: pad, 4: pad, 5: {0: "End", 2: FORMULA}}
        upload = {0: {0: "Hdr"}, 1: pad, 2: {0: "Pad", 1: 9, 2: FORMULA}, 3: pad, 4: {0: "End", 2: FORMULA}}

        result = align_rows(template, upload)

        self.assertIn("identical", result.failure or "")

    def test_repeated_default_values_do_not_look_like_a_moved_column(self):
        template = {row: {0: f"S-{row}", 1: 0, 2: 0, 3: FORMULA} for row in range(6)}
        upload = {row: {0: f"S-{row}", 1: 5, 2: 0, 3: FORMULA} for row in range(7)}

        self.assertIsNone(align_rows(template, upload).failure)

    def test_column_inserted_before_edited_inputs_is_refused(self):
        template = {row: {0: f"S-{row}", 1: 10 * row} for row in range(6)}
        upload = {row: {0: f"S-{row}", 2: 10 * row + 1} for row in range(6)}

        self.assertIn("columns look inserted", align_rows(template, upload).failure or "")

    def test_random_edits_never_move_a_row_to_the_wrong_place(self):
        # When the matcher says rows moved, every matched row must be the right one
        wrong = 0
        for seed in range(400):
            rng = random.Random(seed)
            n = rng.choice([5, 10, 30])
            template = {row: {0: f"S-{row}", 1: 10 + row, 2: FORMULA} for row in range(n)}
            deleted = set(rng.sample(range(n), rng.randint(0, min(3, n - 1))))
            sequence: list = [row for row in range(n) if row not in deleted]
            for _ in range(rng.randint(0, 4)):
                sequence.insert(rng.randint(0, len(sequence)), None)
            upload, truth = {}, {}
            for position, row in enumerate(sequence):
                if row is None:
                    upload[position] = {0: f"NEW{position}", 1: rng.randint(1000, 9999), 2: FORMULA}
                else:
                    upload[position] = {**template[row], **({1: rng.randint(1000, 9999)}
                                                            if rng.random() < 0.5 else {})}
                    truth[position] = row
            result = align_rows(template, upload)
            if result.failure is None and result.moved:
                wrong += sum(1 for u, t in result.mapping.items() if truth.get(u, t) != t)
        self.assertEqual(wrong, 0)


if __name__ == "__main__":
    unittest.main()
