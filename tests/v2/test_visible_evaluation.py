"""An edit must never be answered from Excel's cached value.

After an input changes, every cell downstream of it is stale until the engine
recalculates it. Only the cells inside the blocks a client currently has loaded
are recalculated eagerly; the rest are marked stale and are calculated when
their block is asked for.
"""

import tempfile
import unittest
from pathlib import Path
from unittest import mock

from ops.engine_scale_fixture import build_manifest, expected_formula_values, generate
from tests.v2.support import TMP

SEED = 20261011
SCALE = 0.0001


def _fixture(path):
    generate(path, SCALE, SEED)
    return path


def _rows_of(manifest, sheet=0):
    return sorted({row for index, row, _col, _value
                   in expected_formula_values(manifest, 7, 3) if index == sheet})


class VisibleEvaluationTests(unittest.TestCase):
    def test_an_edit_evaluates_only_the_affected_cells_a_client_has_open(self):
        from pp.engine.formualizer import FormualizerAdapter, VisibleBlocks

        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = _fixture(Path(temp) / "fixture.xlsx")
            manifest = _manifest(source)
            engine = FormualizerAdapter.load(source)
            first = min(_rows_of(manifest))
            engine.set_visible_blocks(VisibleBlocks.of((0, first, first + 1)))

            changes = engine.apply([(0, 0, 0, 9)])

            changed_rows = {row for _sheet, row, _col, _value in changes}
            self.assertTrue(changed_rows)
            self.assertTrue(changed_rows <= {first, first + 1})
            self.assertLess(len(changes), len(engine.formula_anchors))

    def test_a_stale_downstream_cell_is_never_answered_from_the_excel_cache(self):
        from pp.engine.formualizer import FormualizerAdapter, VisibleBlocks

        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = _fixture(Path(temp) / "fixture.xlsx")
            manifest = _manifest(source)
            engine = FormualizerAdapter.load(source)
            engine.set_visible_blocks(VisibleBlocks.of((0, 0, 0)))

            engine.apply([(0, 0, 0, 9)])

            stale = _stale_on_other_sheets(engine)
            self.assertTrue(stale)
            for sheet, row, col in stale:
                expected = _expected(manifest, sheet, row, col, broad=9, ordinary=3)
                self.assertEqual(engine.get(sheet, row, col), expected)

    def test_reading_a_stale_cell_evaluates_it_and_clears_the_mark(self):
        from pp.engine.formualizer import FormualizerAdapter, VisibleBlocks

        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = _fixture(Path(temp) / "fixture.xlsx")
            engine = FormualizerAdapter.load(source)
            engine.set_visible_blocks(VisibleBlocks.of((0, 0, 0)))
            engine.apply([(0, 0, 0, 9)])
            stale = _stale_on_other_sheets(engine)

            sheet, row, col = stale[0]
            engine.get(sheet, row, col)
            first_read = engine.get(sheet, row, col)

            self.assertEqual(first_read, engine.workbook.get_value(
                engine.names[sheet], row + 1, col + 1))
            self.assertNotIn((sheet, row, col), engine.stale_cells)

    def test_a_cell_outside_the_block_stays_stale_until_its_block_is_opened(self):
        from pp.engine.formualizer import FormualizerAdapter, VisibleBlocks

        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = _fixture(Path(temp) / "fixture.xlsx")
            manifest = _manifest(source)
            engine = FormualizerAdapter.load(source)
            rows = _rows_of(manifest)
            engine.set_visible_blocks(VisibleBlocks.of((0, rows[0], rows[0])))

            engine.apply([(0, 0, 0, 9)])
            hidden = _stale_on_other_sheets(engine)
            self.assertTrue(hidden)

            sheet, row, col = hidden[0]
            engine.set_visible_blocks(VisibleBlocks.of((0, rows[0], rows[0]),
                                                        (sheet, row, row)))
            engine.recalculate_stale()

            self.assertEqual(engine.get(sheet, row, col),
                             _expected(manifest, sheet, row, col, broad=9, ordinary=3))

    def test_opening_every_block_restores_the_full_affected_set(self):
        from pp.engine.formualizer import FormualizerAdapter, VisibleBlocks

        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = _fixture(Path(temp) / "fixture.xlsx")
            manifest = _manifest(source)
            engine = FormualizerAdapter.load(source)
            engine.set_visible_blocks(VisibleBlocks.of((0, 0, 0)))
            engine.apply([(0, 0, 0, 9)])

            stale = sorted(engine.stale_cells)
            engine.set_visible_blocks(VisibleBlocks.of_rows(engine.formula_cells))
            changes = engine.recalculate_stale()

            self.assertEqual(len(changes), len(stale))
            self.assertEqual(engine.stale_cells, frozenset())
            for sheet, row, col in stale:
                self.assertEqual(engine.get(sheet, row, col),
                                 _expected(manifest, sheet, row, col, broad=9, ordinary=3))

    def test_untouched_cells_keep_serving_the_stored_cache(self):
        from pp.engine.formualizer import FormualizerAdapter

        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = _fixture(Path(temp) / "fixture.xlsx")
            manifest = _manifest(source)
            engine = FormualizerAdapter.load(source)
            row, col = min((row, col) for index, row, col, _v
                           in expected_formula_values(manifest, 7, 3) if index == 0)

            with mock.patch("formualizer.Workbook.evaluate_cell",
                            side_effect=AssertionError("an untouched cell was evaluated")):
                value = engine.get(0, row, col)

            self.assertEqual(value, _expected(manifest, 0, row, col, broad=7, ordinary=3))
            self.assertEqual(engine.stale_cells, frozenset())

    def test_a_visible_only_broad_edit_publishes_fewer_cells_than_the_full_one(self):
        from pp.engine.formualizer import FormualizerAdapter, VisibleBlocks

        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = Path(temp) / "fixture.xlsx"
            generate(source, 0.005, SEED)
            manifest = build_manifest(0.005, SEED)
            engine = FormualizerAdapter.load(source)
            rows = sorted({row for index, row, _col, _value
                           in expected_formula_values(manifest, 7, 3) if index == 0})
            engine.set_visible_blocks(VisibleBlocks.of((0, rows[0], rows[0] + 39)))

            visible_changes = engine.apply([(0, 0, 0, 9)])
            stale = set(engine.stale_cells)
            engine.set_visible_blocks(VisibleBlocks.of_rows(engine.formula_cells))
            rest = engine.recalculate_stale()

            self.assertLess(len(visible_changes), len(engine.formula_anchors))
            self.assertEqual(len(visible_changes) + len(rest), len(engine.formula_anchors))
            self.assertEqual(stale & set(engine.stale_cells), set())

    def test_a_second_edit_keeps_the_union_of_unevaluated_cells_stale(self):
        from pp.engine.formualizer import FormualizerAdapter, VisibleBlocks

        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = _fixture(Path(temp) / "fixture.xlsx")
            engine = FormualizerAdapter.load(source)
            engine.set_visible_blocks(VisibleBlocks.of((0, 0, 0)))

            engine.apply([(0, 0, 0, 9)])
            after_first = set(engine.stale_cells)
            engine.apply([(0, 0, 1, 4)])

            self.assertTrue(after_first <= set(engine.stale_cells))


def _manifest(source):
    from ops.engine_scale_fixture import build_manifest

    return build_manifest(SCALE, SEED)


def _stale_on_other_sheets(engine):
    return sorted(engine.stale_cells)[:5]


def _expected(manifest, sheet, row, col, broad, ordinary):
    for index, expected_row, expected_col, value in expected_formula_values(
            manifest, broad, ordinary):
        if (index, expected_row, expected_col) == (sheet, row, col):
            return value
    raise AssertionError(f"cell {(sheet, row, col)} is not a formula cell")


def _tmp():
    TMP.mkdir(parents=True, exist_ok=True)
    return TMP


if __name__ == "__main__":
    unittest.main()