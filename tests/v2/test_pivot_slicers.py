"""Excel pivots and slicers (P6.4): pivot output is Excel's, slicer selections replay on a fresh engine."""

import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

from openpyxl import Workbook

from pp.engine.excel import REFRESH_ALL
from pp.engine.worker import WorkerBatch
from tests.v2.support import TMP
from tests.v2.test_engine_cached import _replace_formula_values
from tests.v2.test_excel_macros import ComHostTests, FakeExcelRange, FakeExcelSheet, MacroApp

PIVOT = (0, 2, 0, 3, 0)  # Calc!A3:A4, the pivot output


def pivot_workbook(directory):
    """Calc!A1 input 2; B1 = A1*2 (cached 4); A3:A4 pivot output (10, 20); C1 = A3+A4 (cached 30)."""
    path = Path(directory) / "pivot.xlsx"
    book = Workbook()
    calc = book.active
    calc.title = "Calc"
    calc["A1"] = 2
    calc["B1"] = "=A1*2"
    calc["A3"] = 10
    calc["A4"] = 20
    calc["C1"] = "=A3+A4"
    book.save(path)
    _replace_formula_values(path, {"xl/worksheets/sheet1.xml": {"B1": ("4", None), "C1": ("30", None)}})
    return path


def north_hook(host, cache, selected):
    """Excel's pivot shows 100 when the north item is selected, 1 otherwise."""
    host.write(0, [(2, 0, 100 if "north" in selected else 1)])


def refresh_hook(host):
    host.write(0, [(2, 0, 11), (3, 0, 21)])
    return []


class PivotEngineTests(unittest.TestCase):
    def load(self, hooks=True):
        from pp.engine.excel import ExcelEngine, FormualizerHost

        TMP.mkdir(parents=True, exist_ok=True)
        directory = tempfile.TemporaryDirectory(dir=TMP)
        self.addCleanup(directory.cleanup)
        source = pivot_workbook(directory.name)
        hosts = []

        def factory(path, names):
            host = FormualizerHost(path, names)
            host.pivots = [PIVOT]
            if hooks:
                host.macros = {REFRESH_ALL: refresh_hook}
                host.refresh_hook = refresh_hook
                host.slicer_hook = north_hook
            hosts.append(host)
            return host

        engine = ExcelEngine.load(source, cache_dir=Path(directory.name) / "cache", host_factory=factory)
        self.addCleanup(engine.close)
        return engine, hosts

    def test_pivot_cells_are_published_computed_cells_not_overrides(self):
        engine, _ = self.load()

        self.assertEqual(engine.pivot_cells, {(0, 2, 0), (0, 3, 0)})
        self.assertIn((0, 2, 0), engine.formula_anchors)
        self.assertEqual(engine.get(0, 2, 0), 10)
        self.assertEqual(engine.get(0, 0, 2), 30)

    def test_refresh_publishes_pivot_changes_and_dependents_but_no_contents(self):
        engine, _ = self.load()

        result = engine.refresh_data()

        self.assertEqual(result["contents"], [])
        self.assertEqual(result["changes"], [(0, 0, 2, 32), (0, 2, 0, 11), (0, 3, 0, 21)])
        self.assertEqual(engine.overlay_snapshot(), [(0, 0, 2, 32), (0, 2, 0, 11), (0, 3, 0, 21)])

    def test_replay_and_edits_never_write_inside_a_pivot_range(self):
        engine, hosts = self.load()
        engine.refresh_data()
        with mock.patch.object(hosts[0], "write", wraps=hosts[0].write) as write:
            changes = engine.apply([(0, 2, 0, 99), (0, 0, 0, 5)])

        written = [cell for call in write.call_args_list for cell in call.args[1]]
        self.assertEqual([cell[:2] for cell in written], [(0, 0)])
        self.assertEqual(changes, [(0, 0, 1, 10)])  # C1 reads the pivot cells, which did not move
        self.assertEqual(engine.get(0, 2, 0), 11)

    def test_set_slicer_records_pivot_output_as_changes(self):
        engine, _ = self.load()

        result = engine.set_slicer("Slicer_Region", ["north"])

        self.assertEqual((result["messages"], result["contents"]), ([], []))
        self.assertEqual(result["changes"], [(0, 0, 2, 120), (0, 2, 0, 100)])
        self.assertEqual(engine.get(0, 2, 0), 100)

    def test_stored_selection_replays_on_a_fresh_engine_to_the_same_pivot_cells(self):
        original, _ = self.load()
        original.set_slicer("Slicer_Region", ["north"])
        replay, _ = self.load()

        result = replay.restore_pivots({"slicers": {"Slicer_Region": ["north"]}, "refreshed": False})

        self.assertEqual(result["changes"], [(0, 0, 2, 120), (0, 2, 0, 100)])
        cells = [(0, 2, 0), (0, 3, 0), (0, 0, 2)]
        self.assertEqual([original.get(*cell) for cell in cells], [replay.get(*cell) for cell in cells])

    def test_refreshed_state_refreshes_the_pivots_before_slicers_apply(self):
        engine, _ = self.load()

        engine.restore_pivots({"slicers": {}, "refreshed": True})

        self.assertEqual(engine.get(0, 2, 0), 11)
        self.assertEqual(engine.get(0, 0, 2), 32)


class PivotWorkerTests(unittest.TestCase):
    def copy(self):
        from pp.engine.worker import WarmCopy

        TMP.mkdir(parents=True, exist_ok=True)
        directory = tempfile.TemporaryDirectory(dir=TMP)
        self.addCleanup(directory.cleanup)
        source = pivot_workbook(directory.name)

        def factory(path, names):
            from pp.engine.excel import FormualizerHost

            host = FormualizerHost(path, names)
            host.pivots, host.slicer_hook, host.refresh_hook = [PIVOT], north_hook, refresh_hook
            host.macros = {REFRESH_ALL: refresh_hook}
            return host

        patch = mock.patch("pp.engine.excel.default_host", return_value=factory)
        patch.start()
        self.addCleanup(patch.stop)
        copy = WarmCopy("copy", "excel")
        self.addCleanup(copy.close)
        return copy, str(source)

    def test_fresh_load_applies_stored_selection_and_warm_copy_ignores_it(self):
        copy, path = self.copy()
        state = {"slicers": {"Slicer_Region": ["north"]}, "refreshed": False}

        first = copy.apply(1, path, WorkerBatch([], None, state))

        self.assertIn((0, 2, 0, 100), first["changes"])
        second = copy.apply(2, path, WorkerBatch([], 1, state))
        self.assertNotIn((0, 2, 0, 100), second["changes"])

    def test_slicer_request_selects_and_folds_nothing_into_overrides(self):
        copy, path = self.copy()

        result = copy.run_slicer(1, path, [], None, "Slicer_Region", ["north"])

        self.assertEqual((result["contents"], result["messages"]), ([], []))
        self.assertIn((0, 2, 0, 100), result["changes"])
        self.assertEqual(copy.overrides, {})

    def test_wire_tuples_carry_the_pivot_state_and_keep_the_plain_forms(self):
        from pp.engine.worker import _serve

        copy, path = self.copy()
        state = {"slicers": {"Slicer_Region": ["south"]}, "refreshed": False}
        slicer = _serve(copy, ("slicer", 1, path, [], None, "Slicer_Region", ["north"], state))
        plain = _serve(copy, (2, path, [], 1))

        self.assertIn((0, 2, 0, 100), slicer["changes"])
        self.assertEqual(plain["revision"], 2)


class ComPivotTests(unittest.TestCase):
    com_host = ComHostTests.com_host  # the same COM fakes and cleanup as the macro host tests

    def com_pivot_host(self, pivots, slicers):
        from tests.v2.test_excel_macros import FakeGui

        gui, killed = FakeGui(), []
        app = PivotApp(gui, slicers)
        sheet = FakePivotSheet(pivots)
        app.sheets["Calc"] = sheet
        host = self.com_host(app, gui, killed)
        return host, app, sheet

    def test_pivot_ranges_come_from_every_pivot_table_range(self):
        host, _, _ = self.com_pivot_host([FakePivot(top=3, left=1, rows=2, cols=1)], {})

        self.assertEqual(host.pivot_ranges(), [(0, 2, 0, 3, 0)])

    def test_refresh_pivots_refreshes_each_table_then_waits_for_queries(self):
        host, app, sheet = self.com_pivot_host([FakePivot(top=3, left=1, rows=2, cols=1)], {})

        host.refresh_pivots(5)

        self.assertEqual(sheet.pivots[0].refreshes, 1)
        self.assertIn("CalculateUntilAsyncQueriesDone", app.calls)

    def test_set_slicer_selects_exactly_the_named_items_with_pivots_held(self):
        cache = FakeSlicerCache(["north", "south", "(blank)"], [FakePivot(3, 1, 2, 1)])
        host, app, sheet = self.com_pivot_host([FakePivot(3, 1, 2, 1)], {"Slicer_Region": cache})

        host.set_slicer("Slicer_Region", ["south", "(blank)"])

        self.assertEqual([item.Selected for item in cache.SlicerItems.members], [False, True, True])
        self.assertEqual([pivot.ManualUpdate for pivot in cache.PivotTables.members], [False])
        self.assertEqual(cache.PivotTables.members[0].manual_history, [True, False])

    def test_set_slicer_with_an_item_excel_does_not_have_fails_and_restores_manual_update(self):
        from pp.engine.formualizer import EngineError

        cache = FakeSlicerCache(["north"], [FakePivot(3, 1, 2, 1)])
        host, _, _ = self.com_pivot_host([FakePivot(3, 1, 2, 1)], {"Slicer_Region": cache})

        with self.assertRaises(EngineError):
            host.set_slicer("Slicer_Region", ["north", "gone"])
        self.assertEqual(cache.PivotTables.members[0].ManualUpdate, False)


class FakePivot:
    def __init__(self, top, left, rows, cols):
        self.area = FakeExcelRange(None, (top, left), (top + rows - 1, left + cols - 1))
        self.refreshes, self.manual_history = 0, []
        self._manual = False

    @property
    def TableRange2(self):
        return self.area

    @property
    def ManualUpdate(self):
        return self._manual

    @ManualUpdate.setter
    def ManualUpdate(self, value):
        self._manual = value
        self.manual_history.append(value)

    def RefreshTable(self):
        self.refreshes += 1


class FakeCollection:
    def __init__(self, members):
        self.members = list(members)
        self.Count = len(self.members)

    def Item(self, position):
        return self.members[position - 1]


class FakeSlicerItem:
    def __init__(self, name):
        self.Name, self.Selected = name, False


class FakeSlicerCache:
    def __init__(self, items, pivots):
        self.PivotTables = FakeCollection(pivots)
        self.SlicerItems = FakeCollection([FakeSlicerItem(name) for name in items])


class FakePivotSheet(FakeExcelSheet):
    def __init__(self, pivots):
        super().__init__()
        self.pivots = pivots

    def PivotTables(self):
        return FakeCollection(self.pivots)


class PivotApp(MacroApp):
    def __init__(self, gui, slicers):
        super().__init__(gui, lambda app, name: None)
        self.slicers = slicers

    def open(self, path, **options):
        book = super().open(path, **options)
        book.SlicerCaches = lambda name: self.slicers[name]
        return book


class SlicerSmokeTests(unittest.TestCase):
    def test_smoke_selects_and_restores_each_slicer_then_replays_the_pivot_cells(self):
        from ops.excel_smoke import check_slicers
        from pp.engine.excel import ExcelEngine, FormualizerHost

        TMP.mkdir(parents=True, exist_ok=True)
        directory = tempfile.TemporaryDirectory(dir=TMP)
        self.addCleanup(directory.cleanup)
        source = pivot_workbook(directory.name)

        def open_engine(path):
            def factory(workbook, names):
                host = FormualizerHost(workbook, names)
                host.pivots = [PIVOT]
                host.macros = {REFRESH_ALL: refresh_hook}
                host.refresh_hook, host.slicer_hook = refresh_hook, north_hook
                return host

            return ExcelEngine.load(path, cache_dir=Path(directory.name) / "cache", host_factory=factory)

        snapshot = SimpleNamespace(source_path=source, sheets=[{"slicers": [{
            "engine": "excel", "cache": "Slicer_Region",
            "items": [{"caption": "north", "no_data": False, "selected": True},
                      {"caption": "south", "no_data": False, "selected": False}],
        }]}])
        lines = []

        failures = check_slicers(snapshot, open_engine, report=lines.append)

        self.assertEqual(failures, 0)
        self.assertEqual(lines[-1], "PASS")
        self.assertFalse(any("north" in line or "south" in line for line in lines))
