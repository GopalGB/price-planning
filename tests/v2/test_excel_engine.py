"""Excel engine: Excel calculates, the engine publishes exactly what changed."""

import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from tests.v2.support import TMP
from tests.v2.test_engine_cached import _replace_formula_values


def workbook(directory):
    path = Path(directory) / "plan.xlsx"
    book = Workbook()
    calc = book.active
    calc.title = "Calc"
    data = book.create_sheet("Data Src")
    calc["A1"] = 2
    calc["B1"] = "=A1*2"
    calc["C1"] = "=SUM(B:B)"
    calc["D1"] = "=Rate*10"
    calc["E1"] = "=7"
    data["B2"] = 3
    book.defined_names["Rate"] = DefinedName("Rate", attr_text="'Data Src'!$B$2")
    book.save(path)
    _replace_formula_values(path, {"xl/worksheets/sheet1.xml": {
        "B1": ("4", None), "C1": ("4", None), "D1": ("30", None), "E1": ("7", None)}})
    return path


class ExcelEngineTests(unittest.TestCase):
    def load(self):
        from pp.engine.excel import ExcelEngine, FormualizerHost

        TMP.mkdir(parents=True, exist_ok=True)
        directory = tempfile.TemporaryDirectory(dir=TMP)
        self.addCleanup(directory.cleanup)
        source = workbook(directory.name)
        engine = ExcelEngine.load(source, cache_dir=Path(directory.name) / "cache",
                                  host_factory=FormualizerHost)
        self.addCleanup(engine.close)
        return engine

    def test_edit_publishes_reachable_changes_and_reverting_removes_them(self):
        engine = self.load()
        self.assertEqual(engine.get(0, 0, 1), 4)

        changes = engine.apply([(0, 0, 0, 5)])

        self.assertEqual(changes, [(0, 0, 1, 10), (0, 0, 2, 10)])
        self.assertEqual(engine.overlay_snapshot(), [(0, 0, 1, 10), (0, 0, 2, 10)])
        engine.apply([(0, 0, 0, 2)])
        self.assertEqual(engine.overlay_snapshot(), [])
        self.assertEqual(engine.overlay_delta(), ([], [(0, 0, 1), (0, 0, 2)]))

    def test_cross_sheet_name_and_clear(self):
        engine = self.load()
        self.assertEqual(engine.apply([(1, 1, 1, 4)]), [(0, 0, 3, 40)])
        self.assertEqual(engine.apply([(0, 0, 0, None)]), [(0, 0, 1, 0), (0, 0, 2, 0)])
        self.assertTrue(engine.handles_clears)

    def test_unrelated_formulas_are_never_read(self):
        engine = self.load()
        with mock.patch.object(engine.host, "read", wraps=engine.host.read) as read:
            engine.apply([(1, 1, 1, 9)])
        self.assertEqual([call.args[0] for call in read.call_args_list], [0])
        self.assertEqual(engine.get(0, 0, 4), 7)


class ExcelWorkerTests(unittest.TestCase):
    def test_worker_uses_excel_engine_and_clears_without_reloading(self):
        from pp.engine.worker import WarmCopy

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as directory, \
                mock.patch.dict("os.environ", {"PP_ENGINE_CACHE_DIR": str(Path(directory) / "cache")}):
            source = workbook(directory)
            copy = WarmCopy("copy", "excel")
            first = copy.apply(1, str(source), [(0, 0, 0, 5)])
            with mock.patch.object(WarmCopy, "_reload_with_clears", side_effect=AssertionError("reload")):
                cleared = copy.apply(2, str(source), [(0, 0, 0, None)])
            engine = copy.engine
            with mock.patch.object(engine, "close", wraps=engine.close) as close:
                copy.close()

        self.assertEqual(type(engine).__name__, "ExcelEngine")
        self.assertEqual(first["changes"], [(0, 0, 1, 10), (0, 0, 2, 10)])
        self.assertEqual(cleared["changes"], [(0, 0, 1, 0), (0, 0, 2, 0)])
        close.assert_called_once()


class ExcelRuntimeTests(unittest.TestCase):
    def test_excel_runtime_warms_template_caches_and_builds_an_excel_pool(self):
        import json

        from pp.runtime import create_runtime
        from tests.v2.test_end_to_end import runtime_root

        temporary, root = runtime_root()
        self.addCleanup(temporary.cleanup)
        settings = json.loads((root / "settings.json").read_text())
        (root / "settings.json").write_text(json.dumps({**settings, "engine": "excel"}))
        with mock.patch("pp.engine.excel.warm_caches") as warm, \
                mock.patch("threading.Thread.start", lambda self: self.run()):
            runtime = create_runtime(root)
        self.addCleanup(runtime.close)
        warm.assert_called_once_with(runtime.snapshot.source_path)
        self.assertEqual(runtime.snapshot.results, "saved")  # no import engine pass, edits stay on
        self.assertEqual(runtime.pool._engine, "excel")


class ValueConversionTests(unittest.TestCase):
    def test_com_values(self):
        from pp.engine.excel import COM_ERROR_BASE, from_com, to_com

        self.assertEqual(from_com(COM_ERROR_BASE + 2007), {"error": "#DIV/0!"})
        self.assertEqual(from_com(COM_ERROR_BASE + 2042), {"error": "#N/A"})
        self.assertEqual(from_com(5.0), 5)
        self.assertEqual(from_com(2.5), 2.5)
        self.assertIs(from_com(True), True)
        self.assertEqual(from_com("x"), "x")
        self.assertIsNone(from_com(None))
        self.assertEqual(from_com({"error": "#REF!"}), {"error": "#REF!"})
        self.assertEqual(to_com("1-2"), "'1-2")
        self.assertEqual(to_com(""), "")
        self.assertEqual(to_com(3), 3)


class FakeRange:
    def __init__(self, sheet, cells):
        self.sheet, self.cells = sheet, cells

    @property
    def Value2(self):
        (top, left), (bottom, right) = self.cells
        rows = tuple(tuple(self.sheet.values.get((r, c)) for c in range(left, right + 1))
                     for r in range(top, bottom + 1))
        return rows[0][0] if (top, left) == (bottom, right) else rows


class FakeCell:
    def __init__(self, sheet, row, col):
        self.sheet, self.at = sheet, (row, col)

    @property
    def Value2(self):
        return self.sheet.values.get(self.at)

    @Value2.setter
    def Value2(self, value):
        self.sheet.values[self.at] = value


class FakeSheet:
    def __init__(self):
        self.values = {}

    def Cells(self, row, col):
        return FakeCell(self, row, col)

    def Range(self, first, last):
        return FakeRange(self, (first.at, last.at))


class FakeApp:
    def __init__(self):
        self.Hwnd, self.calculated, self.quit, self.sheets = 77, 0, False, {}
        self.Workbooks = types.SimpleNamespace(Open=self.open)

    def open(self, path, **options):
        self.opened = (path, options)
        return types.SimpleNamespace(Worksheets=lambda name: self.sheets.setdefault(name, FakeSheet()),
                                     Close=lambda SaveChanges: setattr(self, "closed", SaveChanges))

    def Calculate(self):
        self.calculated += 1

    def Quit(self):
        self.quit = True


class ComHostTests(unittest.TestCase):
    def test_hidden_macro_free_manual_excel_on_a_private_copy(self):
        app = FakeApp()
        killed = []
        modules = {
            "pythoncom": types.SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None),
            "win32com.client": types.SimpleNamespace(DispatchEx=lambda name: app),
            "win32process": types.SimpleNamespace(GetWindowThreadProcessId=lambda hwnd: (1, 4242)),
        }
        modules["win32com"] = types.SimpleNamespace(client=modules["win32com.client"])
        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as directory, \
                mock.patch.dict(sys.modules, modules), \
                mock.patch("pp.engine.excel._kill", side_effect=killed.append):
            from pp.engine.excel import ComExcelHost

            source = Path(directory) / "plan.xlsm"
            source.write_bytes(b"template")
            host = ComExcelHost(source, ["Calc", "Data"])
            host.write(0, [(0, 0, 5), (1, 2, "1-2")])
            host.calculate()
            app.sheets["Calc"].values[(2, 2)] = 7.0  # Excel row 2, column B
            block = host.read(0, 0, 0, 1, 2)
            single = host.read(0, 0, 0, 0, 0)
            opened = Path(app.opened[0])
            self.assertNotEqual(opened, source)
            self.assertEqual(opened.read_bytes(), b"template")
            host.close()

        self.assertEqual((app.Visible, app.DisplayAlerts, app.EnableEvents, app.AutomationSecurity,
                          app.Calculation), (False, False, False, 1, -4135))
        self.assertEqual(app.opened[1]["UpdateLinks"], 0)
        self.assertEqual(app.calculated, 1)
        self.assertEqual(block, [[5, None, None], [None, 7.0, "'1-2"]])
        self.assertEqual(single, [[5]])
        self.assertTrue(app.quit)
        self.assertEqual(killed, [4242])
        self.assertFalse(opened.parent.exists())


if __name__ == "__main__":
    unittest.main()
