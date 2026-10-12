"""Excel macros: what a macro changed is found by Excel and replayed as overrides."""

import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

from openpyxl import Workbook

from tests.v2.support import TMP
from tests.v2.test_engine_cached import _replace_formula_values

EXCEL_PID, OTHER_PID = 4242, 5555
MAIN, DIALOG, OTHER = 9000, 9001, 9002
NO_CELLS = -2146827284


def macro_workbook(directory):
    """Calc!A1 input 2; B1 = A1*2 (cached 4); C1 = A2+1 (cached 1) reaches A2 only."""
    path = Path(directory) / "macro.xlsx"
    book = Workbook()
    calc = book.active
    calc.title = "Calc"
    calc["A1"] = 2
    calc["B1"] = "=A1*2"
    calc["C1"] = "=A2+1"
    book.save(path)
    _replace_formula_values(path, {"xl/worksheets/sheet1.xml": {"B1": ("4", None), "C1": ("1", None)}})
    return path


def set_a1(host):
    host.write(0, [(0, 0, 5)])
    return []


def chain(host):
    host.write_formulas(0, [(1, 0, "=A1*3")])
    return []


def detach_b1(host):
    host.write(0, [(0, 1, 9)])
    return []


def set_and_chain(host):
    host.write(0, [(0, 0, 5)])
    host.write_formulas(0, [(1, 0, "=A1*3"), (1, 1, "=A2+1")])
    return ["Chain built"]


class MacroEngineTests(unittest.TestCase):
    def load(self, macros):
        from pp.engine.excel import ExcelEngine, FormualizerHost

        TMP.mkdir(parents=True, exist_ok=True)
        directory = tempfile.TemporaryDirectory(dir=TMP)
        self.addCleanup(directory.cleanup)
        source = macro_workbook(directory.name)

        def factory(path, names):
            host = FormualizerHost(path, names)
            host.macros = macros
            return host

        engine = ExcelEngine.load(source, cache_dir=Path(directory.name) / "cache", host_factory=factory)
        self.addCleanup(engine.close)
        return engine

    def test_constant_changes_publish_dependents(self):
        engine = self.load({"SetA1": set_a1})

        result = engine.run_macro("SetA1")

        self.assertEqual(result, {"messages": [], "contents": [(0, 0, 0, 5)], "changes": [(0, 0, 1, 10)]})
        self.assertEqual(engine.overlay_snapshot(), [(0, 0, 1, 10)])

    def test_macro_formula_is_an_override_and_edits_reach_it_through_both_indexes(self):
        engine = self.load({"Chain": chain})

        result = engine.run_macro("Chain")

        self.assertEqual(result["contents"], [(0, 1, 0, {"formula": "=A1*3"})])
        self.assertEqual(result["changes"], [(0, 0, 2, 7), (0, 1, 0, 6)])
        # C1 is a template formula that only reads the override A2: the fixpoint must reach it
        self.assertEqual(engine.apply([(0, 0, 0, 4)]), [(0, 0, 1, 8), (0, 0, 2, 13), (0, 1, 0, 12)])
        self.assertEqual(engine.get(0, 1, 0), 12)

    def test_macro_turns_a_formula_into_a_value(self):
        engine = self.load({"Detach": detach_b1})

        result = engine.run_macro("Detach")

        self.assertEqual(result, {"messages": [], "contents": [(0, 0, 1, 9)], "changes": []})
        self.assertNotIn((0, 0, 1), engine.formula_anchors)
        self.assertEqual(engine.apply([(0, 0, 0, 3)]), [])
        self.assertEqual(engine.get(0, 0, 1), 9)

    def test_replay_of_recorded_contents_reads_identical_values(self):
        original = self.load({"Build": set_and_chain})
        replay = self.load({})

        result = original.run_macro("Build")
        replay.apply(result["contents"])

        self.assertEqual(result["messages"], ["Chain built"])
        cells = [(0, row, col) for row in range(2) for col in range(3)]
        self.assertEqual([original.get(*cell) for cell in cells], [replay.get(*cell) for cell in cells])
        self.assertEqual([original.get(*cell) for cell in cells], [5, 10, 16, 15, 16, None])


class FakeExcelSheet:
    """Excel worksheet with 1-based cells: values, formulas and used range."""

    def __init__(self):
        self.values, self.formulas = {}, {}

    def Cells(self, row, col):
        return FakeExcelCell(self, (row, col))

    def Range(self, first, last):
        return FakeExcelRange(self, first.at, last.at)

    @property
    def UsedRange(self):
        cells = set(self.values) | set(self.formulas)
        rows = max((row for row, _ in cells), default=1)
        cols = max((col for _, col in cells), default=1)
        return FakeExcelRange(self, (1, 1), (rows, cols))


class FakeExcelCell:
    def __init__(self, sheet, at):
        self.sheet, self.at = sheet, at

    @property
    def Value2(self):
        return self.sheet.values.get(self.at)

    @Value2.setter
    def Value2(self, value):
        self.sheet.values[self.at] = value
        self.sheet.formulas.pop(self.at, None)

    @property
    def Formula2(self):
        return self.sheet.formulas.get(self.at, self.sheet.values.get(self.at))

    @Formula2.setter
    def Formula2(self, text):
        self.sheet.formulas[self.at] = text
        self.sheet.values.pop(self.at, None)


class FakeExcelRange:
    def __init__(self, sheet, first, last):
        self.sheet, self.first, self.last = sheet, first, last

    @property
    def Row(self):
        return self.first[0]

    @property
    def Column(self):
        return self.first[1]

    @property
    def Rows(self):
        return SimpleNamespace(Count=self.last[0] - self.first[0] + 1)

    @property
    def Columns(self):
        return SimpleNamespace(Count=self.last[1] - self.first[1] + 1)

    def _grid(self, lookup):
        rows = tuple(tuple(lookup(at) for at in ((row, col) for col in range(self.first[1], self.last[1] + 1)))
                     for row in range(self.first[0], self.last[0] + 1))
        return rows[0][0] if self.first == self.last else rows

    @property
    def Value2(self):
        return self._grid(self.sheet.values.get)

    @property
    def Formula2(self):
        return self._grid(lambda at: self.sheet.formulas.get(at, self.sheet.values.get(at)))

    def SpecialCells(self, kind):
        if not self.sheet.formulas:
            raise ComError(NO_CELLS)
        areas = [FakeExcelRange(self.sheet, at, at) for at in sorted(self.sheet.formulas)]
        return SimpleNamespace(Areas=SimpleNamespace(Count=len(areas), Item=lambda index: areas[index - 1]))


class ComError(Exception):
    def __init__(self, hresult):
        super().__init__(hresult)
        self.hresult = hresult


class FakeGui:
    """Top-level windows and their children, as win32gui reports them."""

    def __init__(self):
        self.info = {  # handle: (class, process, text)
            MAIN: ("XLMAIN", EXCEL_PID, ""),
            DIALOG: ("#32770", EXCEL_PID, ""),
            OTHER: ("#32770", OTHER_PID, ""),
        }
        self.children = {DIALOG: [(701, "Static", "Done: 3 rows"), (702, "Static", "Saved"),
                                  (703, "Button", "OK")], OTHER: [(801, "Static", "Not ours")]}
        for owner in self.children.values():
            for child, kind, text in owner:
                self.info[child] = (kind, None, text)
        self.shown = {MAIN, OTHER}
        self.posted = []
        self.closed = threading.Event()

    def show(self):
        self.shown.add(DIALOG)

    def process_of(self, hwnd):
        return (1, EXCEL_PID if hwnd == 77 else self.info[hwnd][1])

    def EnumWindows(self, callback, extra):
        for hwnd in sorted(self.shown):
            callback(hwnd, extra)

    def GetClassName(self, hwnd):
        return self.info[hwnd][0]

    def GetWindowText(self, hwnd):
        return self.info[hwnd][2]

    def EnumChildWindows(self, hwnd, callback, extra):
        for child, _, _ in self.children.get(hwnd, []):
            callback(child, extra)

    def PostMessage(self, hwnd, message, wparam, lparam):
        self.posted.append((hwnd, message, wparam, lparam))
        if hwnd == DIALOG:
            self.shown.discard(DIALOG)
            self.closed.set()


class ComHostTests(unittest.TestCase):
    def com_host(self, app, gui, killed):
        from pp.engine.excel import ComExcelHost

        client = SimpleNamespace(DispatchEx=lambda name: app)
        modules = {
            "pythoncom": SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None),
            "win32com": SimpleNamespace(client=client),
            "win32com.client": client,
            "win32process": SimpleNamespace(GetWindowThreadProcessId=gui.process_of),
            "win32gui": gui,
            "win32con": SimpleNamespace(WM_COMMAND=0x111, IDOK=1),
        }
        TMP.mkdir(parents=True, exist_ok=True)
        directory = tempfile.TemporaryDirectory(dir=TMP)
        self.addCleanup(directory.cleanup)
        patches = [mock.patch.dict(sys.modules, modules),
                   mock.patch("pp.engine.excel._kill", side_effect=lambda pid: (killed.append(pid), app.killed.set()))]
        for patch in patches:
            patch.start()
            self.addCleanup(patch.stop)
        source = Path(directory.name) / "plan.xlsm"
        source.write_bytes(b"template")
        host = ComExcelHost(source, ["Calc"])
        self.addCleanup(host.close)
        return host

    def test_dialog_is_pressed_ok_and_its_text_returned(self):
        gui, killed = FakeGui(), []
        app = MacroApp(gui, lambda app, name: run_showing_dialog(app, name))
        host = self.com_host(app, gui, killed)

        messages = host.run_macro("Refresh", 5)

        self.assertEqual(messages, ["Done: 3 rows\nSaved"])
        self.assertEqual(gui.posted, [(DIALOG, 0x111, 1, 0)])
        self.assertEqual(app.ran, "'plan.xlsm'!Refresh")
        self.assertEqual(killed, [])

    def test_macro_that_runs_past_the_limit_gets_excel_killed(self):
        gui, killed = FakeGui(), []
        app = MacroApp(gui, lambda app, name: wait_until_killed(app))
        host = self.com_host(app, gui, killed)

        with self.assertRaises(Exception) as raised:
            host.run_macro("Hang", 0.3)

        self.assertEqual(str(raised.exception), "macro timed out")
        self.assertEqual(killed, [EXCEL_PID])

    def test_com_error_from_run_is_a_macro_failure(self):
        gui, killed = FakeGui(), []
        app = MacroApp(gui, lambda app, name: fail(app, name))
        host = self.com_host(app, gui, killed)

        with self.assertRaises(Exception) as raised:
            host.run_macro("Broken", 5)

        self.assertEqual(str(raised.exception), "macro failed")

    def test_formula_and_used_range_calls_on_a_sheet(self):
        gui, killed = FakeGui(), []
        app = MacroApp(gui, lambda app, name: None)
        sheet = app.sheets.setdefault("Calc", FakeExcelSheet())
        host = self.com_host(app, gui, killed)

        self.assertEqual(host.formula_areas(0), [])
        self.assertIsNone(host.used_range(0))
        sheet.values[(1, 1)] = 5
        sheet.formulas[(2, 2)] = "=A1*2"
        self.assertEqual(host.formula_areas(0), [(1, 1, 1, 1)])
        self.assertEqual(host.used_range(0), (0, 0, 1, 1))
        self.assertEqual(host.read_formulas(0, 0, 0, 1, 1), [[None, None], [None, "=A1*2"]])
        host.write_formulas(0, [(2, 0, "=A1*3")])
        self.assertEqual(sheet.formulas[(3, 1)], "=A1*3")
        self.assertEqual(app.AutomationSecurity, 1)
        self.assertEqual(app.EnableEvents, False)


class MacroApp:
    Hwnd = 77

    def __init__(self, gui, run):
        self.gui, self.run_impl = gui, run
        self.sheets, self.killed = {}, threading.Event()
        self.Workbooks = SimpleNamespace(Open=self.open)

    def open(self, path, **options):
        self.opened = Path(path)
        return SimpleNamespace(Name=Path(path).name,
                               Worksheets=lambda name: self.sheets.setdefault(name, FakeExcelSheet()),
                               Close=lambda SaveChanges: None)

    def Run(self, name):
        self.ran = name
        return self.run_impl(self, name)

    def Calculate(self):
        pass

    def Quit(self):
        pass


def run_showing_dialog(app, name):
    time.sleep(0.2)  # the macro takes a moment before its MsgBox appears
    app.gui.show()
    if not app.gui.closed.wait(5):
        raise RuntimeError("dialog never closed")


def wait_until_killed(app):
    if app.killed.wait(5):
        raise RuntimeError("the server is unavailable")


def fail(app, name):
    raise RuntimeError("macro error")


if __name__ == "__main__":
    unittest.main()
