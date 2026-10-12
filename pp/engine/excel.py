"""Calculate copies in Microsoft Excel itself (office Windows server).

Excel recalculates natively and exactly. The engine writes a batch of input
values, asks Excel to calculate, then reads back only the formulas the static
dependency index says the batch can reach, in bulk row blocks, and diffs them
against the last value it knew (Excel's saved value until something changes).

The host is the only part that talks to Excel: ``ComExcelHost`` drives a
hidden, private Excel instance through pywin32; ``FormualizerHost`` stands in
for it in tests on machines without Excel.
"""

from __future__ import annotations

import hashlib
import os
import pickle
import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

from ops.typed import typed

from ..values import exactly_equal
from .depindex import DependencyIndex
from .formualizer import EngineError, _cached_formula_map, _json_value
from .native_load import prepared_template

BLOCK_ROWS = 256
# COM reports a cell error as the HRESULT 0x800A0000 + Excel's CVErr number.
COM_ERROR_BASE = -2146828288
EXCEL_ERRORS = {2000: "#NULL!", 2007: "#DIV/0!", 2015: "#VALUE!", 2023: "#REF!", 2029: "#NAME?",
                2036: "#NUM!", 2042: "#N/A", 2043: "#GETTING_DATA", 2045: "#SPILL!",
                2047: "#BLOCKED!", 2048: "#UNKNOWN!", 2049: "#FIELD!", 2050: "#CALC!"}


def from_com(value):
    """Convert a COM Value2 into the app's JSON cell value."""
    if value is None or isinstance(value, (bool, str)):
        return value
    if isinstance(value, dict) and set(value) == {"error"}:
        return value
    if isinstance(value, int) and not isinstance(value, bool):
        code = value - COM_ERROR_BASE
        if code in EXCEL_ERRORS:
            return {"error": EXCEL_ERRORS[code]}
    if isinstance(value, float) and value.is_integer() and abs(value) < 2**53:
        return int(value)
    kind, normalized = typed(value)
    return {"error": normalized} if kind == "error" else normalized


def to_com(value):
    """Convert an input value for Value2; text is stored literally, never re-parsed."""
    if isinstance(value, dict) and set(value) == {"error"}:
        raise EngineError("errors cannot be typed into input cells")
    if isinstance(value, str) and value:
        return "'" + value
    return value


class ExcelHost:
    """What the engine needs from Excel. Coordinates are zero-based."""

    def write(self, sheet: int, cells: list[tuple[int, int, object]]) -> None: ...
    def calculate(self) -> None: ...
    def read(self, sheet: int, top: int, left: int, bottom: int, right: int) -> list[list]: ...
    def memory_bytes(self) -> int: ...
    def close(self) -> None: ...


class ComExcelHost(ExcelHost):
    """A private hidden Excel with macros and events off, on its own file copy."""

    def __init__(self, workbook: Path, sheet_names: list[str]):
        import pythoncom
        import win32com.client
        import win32process

        pythoncom.CoInitialize()
        self.directory = tempfile.mkdtemp(prefix="pp-excel-")
        self.path = Path(self.directory) / workbook.name
        shutil.copy2(workbook, self.path)
        self.app = win32com.client.DispatchEx("Excel.Application")
        self.pid = win32process.GetWindowThreadProcessId(self.app.Hwnd)[1]
        try:
            for name, value in (("Visible", False), ("DisplayAlerts", False), ("ScreenUpdating", False),
                                ("EnableEvents", False), ("AskToUpdateLinks", False),
                                ("AutomationSecurity", 3)):  # 3 = msoAutomationSecurityForceDisable
                setattr(self.app, name, value)
            self.book = self.app.Workbooks.Open(str(self.path), UpdateLinks=0, ReadOnly=False,
                                                IgnoreReadOnlyRecommended=True, AddToMru=False,
                                                Notify=False)
            self.app.Calculation = -4135  # xlCalculationManual: recalculate once per batch
            self.sheets = [self.book.Worksheets(name) for name in sheet_names]
        except Exception:
            self.close()
            raise

    def write(self, sheet, cells):
        worksheet = self.sheets[sheet]
        for row, col, value in cells:
            worksheet.Cells(row + 1, col + 1).Value2 = to_com(value)

    def calculate(self):
        self.app.Calculate()

    def read(self, sheet, top, left, bottom, right):
        worksheet = self.sheets[sheet]
        values = worksheet.Range(worksheet.Cells(top + 1, left + 1),
                                 worksheet.Cells(bottom + 1, right + 1)).Value2
        if not isinstance(values, tuple):
            return [[values]]
        return [list(row) for row in values]

    def memory_bytes(self):
        import win32api
        import win32con
        import win32process

        handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ,
                                      False, self.pid)
        try:
            return int(win32process.GetProcessMemoryInfo(handle)["PeakWorkingSetSize"])
        finally:
            win32api.CloseHandle(handle)

    def close(self):
        import pythoncom

        try:
            if getattr(self, "book", None) is not None:
                self.book.Close(SaveChanges=False)
            self.app.Quit()
        except Exception:
            pass
        finally:
            self.book = self.sheets = None
            self.app = None
            _kill(getattr(self, "pid", None))
            shutil.rmtree(self.directory, ignore_errors=True)
            pythoncom.CoUninitialize()


def _kill(pid):
    """Make sure a hung Excel never outlives its worker."""
    if not pid:
        return
    try:
        import win32api
        import win32con

        handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
        win32api.TerminateProcess(handle, 0)
        win32api.CloseHandle(handle)
    except Exception:
        pass


class FormualizerHost(ExcelHost):
    """Test stand-in: formualizer computes what Excel would for small workbooks."""

    def __init__(self, workbook: Path, sheet_names: list[str]):
        import formualizer

        self.book = formualizer.load_workbook(str(workbook))
        self.names = sheet_names

    def write(self, sheet, cells):
        for row, col, value in cells:
            self.book.set_value(self.names[sheet], row + 1, col + 1, value)

    def calculate(self):
        pass

    def read(self, sheet, top, left, bottom, right):
        name = self.names[sheet]
        cells = [(name, row + 1, col + 1) for row in range(top, bottom + 1)
                 for col in range(left, right + 1)]
        self.book.evaluate_cells(cells)
        width = right - left + 1
        values = [_json_value(self.book.get_value(*cell), False) for cell in cells]
        return [values[index:index + width] for index in range(0, len(values), width)]

    def memory_bytes(self):
        return 0

    def close(self):
        self.book = None


def default_host():
    return ComExcelHost if os.name == "nt" else FormualizerHost


def _cached_index(prepared: Path, formulas: dict, names: list[str]) -> DependencyIndex:
    code = hashlib.sha256(Path(__file__).with_name("depindex.py").read_bytes()).hexdigest()[:12]
    path = prepared.with_name(f"{prepared.name}.index-{code}.pickle")
    try:
        with path.open("rb") as handle:
            return pickle.load(handle)
    except FileNotFoundError:
        pass
    except Exception:
        path.unlink(missing_ok=True)
    index = DependencyIndex.build(prepared, formulas, names)
    pending = path.with_name(f"{path.name}.{os.getpid()}.tmp")
    try:
        with pending.open("wb") as handle:
            pickle.dump(index, handle, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(pending, path)
    except OSError:
        pending.unlink(missing_ok=True)
    return index


def warm_caches(template_path: str | Path, cache_dir: str | Path | None = None) -> None:
    """Build the prepared package, formula map and index once, before the first edit needs them."""
    prepared = prepared_template(Path(template_path), cache_dir)
    _, formulas, names = _cached_formula_map(prepared)[:3]
    _cached_index(prepared, formulas, names)


@dataclass
class ExcelEngine:
    host: ExcelHost
    index: DependencyIndex
    formula_anchors: frozenset
    baseline: dict  # Excel's saved values for formula cells
    current: dict = field(default_factory=dict)  # formula cells whose value now differs
    overlay_upserts: dict = field(default_factory=dict)
    overlay_removals: set = field(default_factory=set)
    handles_clears = True

    @classmethod
    def load(cls, template_path: str | Path, cache_dir: str | Path | None = None,
             host_factory=None) -> ExcelEngine:
        path = Path(template_path)
        try:
            prepared = prepared_template(path, cache_dir)
            cells, formulas, names, cached = _cached_formula_map(prepared)[:4]
            index = _cached_index(prepared, formulas, names)
            host = (host_factory or default_host())(path, names)
        except EngineError:
            raise
        except Exception as exc:
            raise EngineError("Excel calculation engine failed to load") from exc
        baseline = {key: from_com(cached.get(key)) for key in cells}
        return cls(host, index, frozenset(cells), baseline)

    @property
    def formula_cells(self):
        return self.formula_anchors

    def apply(self, changes: list[tuple[int, int, int, object]]) -> list[tuple[int, int, int, object]]:
        self.overlay_upserts, self.overlay_removals = {}, set()
        if not changes:
            return []
        try:
            by_sheet: dict[int, list] = {}
            for sheet, row, col, value in changes:
                by_sheet.setdefault(sheet, []).append((row, col, value))
            for sheet, cells in by_sheet.items():
                self.host.write(sheet, cells)
            self.host.calculate()
            affected = self.index.cone([(sheet, row, col) for sheet, row, col, _ in changes])
            return self._publish(affected)
        except EngineError:
            raise
        except Exception as exc:
            raise EngineError("Excel calculation engine failed to apply changes") from exc

    def _publish(self, affected) -> list[tuple[int, int, int, object]]:
        published = []
        for sheet, top, left, bottom, right, members in _blocks(affected):
            values = self.host.read(sheet, top, left, bottom, right)
            for row, col in members:
                value = from_com(values[row - top][col - left])
                key = (sheet, row, col)
                if exactly_equal(value, self.current.get(key, self.baseline.get(key))):
                    continue
                published.append((sheet, row, col, value))
                if exactly_equal(value, self.baseline.get(key)):
                    self.current.pop(key, None)
                    self.overlay_removals.add(key)
                    self.overlay_upserts.pop(key, None)
                else:
                    self.current[key] = value
                    self.overlay_upserts[key] = value
                    self.overlay_removals.discard(key)
        return sorted(published, key=lambda item: item[:3])

    def get(self, sheet: int, row: int, col: int):
        key = (sheet, row, col)
        if key in self.formula_anchors:
            return self.current.get(key, self.baseline.get(key))
        try:
            return from_com(self.host.read(sheet, row, col, row, col)[0][0])
        except Exception as exc:
            raise EngineError("Excel calculation engine failed to read a cell") from exc

    def overlay_snapshot(self):
        return [(*key, value) for key, value in sorted(self.current.items())]

    def overlay_delta(self):
        upserts = [(*key, value) for key, value in sorted(self.overlay_upserts.items())]
        return upserts, sorted(self.overlay_removals)

    def memory_bytes(self) -> int:
        return self.host.memory_bytes()

    def close(self) -> None:
        self.host.close()


def _blocks(affected):
    """Group cells into row blocks with one bounding column span each."""
    grouped: dict[tuple[int, int], list] = {}
    for sheet, row, col in affected:
        grouped.setdefault((sheet, row // BLOCK_ROWS), []).append((row, col))
    for (sheet, _), members in sorted(grouped.items()):
        rows = [row for row, _ in members]
        cols = [col for _, col in members]
        yield sheet, min(rows), min(cols), max(rows), max(cols), members
