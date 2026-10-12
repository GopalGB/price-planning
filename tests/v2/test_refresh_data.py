import json
import tempfile
import unittest
import zipfile
from pathlib import Path
from types import SimpleNamespace

from ops.fixtures import make_semantic
from pp.engine.formualizer import EngineError, MacroTimeout
from tests.v2.api_support import APIFixture
from tests.v2.support import TMP
from tests.v2.test_excel_macros import macro_workbook
from tests.v2.test_macro_buttons import write_button_fixture
from tests.v2.test_macro_api import FakeMacroPool

SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def add_connections(path, count=1):
    """Give a saved synthetic workbook xl/connections.xml with ``count`` Power Query connections."""
    with zipfile.ZipFile(path) as source:
        members = {info.filename: source.read(info.filename) for info in source.infolist()}
    body = "".join(f'<connection id="{number}" name="Query{number}" type="5"/>' for number in range(1, count + 1))
    members["xl/connections.xml"] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                                     f'<connections xmlns="{SPREADSHEET_NS}">{body}</connections>').encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in members.items():
            target.writestr(name, data)


def refreshable_fixture(path):
    write_button_fixture(path)
    add_connections(path)


def refresh_table(host):
    """Stands in for Excel's RefreshAll: a query writes A3 and a new formula in D1 reads it."""
    host.write(0, [(2, 0, 7)])
    host.write_formulas(0, [(0, 3, "=A3*10")])
    return ["Query loaded"]


def open_on_formualizer(directory, macros):
    from pp.engine.excel import ExcelEngine, FormualizerHost

    source = macro_workbook(directory)

    def factory(path, names):
        host = FormualizerHost(path, names)
        host.macros = macros
        return host

    return source, lambda: ExcelEngine.load(source, cache_dir=Path(directory) / "cache", host_factory=factory)


class RefreshableFlagTests(unittest.TestCase):
    def import_snapshot(self, builder):
        from pp.template.importer import import_template

        TMP.mkdir(parents=True, exist_ok=True)
        directory = tempfile.TemporaryDirectory(dir=TMP)
        self.addCleanup(directory.cleanup)
        source = Path(directory.name) / "template.xlsm"
        builder(source)
        return import_template(source, Path(directory.name) / "data")

    def test_import_marks_a_workbook_with_a_data_connection_as_refreshable(self):
        meta = json.loads((self.import_snapshot(refreshable_fixture) / "meta.json").read_text())
        self.assertIs(meta["refreshable"], True)

    def test_import_marks_a_workbook_without_connections_as_not_refreshable(self):
        meta = json.loads((self.import_snapshot(make_semantic) / "meta.json").read_text())
        self.assertIs(meta["refreshable"], False)

    def test_a_connections_part_with_no_connection_is_not_refreshable(self):
        def empty(path):
            make_semantic(path)
            with zipfile.ZipFile(path) as source:
                members = {info.filename: source.read(info.filename) for info in source.infolist()}
            members["xl/connections.xml"] = f'<connections xmlns="{SPREADSHEET_NS}"/>'.encode()
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
                for name, data in members.items():
                    target.writestr(name, data)

        meta = json.loads((self.import_snapshot(empty) / "meta.json").read_text())
        self.assertIs(meta["refreshable"], False)

    def test_snapshot_without_the_flag_reads_as_not_refreshable(self):
        from pp.workbook_service import TemplateSnapshot

        snapshot = self.import_snapshot(make_semantic)
        meta_path = snapshot / "meta.json"
        meta = json.loads(meta_path.read_text())
        del meta["refreshable"]
        meta_path.write_text(json.dumps(meta))
        self.assertIs(TemplateSnapshot(snapshot).refreshable, False)


class RefreshEngineTests(unittest.TestCase):
    def load(self, macros):
        TMP.mkdir(parents=True, exist_ok=True)
        directory = tempfile.TemporaryDirectory(dir=TMP)
        self.addCleanup(directory.cleanup)
        _, open_engine = open_on_formualizer(directory.name, macros)
        engine = open_engine()
        self.addCleanup(engine.close)
        return engine

    def test_refresh_records_what_excel_wrote_and_what_it_changed(self):
        engine = self.load({"RefreshAll": refresh_table})

        result = engine.refresh_data()

        self.assertEqual(result["messages"], ["Query loaded"])
        self.assertEqual(result["contents"], [(0, 0, 3, {"formula": "=A3*10"}), (0, 2, 0, 7)])
        self.assertEqual(result["changes"], [(0, 0, 3, 70)])
        self.assertEqual(result["before"], {(0, 0, 3): None, (0, 2, 0): None})
        self.assertEqual(engine.get(0, 0, 3), 70)

    def test_refresh_contents_replay_on_a_fresh_engine_to_the_same_values(self):
        original = self.load({"RefreshAll": refresh_table})
        replay = self.load({})
        result = original.refresh_data()

        replay.apply(result["contents"])

        self.assertEqual(replay.get(0, 0, 3), original.get(0, 0, 3))
        self.assertEqual(replay.get(0, 2, 0), 7)


class RefreshWorkerTests(unittest.TestCase):
    def test_refresh_macro_name_runs_the_excel_refresh_in_the_warm_copy(self):
        from pp.engine.excel import ExcelEngine, FormualizerHost
        from pp.engine.worker import WarmCopy

        TMP.mkdir(parents=True, exist_ok=True)
        directory = tempfile.TemporaryDirectory(dir=TMP)
        self.addCleanup(directory.cleanup)
        source = macro_workbook(directory.name)

        def factory(path, names):
            host = FormualizerHost(path, names)
            host.macros = {"RefreshAll": refresh_table}
            return host

        copy = WarmCopy("copy-r", "excel")
        copy.source = str(source)
        copy.engine = ExcelEngine.load(source, cache_dir=Path(directory.name) / "cache", host_factory=factory)
        self.addCleanup(copy.close)

        result = copy.run_macro(1, str(source), [], None, "RefreshAll", 600)

        self.assertEqual(result["messages"], ["Query loaded"])
        self.assertEqual(result["contents"], [(0, 0, 3, {"formula": "=A3*10"}), (0, 2, 0, 7)])
        self.assertIn((0, 0, 3, 70), result["results"])


REFRESH_OUTCOME = {"messages": ["Query loaded"], "results": [(0, 0, 3, 70)],
                   "contents": [(0, 9, 0, "Hi")]}


class RefreshEndpointTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(builder=refreshable_fixture)
        self.client, self.csrf, self.user = self.fixture.signup("refresh@example.test", "Refresher")
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)
        self.copy_id = self.copy["id"]
        self.pool = FakeMacroPool(REFRESH_OUTCOME)
        self.fixture.app.state.services.engine_pool = self.pool

    def tearDown(self):
        self.fixture.close()

    def refresh(self, op_id="refresh-1"):
        return self.client.post(f"/api/copies/{self.copy_id}/refresh", {"op_id": op_id},
                                {"x-csrf-token": self.csrf})

    def history(self):
        with self.fixture.database.connect() as connection:
            return [tuple(row) for row in connection.execute(
                "SELECT sheet,row,col,action,revision FROM pp_history WHERE copy_id=? ORDER BY sheet,row,col",
                (self.copy_id,)).fetchall()]

    def test_refresh_commits_as_one_revision_with_history_action_refresh(self):
        response = self.refresh()

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json(), {"messages": ["Query loaded"], "changed": 1, "input_revision": 1})
        self.assertEqual(self.pool.requests[-1].macro, "RefreshAll")
        self.assertEqual(self.history(), [(0, 9, 0, "refresh", 1)])

    def test_refreshed_cells_are_allowed_past_the_locked_cell_guard(self):
        self.refresh()

        workbooks = self.fixture.app.state.services.workbooks
        _, overrides = workbooks.source_and_overrides(self.copy_id, self.user["id"])

        self.assertEqual(overrides[(0, 9, 0)], "Hi")

    def test_repeated_op_id_returns_the_stored_ack_without_refreshing_again(self):
        first = self.refresh()
        second = self.refresh()

        self.assertEqual((first.status_code, second.status_code), (200, 200))
        self.assertEqual(second.json(), first.json())
        self.assertEqual(len(self.pool.requests), 1)

    def test_refresh_is_refused_for_a_template_without_connections(self):
        fixture = APIFixture(builder=write_button_fixture)
        self.addCleanup(fixture.close)
        client, csrf, _ = fixture.signup("plain@example.test", "Plain")
        _, copy = fixture.make_copy(client, csrf)
        fixture.app.state.services.engine_pool = FakeMacroPool(REFRESH_OUTCOME)

        response = client.post(f"/api/copies/{copy['id']}/refresh", {"op_id": "r-1"}, {"x-csrf-token": csrf})

        self.assertEqual(response.status_code, 409)
        self.assertEqual(response.json()["error"], "refresh_unavailable")
        self.assertEqual(response.json()["message"], "This template has no data connections to refresh.")

    def test_refresh_needs_the_excel_engine(self):
        self.fixture.app.state.services.engine_pool = None

        response = self.refresh()

        self.assertEqual((response.status_code, response.json()["error"]), (409, "macros_unavailable"))

    def test_viewer_is_forbidden_and_csrf_is_required(self):
        viewer, viewer_csrf, _ = self.fixture.signup("refresh-viewer@example.test", "Viewer")
        self.client.put(f"/api/copies/{self.copy_id}/shares",
                        {"email": "refresh-viewer@example.test", "role": "viewer"},
                        {"x-csrf-token": self.csrf})

        response = viewer.post(f"/api/copies/{self.copy_id}/refresh", {"op_id": "r-2"},
                               {"x-csrf-token": viewer_csrf})
        missing = self.client.post(f"/api/copies/{self.copy_id}/refresh", {"op_id": "r-3"})

        self.assertEqual((response.status_code, response.json()["error"]), (403, "forbidden"))
        self.assertEqual(missing.json()["error"], "csrf")

    def test_engine_failure_is_502_and_commits_nothing(self):
        self.pool.error = EngineError("boom")

        response = self.refresh()

        self.assertEqual((response.status_code, response.json()["error"]), (502, "macro_failed"))
        self.assertEqual(self.history(), [])

    def test_refresh_past_the_time_limit_is_502_and_commits_nothing(self):
        self.pool.error = MacroTimeout("slow")

        response = self.refresh()

        self.assertEqual((response.status_code, response.json()["error"]), (502, "macro_failed"))
        self.assertEqual(self.history(), [])


class RefreshSmokeTests(unittest.TestCase):
    def test_smoke_refresh_prints_counts_and_replays_on_a_fresh_engine(self):
        from ops.excel_smoke import check_refresh

        TMP.mkdir(parents=True, exist_ok=True)
        directory = tempfile.TemporaryDirectory(dir=TMP)
        self.addCleanup(directory.cleanup)
        source, open_engine = open_on_formualizer(directory.name, {"RefreshAll": refresh_table})
        snapshot = SimpleNamespace(source_path=source, sheets=[{"rows": 3, "cols": 4}], refreshable=True)

        lines = []
        self.assertEqual(check_refresh(snapshot, lambda path: open_engine(), lines.append), 0)

        self.assertIn("messages 1", lines[0])
        self.assertIn("contents 2", lines[0])
        self.assertIn("published 1", lines[0])
        self.assertIn("mismatched cells 0", lines[1])
        self.assertEqual(lines[-1], "PASS")
        printed = "\n".join(lines)
        self.assertNotIn("Query loaded", printed)
        self.assertNotIn("A3", printed)

    def test_smoke_refresh_failure_counts_as_a_failure(self):
        from ops.excel_smoke import check_refresh

        TMP.mkdir(parents=True, exist_ok=True)
        directory = tempfile.TemporaryDirectory(dir=TMP)
        self.addCleanup(directory.cleanup)

        def broken(host):
            raise EngineError("refresh failed")

        source, open_engine = open_on_formualizer(directory.name, {"RefreshAll": broken})
        snapshot = SimpleNamespace(source_path=source, sheets=[{"rows": 3, "cols": 4}], refreshable=True)

        lines = []
        self.assertEqual(check_refresh(snapshot, lambda path: open_engine(), lines.append), 1)
        self.assertEqual(lines[-1], "FAIL (1)")


if __name__ == "__main__":
    unittest.main()
