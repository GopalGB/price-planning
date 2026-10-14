"""POST /api/copies/<id>/slicers: an Excel slicer change is one revision, its selection is stored per copy."""

import unittest

import ops.fixtures as fixtures
from pp.engine.formualizer import EngineError, MacroTimeout
from pp.workbook_service import MACRO_MESSAGE, SLICER_MESSAGE
from tests.v2.api_support import APIFixture


def external_dashboard(path):
    fixtures.make_dashboard(path, row_count=32, external=True)


class FakeSlicerPool:
    """Canned Excel outcomes; records the requests the server sends."""

    kind = "excel"

    def __init__(self, error=None):
        self.error, self.requests, self.calc = error, [], []

    def apply(self, request):
        self.calc.append(request)
        return self._result(request, changes=[], results=[])

    def run_slicer(self, request):
        self.requests.append(request)
        if self.error is not None:
            raise self.error
        return self._result(request, changes=[(0, 2, 0, 100)], results=[(0, 2, 0, 100)])

    def run_macro(self, request):
        self.requests.append(request)
        return self._result(request, changes=[], results=[])

    def _result(self, request, changes, results):
        return {"copy_id": request.copy_id, "revision": request.revision, "pid": 0, "changes": changes,
                "results": results, "removals": [], "delta": False, "base_revision": request.base_revision,
                "messages": [], "contents": []}


class SlicerEndpointTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(builder=external_dashboard)
        self.client, self.csrf, self.user = self.fixture.signup("slicers@example.test", "Slicer Owner")
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)
        self.copy_id = self.copy["id"]
        self.pool = FakeSlicerPool()
        self.fixture.app.state.services.engine_pool = self.pool
        # the first slicer of the fixture has 10 items, only the first selected in the template
        first = next(item for item in self.excel_slicers() if len(item["items"]) == 10)
        self.cache = first["cache"]
        self.template_pick, self.pick = first["items"][0]["caption"], first["items"][1]["caption"]

    def tearDown(self):
        self.fixture.close()

    def press(self, body=None, op_id="slice-1", client=None, csrf=None):
        payload = {"op_id": op_id, "cache": self.cache, "selected": [self.pick]} if body is None else body
        return (client or self.client).post(f"/api/copies/{self.copy_id}/slicers", payload,
                                            {"x-csrf-token": csrf or self.csrf})

    def workbook(self, client=None):
        return (client or self.client).get(f"/api/copies/{self.copy_id}/workbook").json()

    def excel_slicers(self, workbook=None):
        sheet = next(item for item in (workbook or self.workbook())["sheets"] if item["name"] == "Dashboard")
        return sheet["slicers"]

    def pivot_state(self):
        return self.fixture.app.state.services.workbooks.pivot_state(self.copy_id)

    def history(self):
        with self.fixture.database.connect() as connection:
            return [tuple(row) for row in connection.execute(
                "SELECT sheet,row,col,action,revision FROM pp_history WHERE copy_id=? ORDER BY id",
                (self.copy_id,)).fetchall()]

    def test_payload_lists_excel_slicers_with_the_template_selection(self):
        slicers = self.excel_slicers()

        self.assertEqual(len(slicers), 4)
        self.assertEqual({item["engine"] for item in slicers}, {"excel"})
        ten = next(slicer for slicer in slicers if slicer["cache"] == self.cache)
        self.assertEqual([item["selected"] for item in ten["items"]], [True] + [False] * 9)

    def test_selection_commits_one_slicer_revision_and_is_stored_for_the_copy(self):
        response = self.press()

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json(), {"messages": [], "changed": 0, "input_revision": 1,
                                           "selected": [self.pick]})
        request = self.pool.requests[0]
        self.assertEqual((request.cache, request.selected, request.revision, request.base_revision),
                         (self.cache, [self.pick], 1, 0))
        self.assertEqual(request.pivot_state, {"slicers": {}, "refreshed": False})
        self.assertEqual(self.pivot_state(), {"slicers": {self.cache: [self.pick]}, "refreshed": False})
        self.assertEqual(self.history(), [])  # pivot output only: no cell contents to log
        ten = next(item for item in self.excel_slicers() if item["cache"] == self.cache)["items"]
        self.assertEqual([item["caption"] for item in ten if item["selected"]], [self.pick])

    def test_contents_written_with_a_selection_are_logged_under_the_slicer_action(self):
        original = self.pool.run_slicer

        def with_contents(request):
            return {**original(request), "contents": [(0, 9, 0, "Hi")]}

        self.pool.run_slicer = with_contents
        self.press()

        self.assertEqual(self.history(), [(0, 9, 0, "slicer", 1)])

    def test_same_selection_again_changes_nothing_and_runs_nothing(self):
        self.press()

        response = self.press(op_id="slice-2")

        self.assertEqual(response.json(), {"messages": [], "changed": 0, "input_revision": 1,
                                           "selected": [self.pick]})
        self.assertEqual(len(self.pool.requests), 1)

    def test_repeated_op_id_returns_the_stored_ack_without_running_again(self):
        first, second = self.press(), self.press()

        self.assertEqual((first.status_code, second.status_code), (200, 200))
        self.assertEqual(second.json(), first.json())
        self.assertEqual(len(self.pool.requests), 1)

    def test_stored_selection_reaches_the_payload_of_every_viewer_and_forks(self):
        self.press()
        viewer, viewer_csrf, _ = self.fixture.signup("slicer-viewer@example.test", "Viewer")
        self.client.put(f"/api/copies/{self.copy_id}/shares",
                        {"email": "slicer-viewer@example.test", "role": "viewer"},
                        {"x-csrf-token": self.csrf})
        viewer_slicer = next(item for item in self.excel_slicers(self.workbook(viewer))
                             if item["cache"] == self.cache)

        fork = self.client.post(f"/api/copies/{self.copy_id}/fork", {"name": "Fork"},
                                {"x-csrf-token": self.csrf})
        fork_id = fork.json()["copy"]["id"]
        fork_selection = self.fixture.app.state.services.workbooks.pivot_state(fork_id)

        self.assertEqual([item["caption"] for item in viewer_slicer["items"] if item["selected"]], [self.pick])
        self.assertEqual(fork_selection, {"slicers": {self.cache: [self.pick]}, "refreshed": False})

    def test_calculations_after_a_selection_carry_the_stored_selection_for_a_fresh_engine(self):
        self.press()

        self.client.post(f"/api/copies/{self.copy_id}/fork", {"name": "Baseline"}, {"x-csrf-token": self.csrf})

        self.assertEqual(self.pool.calc[-1].pivot_state, {"slicers": {self.cache: [self.pick]}, "refreshed": False})

    def test_unknown_cache_is_404_and_unknown_item_or_empty_selection_is_422(self):
        self.assertEqual(self.press({"op_id": "x1", "cache": "Nope", "selected": [self.pick]}).status_code, 404)
        unknown = self.press({"op_id": "x2", "cache": self.cache, "selected": ["atlantis"]})
        self.assertEqual((unknown.status_code, unknown.json()["error"]), (422, "unknown_item"))
        empty = self.press({"op_id": "x3", "cache": self.cache, "selected": []})
        self.assertEqual((empty.status_code, empty.json()["error"]), (422, "empty_selection"))
        not_a_list = self.press({"op_id": "x4", "cache": self.cache, "selected": self.pick})
        self.assertEqual(not_a_list.status_code, 422)
        self.assertEqual(self.pool.requests, [])

    def test_mac_values_only_and_missing_excel_are_409_slicers_unavailable(self):
        for engine in (None, FakeSlicerPool()):
            with self.subTest(engine=engine):
                self.fixture.app.state.services.engine_pool = engine
                if engine is not None:
                    engine.kind = "formualizer"
                response = self.press(op_id=f"na-{id(engine)}")
                self.assertEqual((response.status_code, response.json()["error"]),
                                 (409, "slicers_unavailable"))
                self.assertEqual(response.json()["message"],
                                 SLICER_MESSAGE)

    def test_viewer_is_forbidden_and_csrf_is_required(self):
        viewer, viewer_csrf, _ = self.fixture.signup("slicer-spy@example.test", "Viewer")
        self.client.put(f"/api/copies/{self.copy_id}/shares",
                        {"email": "slicer-spy@example.test", "role": "viewer"}, {"x-csrf-token": self.csrf})

        response = self.press(client=viewer, csrf=viewer_csrf, op_id="v1")

        self.assertEqual((response.status_code, response.json()["error"]), (403, "forbidden"))
        self.assertEqual(self.client.post(f"/api/copies/{self.copy_id}/slicers",
                                          {"op_id": "c1", "cache": self.cache, "selected": [self.pick]}).json()["error"],
                         "csrf")

    def test_engine_failure_is_502_and_commits_nothing(self):
        self.pool.error = EngineError("boom")

        response = self.press(op_id="fail-1")

        self.assertEqual((response.status_code, response.json()["error"]), (502, "macro_failed"))
        self.assertEqual(self.pivot_state(), {"slicers": {}, "refreshed": False})

    def test_slicer_past_the_time_limit_is_502_and_commits_nothing(self):
        self.pool.error = MacroTimeout("slow")

        response = self.press(op_id="slow-1")

        self.assertEqual((response.status_code, response.json()["error"]), (502, "macro_failed"))
        self.assertEqual(self.pivot_state(), {"slicers": {}, "refreshed": False})
