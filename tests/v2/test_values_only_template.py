"""A values-only template shows Excel's saved results and saves typed values without recalculating."""

import json
import unittest
from unittest import mock

from tests.v2.api_support import APIFixture


class ValuesOnlyTemplateTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture()
        self.addCleanup(self.fixture.close)

    def test_environment_turns_values_only_on(self):
        # office_run.bat sets it when the computer has no Excel, without rewriting settings.json
        from pp.config import _template_results

        with mock.patch.dict("os.environ", {"PP_TEMPLATE_RESULTS": "cached"}):
            self.assertEqual(_template_results({"template_results": "engine"}), "cached")
        with mock.patch.dict("os.environ", {"PP_TEMPLATE_RESULTS": "bogus"}), self.assertRaises(ValueError):
            _template_results({})

    def test_import_skips_the_engine_and_is_cached_separately(self):
        from pp.template.importer import import_template
        data = self.fixture.root / "values-only"
        with mock.patch("pp.template.importer._synchronize_formula_results",
                        side_effect=AssertionError("engine must not run")):
            path = import_template(self.fixture.source, data, "cached")
        self.assertEqual(json.loads((path / "meta.json").read_text())["results"], "cached")
        with mock.patch("pp.template.importer._synchronize_formula_results") as engine:
            import_template(self.fixture.source, data, "engine")
        engine.assert_called_once()

    def test_values_save_without_recalculation_and_formulas_stay_locked(self):
        client, csrf, _ = self.fixture.signup("values@example.test")
        _, copy = self.fixture.make_copy(client, csrf)
        self.fixture.snapshot.results = "cached"
        meta = client.get(f"/api/copies/{copy['id']}/workbook").json()
        before = client.get(f"/api/copies/{copy['id']}/sheets/0/rows?start=0&count=8").json()
        saved = client.post(f"/api/copies/{copy['id']}/ops", {
            "op_id": "typed", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        }, {"x-csrf-token": csrf})
        locked = client.post(f"/api/copies/{copy['id']}/ops", {
            "op_id": "template-cell", "action": "edit",
            "cells": [{"sheet": 0, "r": 0, "c": 1, "v": 5, "expected_version": 0}],
        }, {"x-csrf-token": csrf})
        typed_formula = client.post(f"/api/copies/{copy['id']}/ops", {
            "op_id": "typed-formula", "action": "edit",
            "cells": [{"sheet": 0, "r": 2, "c": 1, "v": "=1+1", "expected_version": 0}],
        }, {"x-csrf-token": csrf})
        rows = client.get(f"/api/copies/{copy['id']}/sheets/0/rows?start=0&count=8").json()

        self.assertEqual(meta["calculation"], "cached")
        template = client.get("/api/me").json()["template"]
        self.assertEqual((template["sheets"], template["calculation"]), (len(meta["sheets"]), "cached"))
        self.assertEqual(saved.status_code, 200, saved.body)
        self.assertEqual(rows["rows"]["1"]["1"]["v"], 12)
        # nothing recalculates, so the calculated revision follows the typed one and the browser never waits
        self.assertEqual((rows["input_revision"], rows["calculated_revision"]), (1, 1))
        self.assertEqual(before["input_revision"], before["calculated_revision"])
        self.assertEqual((locked.status_code, locked.json()["error"]), (422, "locked"))
        self.assertEqual((typed_formula.status_code, typed_formula.json()["error"]), (422, "invalid"))
        # the download carries the typed value, so Excel recalculates with it when the file is opened
        import io
        import openpyxl
        exported = client.get(f"/api/copies/{copy['id']}/export.xlsm")
        self.assertEqual(exported.status_code, 200)
        workbook = openpyxl.load_workbook(io.BytesIO(exported.body), keep_vba=True)
        self.assertEqual(workbook.worksheets[0].cell(2, 2).value, 12)

    def test_runtime_creates_copies_without_starting_the_engine(self):
        from pp.runtime import create_runtime
        from tests.v2.asgi_client import ASGIClient
        from tests.v2.test_end_to_end import runtime_root
        temporary, root = runtime_root()
        self.addCleanup(temporary.cleanup)
        settings = json.loads((root / "settings.json").read_text())
        (root / "settings.json").write_text(json.dumps({**settings, "template_results": "cached"}))
        with mock.patch("pp.engine.pool.EnginePool.apply", side_effect=AssertionError("engine must not run")):
            runtime = create_runtime(root)
            self.addCleanup(runtime.close)
            client = ASGIClient(runtime.app)
            client.post("/api/auth/signup", {"email": "rt@example.test", "name": "RT",
                                             "password": "correct-horse-battery"})
            csrf = client.get("/api/me").json()["csrf"]
            category = client.post("/api/categories", {"name": "RT"}, {"x-csrf-token": csrf}).json()["category"]
            response = client.post("/api/copies", {"category_id": category["id"], "name": "RT",
                                                   "visibility": "private", "default_role": "viewer"},
                                   {"x-csrf-token": csrf})
        self.assertEqual(response.status_code, 201)
        copy_id = response.json()["copy"]["id"]
        edit = client.post(f"/api/copies/{copy_id}/ops", {
            "op_id": "typed", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 15, "expected_version": 0}],
        }, {"x-csrf-token": csrf})
        self.assertEqual(edit.status_code, 200, edit.body)
        # the live calc_state message must report it calculated, or the browser shows "Calculating..."
        # forever (seen in a browser 2026-10-13); the stored job stays queued so that once Excel works
        # the engine recalculates this copy instead of skipping it
        from pp.server import _calc_state
        self.assertEqual(_calc_state(runtime.app.state.services, copy_id),
                         {"input_revision": 1, "calculated_revision": 1})
        with runtime.database.connect() as connection:
            pending = connection.execute("SELECT COUNT(*) FROM pp_calc_jobs WHERE copy_id=? AND status='pending'",
                                         (copy_id,)).fetchone()[0]
        self.assertGreaterEqual(pending, 1)
        # the browser advances its loaded rows only on a calc message, so values mode sends an empty one
        import asyncio
        from pp.server import _announce
        services, sent = runtime.app.state.services, []

        async def capture(copy, message):
            sent.append(message)
        with mock.patch.object(services.hub, "broadcast", side_effect=capture):
            asyncio.run(_announce(services, copy_id, 1))
        self.assertIn({"type": "calc", "from_calculated_revision": 0, "calculated_revision": 1, "changes": []},
                      sent)
        with mock.patch("pp.engine.pool.EnginePool.apply") as engine:
            restarted = create_runtime(root)  # the queued baseline must not be recovered with the engine
        restarted.close()
        engine.assert_not_called()


if __name__ == "__main__":
    unittest.main()
