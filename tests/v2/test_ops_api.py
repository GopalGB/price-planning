import unittest
from unittest import mock
import threading

from openpyxl import Workbook
from openpyxl.styles import Protection

from tests.v2.api_support import APIFixture


def dynamic_spill_fixture(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Spill"
    sheet.protection.sheet = True
    sheet["A1"] = 2
    sheet["A1"].protection = Protection(locked=False)
    sheet.column_dimensions["B"].protection = Protection(locked=False)
    sheet["B1"] = "=SEQUENCE(A1,1,10,1)"
    workbook.save(path)


class OperationAPIContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture()
        self.client, self.csrf, _ = self.fixture.signup("ops@example.test", "Ops Owner")
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)

    def tearDown(self):
        self.fixture.close()

    def test_edit_recalculates_dependent_and_rows_expose_versions(self):
        """Acknowledging before real exact-revision calculation must fail."""
        copy_id = self.copy["id"]
        response = self.client.post(f"/api/copies/{copy_id}/ops", {
            "op_id": "http-op", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        }, {"x-csrf-token": self.csrf})
        self.assertEqual(response.json(), {"op_id": "http-op", "input_revision": 1, "accepted": 1})

        inputs = self.client.get(f"/api/copies/{copy_id}/sheets/0/rows?start=1&count=1").json()
        calculated = self.client.get(f"/api/copies/{copy_id}/sheets/2/rows?start=1&count=1").json()
        self.assertEqual(inputs["rows"]["1"]["1"]["v"], 12)
        self.assertEqual(inputs["rows"]["1"]["1"]["ver"], 1)
        self.assertEqual(calculated["rows"]["1"]["1"]["v"], 24)
        self.assertEqual((calculated["input_revision"], calculated["calculated_revision"]), (1, 1))
        history = self.client.get(f"/api/copies/{copy_id}/history?sheet=0&r=1&c=1").json()["entries"]
        self.assertEqual((history[0]["op_id"], history[0]["before"], history[0]["after"]),
                         ("http-op", 10, 12))

    def test_api_error_value_recalculates_to_a_typed_formula_error(self):
        """A stored API error must not leave calculation at the prior revision."""
        copy_id = self.copy["id"]
        response = self.client.post(f"/api/copies/{copy_id}/ops", {
            "op_id": "error-input", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1,
                       "v": {"error": "#DIV/0!"}, "expected_version": 0}],
        }, {"x-csrf-token": self.csrf})
        inputs = self.client.get(f"/api/copies/{copy_id}/sheets/0/rows?start=1&count=1").json()
        calculated = self.client.get(f"/api/copies/{copy_id}/sheets/2/rows?start=1&count=1").json()

        self.assertEqual(response.status_code, 200)
        self.assertEqual(inputs["rows"]["1"]["1"]["v"], {"error": "#DIV/0!"})
        self.assertEqual(calculated["rows"]["1"]["1"]["v"], {"error": "#DIV/0!"})
        self.assertEqual((calculated["input_revision"], calculated["calculated_revision"]), (1, 1))

    def test_committed_edit_calculates_after_actor_access_is_revoked(self):
        """Post-commit calculation must not reauthorize the actor who made the edit."""
        actor, actor_csrf, _ = self.fixture.signup("revoked@example.test", "Revoked Editor")
        category_id, copy_id = self.copy["category_id"], self.copy["id"]
        self.client.put(f"/api/categories/{category_id}/members", {
            "email": "revoked@example.test", "role": "member",
        }, {"x-csrf-token": self.csrf})
        self.client.patch(f"/api/copies/{copy_id}", {
            "visibility": "category", "default_role": "editor",
        }, {"x-csrf-token": self.csrf})
        hub, original = self.fixture.app.state.services.hub, self.fixture.app.state.services.hub.broadcast
        revoked = False

        async def revoke_after_commit(target_copy, payload):
            nonlocal revoked
            if payload.get("type") == "op" and not revoked:
                revoked = True
                with self.fixture.database.transaction() as connection:
                    connection.execute(
                        "DELETE FROM pp_category_members WHERE category_id=? AND user_id="
                        "(SELECT id FROM pp_users WHERE email='revoked@example.test')",
                        (category_id,),
                    )
            await original(target_copy, payload)

        with mock.patch.object(hub, "broadcast", new=revoke_after_commit):
            response = actor.post(f"/api/copies/{copy_id}/ops", {
                "op_id": "revoke-after-commit", "action": "edit",
                "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
            }, {"x-csrf-token": actor_csrf})

        calculated = self.client.get(f"/api/copies/{copy_id}/sheets/2/rows?start=1&count=1").json()
        with self.fixture.database.connect() as connection:
            job = connection.execute(
                "SELECT status,attempts FROM pp_calc_jobs WHERE copy_id=? AND revision=1", (copy_id,)
            ).fetchone()
        self.assertEqual(response.status_code, 200)
        self.assertEqual(calculated["rows"]["1"]["1"]["v"], 24)
        self.assertEqual((job["status"], job["attempts"]), ("complete", 0))

    def test_http_conflict_locked_formula_and_idempotency_shapes(self):
        """HTTP errors must preserve atomic op error codes and details."""
        copy_id = self.copy["id"]
        payload = {"op_id": "same", "action": "edit",
                   "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}]}
        first = self.client.post(f"/api/copies/{copy_id}/ops", payload, {"x-csrf-token": self.csrf})
        retry = self.client.post(f"/api/copies/{copy_id}/ops", payload, {"x-csrf-token": self.csrf})
        self.assertEqual(retry.json(), first.json())
        reused = self.client.post(f"/api/copies/{copy_id}/ops", {**payload, "cells": [
            {"sheet": 0, "r": 1, "c": 1, "v": 13, "expected_version": 1}
        ]}, {"x-csrf-token": self.csrf})
        self.assertEqual((reused.status_code, reused.json()["error"]), (409, "op_id_reused"))
        locked = self.client.post(f"/api/copies/{copy_id}/ops", {
            "op_id": "locked", "action": "edit",
            "cells": [{"sheet": 2, "r": 1, "c": 1, "v": 5, "expected_version": 0}],
        }, {"x-csrf-token": self.csrf})
        self.assertEqual((locked.status_code, locked.json()["error"]), (422, "locked"))

    def test_malformed_action_returns_422_without_committing(self):
        """List and object actions must not escape membership validation as a 500."""
        copy_id = self.copy["id"]
        before = (
            self.fixture.count("pp_history"), self.fixture.count("pp_calc_jobs")
        )
        for index, action in enumerate(([], {})):
            with self.subTest(action=action):
                response = self.client.post(f"/api/copies/{copy_id}/ops", {
                    "op_id": f"bad-action-{index}", "action": action,
                    "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
                }, {"x-csrf-token": self.csrf})
                self.assertEqual((response.status_code, response.json()["error"]), (422, "invalid"))
        after = self.fixture.count("pp_history"), self.fixture.count("pp_calc_jobs")
        self.assertEqual(after, before)
        self.assertEqual(self.fixture.copy()["input_revision"], 0)

    def test_stale_precheck_cannot_recalculate_an_idempotent_retry(self):
        """The transaction that stores an op must decide whether calculation runs."""
        copy_id = self.copy["id"]
        payload = {"op_id": "racing-retry", "action": "edit", "cells": [
            {"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}
        ]}
        headers = {"x-csrf-token": self.csrf}
        with mock.patch.object(self.fixture.pool, "apply", wraps=self.fixture.pool.apply) as apply:
            with mock.patch("pp.server._op_exists", return_value=False, create=True):
                first = self.client.post(f"/api/copies/{copy_id}/ops", payload, headers)
                retry = self.client.post(f"/api/copies/{copy_id}/ops", payload, headers)
        self.assertEqual((first.status_code, retry.status_code), (200, 200))
        self.assertEqual(retry.json(), first.json())
        self.assertEqual(apply.call_count, 1)

    def test_calculation_persistence_runs_off_the_asgi_event_loop(self):
        """Full-result SQLite replacement must not block unrelated ASGI traffic."""
        service = self.fixture.app.state.services.calculations
        original, threads = service.complete, []

        def observed(publication):
            threads.append(threading.get_ident())
            return original(publication)

        with mock.patch.object(service, "complete", side_effect=observed):
            response = self.client.post(f"/api/copies/{self.copy['id']}/ops", {
                "op_id": "threaded-persistence", "action": "edit",
                "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
            }, {"x-csrf-token": self.csrf})

        self.assertEqual(response.status_code, 200)
        self.assertNotEqual(threads, [threading.get_ident()])


class DynamicSpillAPIContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(dynamic_spill_fixture)
        self.client, self.csrf, _ = self.fixture.signup("spill@example.test", "Spill Owner")
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)

    def tearDown(self):
        self.fixture.close()

    def test_spill_followers_recalculate_and_remain_locked_formula_results(self):
        copy_id = self.copy["id"]
        response = self.client.post(f"/api/copies/{copy_id}/ops", {
            "op_id": "resize-spill", "action": "edit",
            "cells": [{"sheet": 0, "r": 0, "c": 0, "v": 3, "expected_version": 0}],
        }, {"x-csrf-token": self.csrf})
        rows = self.client.get(f"/api/copies/{copy_id}/sheets/0/rows?start=0&count=3").json()
        workbook = self.client.get(f"/api/copies/{copy_id}/workbook").json()
        cells = [self.client.get(
            f"/api/copies/{copy_id}/sheets/0/cell?r={row}&c=1"
        ).json() for row in range(3)]
        locked = self.client.post(f"/api/copies/{copy_id}/ops", {
            "op_id": "edit-spill-follower", "action": "edit",
            "cells": [{"sheet": 0, "r": 2, "c": 1, "v": 99, "expected_version": 0}],
        }, {"x-csrf-token": self.csrf})

        self.assertEqual(response.status_code, 200)
        self.assertEqual([rows["rows"][str(row)]["1"]["v"] for row in range(3)], [10, 11, 12])
        self.assertEqual(workbook["sheets"][0]["rows"], 3)
        self.assertEqual([cell["k"] for cell in cells], ["f", "f", "f"])
        self.assertEqual([cell["formula"] for cell in cells], ["=SEQUENCE(A1,1,10,1)", None, None])
        self.assertEqual((locked.status_code, locked.json()["error"]), (422, "locked"))


if __name__ == "__main__":
    unittest.main()
