import unittest
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from ops.fixtures import make_dashboard
from tests.v2.api_support import APIFixture


def api_dashboard(path):
    make_dashboard(path, row_count=120)


def api_unlocked_dashboard(path):
    make_dashboard(path, row_count=120)
    with ZipFile(path) as source:
        parts = {name: source.read(name) for name in source.namelist()}
    worksheet = ET.fromstring(parts["xl/worksheets/sheet1.xml"])
    namespace = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    dashboard = worksheet.find(".//s:c[@r='A5']", namespace)
    page_filter = worksheet.find(".//s:c[@r='B1']", namespace)
    input_cell = worksheet.find(".//s:c[@r='H1']", namespace)
    if dashboard is None or page_filter is None or input_cell is None:
        raise AssertionError("Synthetic Dashboard cells are missing")
    dashboard.set("s", input_cell.get("s", "0"))
    page_filter.set("s", input_cell.get("s", "0"))
    parts["xl/worksheets/sheet1.xml"] = ET.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    with ZipFile(path, "w", ZIP_DEFLATED) as output:
        for name, content in parts.items():
            output.writestr(name, content)


class DashboardViewAPIContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(api_dashboard)
        self.client, self.csrf, self.user = self.fixture.signup(
            "dashboard-owner@example.test", "Dashboard Owner"
        )
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)
        self.path = (
            f"/api/copies/{self.copy['id']}/sheets/0/"
            "pivots/DashboardPivot/view"
        )

    def tearDown(self):
        self.fixture.close()

    def test_workbook_exposes_imported_in_grid_dashboard_metadata(self):
        response = self.client.get(f"/api/copies/{self.copy['id']}/workbook")

        self.assertEqual(response.status_code, 200, response.body)
        dashboard = response.json()["sheets"][0]
        self.assertEqual(dashboard["pivots"][0]["name"], "DashboardPivot")
        self.assertEqual(dashboard["pivots"][0]["location"]["bounds"], [4, 0, 13, 5])
        self.assertEqual(dashboard["pivots"][0]["style"]["name"], "PivotStyleLight16")
        self.assertEqual(len(dashboard["slicers"]), 4)
        self.assertEqual(dashboard["slicers"][0]["style"], "SlicerStyleLight1")
        self.assertEqual(dashboard["slicers"][0]["anchor"]["from"]["row"], 16)
        self.assertGreaterEqual(dashboard["rows"], 30)
        self.assertGreaterEqual(dashboard["cols"], 8)

    def test_default_and_saved_views_use_stable_typed_selectors(self):
        default = self.client.get(self.path)

        self.assertEqual(default.status_code, 200, default.body)
        body = default.json()
        self.assertEqual(body["pivot"], "DashboardPivot")
        self.assertIsNone(body["page"][0]["items"])
        self.assertEqual(body["slicers"][0]["items"], [
            {"kind": "string", "value": "Region 01"}
        ])

        with self.fixture.database.connect() as connection:
            revisions_before = tuple(connection.execute(
                "SELECT input_revision,calculated_revision FROM pp_copies WHERE id=?",
                (self.copy["id"],),
            ).fetchone())
            history_before = connection.execute(
                "SELECT COUNT(*) FROM pp_history WHERE copy_id=?", (self.copy["id"],)
            ).fetchone()[0]
        saved = self.client.put(self.path, {
            "page": [{"field": 1, "items": [0]}],
            "slicers": [{"name": "Slicer_Region", "field": 0, "items": [
                {"kind": "string", "value": "Region 99"}
            ]}],
            "collapsed": [],
        }, {"x-csrf-token": self.csrf})

        self.assertEqual(saved.status_code, 200, saved.body)
        self.assertEqual(len(saved.json()["page"]), 3)
        self.assertEqual(len(saved.json()["slicers"]), 4)
        self.assertEqual(saved.json()["page"][0]["items"], [
            {"kind": "string", "value": "Department 01"}
        ])
        self.assertEqual(saved.json()["slicers"][0]["items"], [
            {"kind": "string", "value": "Region 99"}
        ])
        with self.fixture.database.connect() as connection:
            row = connection.execute(
                "SELECT view_json FROM pp_dashboard_views WHERE copy_id=?",
                (self.copy["id"],),
            ).fetchone()
            revisions_after = tuple(connection.execute(
                "SELECT input_revision,calculated_revision FROM pp_copies WHERE id=?",
                (self.copy["id"],),
            ).fetchone())
            history_after = connection.execute(
                "SELECT COUNT(*) FROM pp_history WHERE copy_id=?", (self.copy["id"],)
            ).fetchone()[0]
        self.assertNotIn('"items":[0]', row["view_json"].replace(" ", ""))
        self.assertEqual(revisions_after, revisions_before)
        self.assertEqual(history_after, history_before)

    def test_stale_persisted_view_falls_back_to_current_imported_defaults(self):
        with self.fixture.database.transaction() as connection:
            connection.execute(
                "INSERT INTO pp_dashboard_views"
                "(copy_id,user_id,sheet,pivot_name,view_json,updated_at) VALUES(?,?,?,?,?,?)",
                (
                    self.copy["id"],
                    self.user["id"],
                    0,
                    "DashboardPivot",
                    '{"page":[],"slicers":[{"name":"Removed","field":0,"items":null}],'
                    '"collapsed":[]}',
                    "2026-01-01T00:00:00Z",
                ),
            )

        response = self.client.get(self.path)

        self.assertEqual(response.status_code, 200, response.body)
        self.assertEqual(response.json()["slicers"][0]["name"], "Slicer_Region")
        self.assertIsNone(response.json()["updated_at"])

    def test_views_are_user_copy_isolated_restart_safe_and_reauthorized(self):
        viewer, viewer_csrf, viewer_user = self.fixture.signup(
            "dashboard-viewer@example.test", "Dashboard Viewer"
        )
        with self.fixture.database.transaction() as connection:
            connection.execute(
                "INSERT INTO pp_copy_shares(copy_id,user_id,role) VALUES(?,?,?)",
                (self.copy["id"], viewer_user["id"], "viewer"),
            )
        selected = [{"kind": "string", "value": "Region 02"}]
        response = viewer.put(self.path, {
            "page": [],
            "slicers": [{"name": "Slicer_Region", "field": 0, "items": selected}],
            "collapsed": [],
        }, {"x-csrf-token": viewer_csrf})
        self.assertEqual(response.status_code, 200, response.body)

        owner_view = self.client.get(self.path).json()
        self.assertNotEqual(owner_view["slicers"][0]["items"], selected)
        _, other_copy = self.fixture.make_copy(self.client, self.csrf, name="Other Plan")
        other_path = (
            f"/api/copies/{other_copy['id']}/sheets/0/"
            "pivots/DashboardPivot/view"
        )
        self.assertNotEqual(self.client.get(other_path).json()["slicers"][0]["items"], selected)

        from pp.dashboard_views import DashboardViewQuery, DashboardViewService

        restarted = DashboardViewService(
            self.fixture.database, self.fixture.app.state.services.workbooks.templates
        )
        remembered = restarted.get(DashboardViewQuery(
            self.copy["id"], viewer_user["id"], 0, "DashboardPivot"
        ))
        self.assertEqual(remembered["slicers"][0]["items"], selected)

        with self.fixture.database.transaction() as connection:
            connection.execute(
                "DELETE FROM pp_copy_shares WHERE copy_id=? AND user_id=?",
                (self.copy["id"], viewer_user["id"]),
            )
        self.assertEqual(viewer.get(self.path).status_code, 404)
        self.assertEqual(
            viewer.put(self.path, response.json(), {"x-csrf-token": viewer_csrf}).status_code,
            404,
        )

    def test_update_requires_csrf_and_rejects_malformed_or_nonfinite_selectors(self):
        valid = {"page": [], "slicers": [], "collapsed": []}
        self.assertEqual(self.client.put(self.path, valid).status_code, 403)
        invalid = {
            "page": [{"field": 1, "items": [{"kind": "number", "value": float("inf")}]}],
            "slicers": [],
            "collapsed": [],
        }
        response = self.client.put(self.path, invalid, {"x-csrf-token": self.csrf})
        self.assertEqual(response.status_code, 422)
        self.assertEqual(response.json()["error"], "invalid")

        malformed = [
            {"page": [{"field": [], "items": None}], "slicers": [], "collapsed": []},
            {"page": [], "slicers": [{"name": [], "field": 0, "items": None}],
             "collapsed": []},
            {"page": [{"field": 1, "items": [{"kind": [], "value": "x"}]}],
             "slicers": [], "collapsed": []},
            {"page": [{"field": 1, "items": [
                {"kind": "number", "value": 10**400}
            ]}], "slicers": [], "collapsed": []},
            {"page": [], "slicers": [],
             "collapsed": [{"field": {}, "selector": {"kind": "string", "value": "x"}}]},
            {"page": [{"field": True, "items": None}], "slicers": [], "collapsed": []},
        ]
        for index, body in enumerate(malformed):
            with self.subTest(index=index):
                response = self.client.put(
                    self.path, body, {"x-csrf-token": self.csrf}
                )
                self.assertEqual(response.status_code, 422, response.body)
                self.assertEqual(response.json()["error"], "invalid")


class DashboardOutputLockTests(unittest.TestCase):
    def test_imported_pivot_output_excludes_an_explicitly_unlocked_cell(self):
        fixture = APIFixture(api_unlocked_dashboard)
        try:
            client, csrf, _ = fixture.signup(
                "dashboard-lock@example.test", "Dashboard Lock"
            )
            _, copy = fixture.make_copy(client, csrf)

            snapshot = fixture.snapshot
            self.assertIn("0:4:0", snapshot.input_mask)
            self.assertIn("0:0:1", snapshot.input_mask)
            self.assertTrue(snapshot.is_input(0, 0, 7))
            self.assertFalse(snapshot.is_input(0, 4, 0))
            self.assertFalse(snapshot.is_input(0, 0, 1))

            input_response = client.post(f"/api/copies/{copy['id']}/ops", {
                "op_id": "edit-unrelated-dashboard-input",
                "action": "edit",
                "cells": [{
                    "sheet": 0, "r": 0, "c": 7, "v": 3, "expected_version": 0
                }],
            }, {"x-csrf-token": csrf})

            response = client.post(f"/api/copies/{copy['id']}/ops", {
                "op_id": "edit-derived-dashboard-cell",
                "action": "edit",
                "cells": [{
                    "sheet": 0, "r": 4, "c": 0, "v": "Changed", "expected_version": 0
                }],
            }, {"x-csrf-token": csrf})
            page_response = client.post(f"/api/copies/{copy['id']}/ops", {
                "op_id": "edit-derived-page-filter-cell",
                "action": "edit",
                "cells": [{
                    "sheet": 0, "r": 0, "c": 1, "v": "Changed", "expected_version": 0
                }],
            }, {"x-csrf-token": csrf})

            self.assertEqual(input_response.status_code, 200, input_response.body)
            self.assertEqual(response.status_code, 422, response.body)
            self.assertEqual(response.json()["error"], "locked")
            self.assertEqual(page_response.status_code, 422, page_response.body)
            self.assertEqual(page_response.json()["error"], "locked")
        finally:
            fixture.close()


if __name__ == "__main__":
    unittest.main()
