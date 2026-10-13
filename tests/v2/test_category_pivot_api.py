import json
import unittest

from ops.fixtures import make_dashboard
from tests.v2.api_support import APIFixture


COPY_SLICER = "__pp_copy__"


def api_dashboard(path):
    make_dashboard(path, row_count=120)


def total(body, field=9):
    return next(
        cell["v"] for cell in body["cells"]
        if cell["row_kind"] == "grand_total" and cell.get("data_field") == field
    )


class CategoryPivotAPIContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(api_dashboard)
        self.owner, self.owner_csrf, _ = self.fixture.signup(
            "category-owner@example.test", "Category Owner"
        )
        self.viewer, self.viewer_csrf, self.viewer_user = self.fixture.signup(
            "category-viewer@example.test", "Category Viewer"
        )
        self.hidden_owner, self.hidden_csrf, _ = self.fixture.signup(
            "hidden-owner@example.test", "Hidden Owner"
        )
        self.category, self.first = self.fixture.make_copy(
            self.owner, self.owner_csrf, name="Duplicate Plan", visibility="category"
        )
        self.owner.put(
            f"/api/categories/{self.category['id']}/members",
            {"email": "category-viewer@example.test", "role": "member"},
            {"x-csrf-token": self.owner_csrf},
        )
        second = self.owner.post(
            "/api/copies",
            {
                "category_id": self.category["id"],
                "name": "Duplicate Plan",
                "visibility": "category",
                "default_role": "viewer",
            },
            {"x-csrf-token": self.owner_csrf},
        )
        self.assertEqual(second.status_code, 201, second.body)
        self.second = second.json()["copy"]
        self.owner.put(
            f"/api/categories/{self.category['id']}/members",
            {"email": "hidden-owner@example.test", "role": "member"},
            {"x-csrf-token": self.owner_csrf},
        )
        hidden = self.hidden_owner.post(
            "/api/copies",
            {
                "category_id": self.category["id"],
                "name": "Private Hidden Plan",
                "visibility": "private",
                "default_role": "viewer",
            },
            {"x-csrf-token": self.hidden_csrf},
        )
        self.assertEqual(hidden.status_code, 201, hidden.body)
        self.hidden = hidden.json()["copy"]
        self.prefix = (
            f"/api/copies/{self.first['id']}/sheets/0/pivots/DashboardPivot/category"
        )

    def tearDown(self):
        self.fixture.close()

    def test_union_uses_current_formula_results_and_stable_duplicate_safe_copy_ids(self):
        edit = self.owner.post(
            f"/api/copies/{self.second['id']}/ops",
            {
                "op_id": "category-formula-source",
                "action": "edit",
                "cells": [
                    {"sheet": 0, "r": 0, "c": 7, "v": 5, "expected_version": 0}
                ],
            },
            {"x-csrf-token": self.owner_csrf},
        )
        self.assertEqual(edit.status_code, 200, edit.body)
        one = self.viewer.post(
            f"/api/copies/{self.first['id']}/sheets/0/pivots/DashboardPivot/refresh", {}
        ).json()
        two = self.viewer.post(
            f"/api/copies/{self.second['id']}/sheets/0/pivots/DashboardPivot/refresh", {}
        ).json()

        response = self.viewer.post(f"{self.prefix}/refresh", {})

        self.assertEqual(response.status_code, 200, response.body)
        body = response.json()
        self.assertEqual(total(body), total(one) + total(two))
        self.assertEqual(body["pivot"]["scope"], "category")
        self.assertEqual(
            {item["copy_id"] for item in body["pivot"]["revisions"]},
            {self.first["id"], self.second["id"]},
        )
        encoded = json.dumps(body)
        self.assertNotIn(self.hidden["id"], encoded)
        self.assertNotIn("Private Hidden Plan", encoded)
        copy_slicer = next(item for item in body["slicers"] if item["name"] == COPY_SLICER)
        self.assertEqual([item["caption"] for item in copy_slicer["items"]], [
            "Duplicate Plan", "Duplicate Plan"
        ])
        self.assertEqual(
            {item["selector"]["value"] for item in copy_slicer["items"]},
            {self.first["id"], self.second["id"]},
        )
        self.assertEqual(copy_slicer["anchor"]["from"]["row"], 16)
        self.assertEqual(copy_slicer["anchor"]["from"]["column"], 8)
        self.assertEqual(copy_slicer["anchor"]["to"]["column"], 11)

        selected = self.viewer.post(f"{self.prefix}/refresh", {
            "slicers": [{
                "name": COPY_SLICER,
                "field": -1,
                "items": [{"kind": "string", "value": self.second["id"]}],
            }]
        })
        self.assertEqual(selected.status_code, 200, selected.body)
        self.assertEqual(total(selected.json()), total(two))

    def test_refresh_reauthorizes_and_names_pending_and_incompatible_states_without_hidden_metadata(self):
        first = self.viewer.post(f"{self.prefix}/refresh", {})
        self.assertEqual(first.status_code, 200, first.body)
        with self.fixture.database.transaction() as connection:
            connection.execute(
                "UPDATE pp_copies SET input_revision=1,calculated_revision=0 WHERE id=?",
                (self.second["id"],),
            )
        pending = self.viewer.post(f"{self.prefix}/refresh", {})
        self.assertEqual(pending.status_code, 409, pending.body)
        self.assertEqual(pending.json()["error"], "category_calculation_pending")
        self.assertNotIn("Private Hidden Plan", pending.body.decode())

        with self.fixture.database.transaction() as connection:
            connection.execute(
                "UPDATE pp_copies SET input_revision=0,calculated_revision=0 WHERE id=?",
                (self.second["id"],),
            )
            connection.execute(
                "INSERT INTO pp_templates(id,name,source_path,snapshot_path,created_at) "
                "SELECT 'other-template',name,source_path,snapshot_path,created_at "
                "FROM pp_templates LIMIT 1"
            )
            connection.execute(
                "UPDATE pp_copies SET template_id='other-template' WHERE id=?",
                (self.second["id"],),
            )
        incompatible = self.viewer.post(f"{self.prefix}/refresh", {})
        self.assertEqual(incompatible.status_code, 409, incompatible.body)
        self.assertEqual(incompatible.json(), {
            "error": "incompatible_templates",
            "message": "Category copies use incompatible templates",
        })

        with self.fixture.database.transaction() as connection:
            connection.execute(
                "DELETE FROM pp_category_members WHERE category_id=? AND user_id=?",
                (self.category["id"], self.viewer_user["id"]),
            )
        revoked = self.viewer.post(f"{self.prefix}/refresh", {})
        self.assertEqual(revoked.status_code, 404, revoked.body)

    def test_copy_filter_rejects_null_unhashable_and_inaccessible_selectors(self):
        malformed = [
            {"slicers": None},
            {"slicers": [{"name": COPY_SLICER, "field": -1, "items": [
                {"kind": "string", "value": []}
            ]}]},
            {"slicers": [{"name": COPY_SLICER, "field": -1, "items": [
                {"kind": "string", "value": {}}
            ]}]},
            {"slicers": [{"name": COPY_SLICER, "field": -1, "items": [
                {"kind": "string", "value": self.hidden["id"]}
            ]}]},
        ]
        for index, body in enumerate(malformed):
            with self.subTest(index=index):
                response = self.viewer.post(f"{self.prefix}/refresh", body)
                self.assertEqual(response.status_code, 422, response.body)
                self.assertEqual(response.json()["error"], "invalid")

    def test_category_view_is_personal_category_scoped_and_separate_from_copy_view(self):
        default_copy = self.viewer.get(
            f"/api/copies/{self.first['id']}/sheets/0/pivots/DashboardPivot/view"
        ).json()
        category_view = self.viewer.get(f"{self.prefix}/view")
        self.assertEqual(category_view.status_code, 200, category_view.body)
        copy_filter = next(
            item for item in category_view.json()["slicers"] if item["name"] == COPY_SLICER
        )
        self.assertIsNone(copy_filter["items"])
        selected = {**category_view.json()}
        selected.pop("pivot")
        selected.pop("updated_at")
        copy_filter = next(item for item in selected["slicers"] if item["name"] == COPY_SLICER)
        copy_filter["items"] = [{"kind": "string", "value": self.second["id"]}]

        saved = self.viewer.put(
            f"{self.prefix}/view", selected, {"x-csrf-token": self.viewer_csrf}
        )

        self.assertEqual(saved.status_code, 200, saved.body)
        self.assertEqual(self.viewer.get(f"{self.prefix}/view").json(), saved.json())
        self.assertEqual(
            self.viewer.get(
                f"/api/copies/{self.first['id']}/sheets/0/pivots/DashboardPivot/view"
            ).json(),
            default_copy,
        )
        owner_view = self.owner.get(f"{self.prefix}/view").json()
        owner_copy_filter = next(
            item for item in owner_view["slicers"] if item["name"] == COPY_SLICER
        )
        self.assertIsNone(owner_copy_filter["items"])

    def test_revoked_private_copy_is_removed_from_cached_results_and_saved_view(self):
        private = self.owner.post(
            "/api/copies",
            {
                "category_id": self.category["id"],
                "name": "Private Shared Plan",
                "visibility": "private",
                "default_role": "viewer",
            },
            {"x-csrf-token": self.owner_csrf},
        ).json()["copy"]
        shared = self.owner.put(
            f"/api/copies/{private['id']}/shares",
            {"email": "category-viewer@example.test", "role": "viewer"},
            {"x-csrf-token": self.owner_csrf},
        )
        self.assertEqual(shared.status_code, 200, shared.body)
        cached = self.viewer.post(f"{self.prefix}/refresh", {})
        self.assertEqual(cached.status_code, 200, cached.body)
        self.assertIn(private["id"], cached.body.decode())
        view = self.viewer.get(f"{self.prefix}/view").json()
        view.pop("pivot")
        view.pop("updated_at")
        copy_filter = next(item for item in view["slicers"] if item["name"] == COPY_SLICER)
        copy_filter["items"] = [{"kind": "string", "value": private["id"]}]
        saved = self.viewer.put(
            f"{self.prefix}/view", view, {"x-csrf-token": self.viewer_csrf}
        )
        self.assertEqual(saved.status_code, 200, saved.body)

        revoked = self.owner.put(
            f"/api/copies/{private['id']}/shares",
            {"email": "category-viewer@example.test", "role": None},
            {"x-csrf-token": self.owner_csrf},
        )
        self.assertEqual(revoked.status_code, 200, revoked.body)

        refreshed = self.viewer.post(f"{self.prefix}/refresh", {})
        restored = self.viewer.get(f"{self.prefix}/view")
        self.assertEqual(refreshed.status_code, 200, refreshed.body)
        self.assertNotIn(private["id"], refreshed.body.decode())
        self.assertEqual(restored.status_code, 200, restored.body)
        self.assertNotIn(private["id"], restored.body.decode())
        sanitized = next(
            item for item in restored.json()["slicers"] if item["name"] == COPY_SLICER
        )
        self.assertIsNone(sanitized["items"])


if __name__ == "__main__":
    unittest.main()
