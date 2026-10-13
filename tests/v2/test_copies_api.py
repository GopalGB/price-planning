import threading
import unittest
from unittest import mock

from tests.v2.api_support import APIFixture
from tests.v2.ws_client import WebSocketClient


class CopyAPIContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture()
        self.owner, self.owner_csrf, self.owner_user = self.fixture.signup("copy-owner@example.test", "Copy Owner")
        self.member, self.member_csrf, self.member_user = self.fixture.signup("copy-member@example.test", "Copy Member")
        category = self.owner.post("/api/categories", {"name": "Category A"},
                                   {"x-csrf-token": self.owner_csrf}).json()["category"]
        self.category_id = category["id"]
        self.owner.put(f"/api/categories/{self.category_id}/members",
                       {"email": "copy-member@example.test", "role": "member"},
                       {"x-csrf-token": self.owner_csrf})

    def tearDown(self):
        self.fixture.close()

    def test_category_visibility_named_role_precedence_and_private_invisibility(self):
        """Copy listing must calculate the highest role and hide private copies."""
        created = self.owner.post("/api/copies", {
            "category_id": self.category_id, "name": "Shared Plan",
            "visibility": "category", "default_role": "viewer",
        }, {"x-csrf-token": self.owner_csrf})
        self.assertEqual(created.status_code, 201)
        copy_id = created.json()["copy"]["id"]
        listed = self.member.get(f"/api/copies?category={self.category_id}").json()["copies"]
        self.assertEqual(listed[0]["my_role"], "viewer")

        shared = self.owner.put(f"/api/copies/{copy_id}/shares",
                                {"email": "copy-member@example.test", "role": "editor"},
                                {"x-csrf-token": self.owner_csrf})
        self.assertEqual(shared.status_code, 200)
        self.assertEqual(self.member.get("/api/copies").json()["copies"][0]["my_role"], "editor")

        private = self.owner.patch(f"/api/copies/{copy_id}", {"visibility": "private"},
                                   {"x-csrf-token": self.owner_csrf})
        self.assertEqual(private.status_code, 200)
        self.owner.put(f"/api/copies/{copy_id}/shares",
                       {"email": "copy-member@example.test", "role": None},
                       {"x-csrf-token": self.owner_csrf})
        self.assertEqual(self.member.get("/api/copies").json()["copies"], [])

    def test_listing_reports_authorized_live_team_and_history_stats(self):
        """Stale, revoked, and duplicate tabs must not inflate the live team."""
        copy = self.owner.post("/api/copies", {
            "category_id": self.category_id, "name": "Team Plan",
            "visibility": "category", "default_role": "editor",
        }, {"x-csrf-token": self.owner_csrf}).json()["copy"]

        def connect(client):
            socket = WebSocketClient(
                self.fixture.app, f"/api/copies/{copy['id']}/ws", client.cookies
            )
            self.addCleanup(socket.close)
            self.assertTrue(socket.accepted)
            socket.send_json({"type": "hello", "last_revision": 0})
            while socket.receive_json()["type"] != "calc_state":
                pass
            return socket

        owner_first, owner_second = connect(self.owner), connect(self.owner)
        member = connect(self.member)
        listed = self.owner.get(f"/api/copies?category={self.category_id}").json()["copies"][0]
        self.assertEqual(
            [(person["id"], person["name"]) for person in listed["occupants"]],
            [(self.member_user["id"], "Copy Member"), (self.owner_user["id"], "Copy Owner")],
        )
        self.assertEqual((listed["edit_count"], listed["last_edit"]), (0, None))

        edited = self.member.post(f"/api/copies/{copy['id']}/ops", {
            "op_id": "team-edit", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        }, {"x-csrf-token": self.member_csrf})
        self.assertEqual(edited.status_code, 200)
        listed = self.owner.get(f"/api/copies?category={self.category_id}").json()["copies"][0]
        self.assertEqual(listed["edit_count"], 1)
        self.assertEqual(listed["last_edit"]["by"], {
            "id": self.member_user["id"], "name": "Copy Member",
        })

        owner_first.close()
        listed = self.owner.get(f"/api/copies?category={self.category_id}").json()["copies"][0]
        self.assertEqual({person["id"] for person in listed["occupants"]}, {
            self.member_user["id"], self.owner_user["id"],
        })
        member.close()
        listed = self.owner.get(f"/api/copies?category={self.category_id}").json()["copies"][0]
        self.assertEqual([person["id"] for person in listed["occupants"]], [self.owner_user["id"]])

        revoked = connect(self.member)
        response = self.owner.put(f"/api/categories/{self.category_id}/members", {
            "email": "copy-member@example.test", "role": None,
        }, {"x-csrf-token": self.owner_csrf})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(revoked.wait_closed(timeout=1), 4403)
        listed = self.owner.get(f"/api/copies?category={self.category_id}").json()["copies"][0]
        self.assertEqual([person["id"] for person in listed["occupants"]], [self.owner_user["id"]])
        owner_second.close()

    def test_unfiltered_listing_keeps_named_private_share_across_categories_and_revokes_it(self):
        """A private named share must list without leaking unrelated private copies."""
        outsider, outsider_csrf, _ = self.fixture.signup("copy-outsider@example.test", "Copy Outsider")
        other_category = outsider.post("/api/categories", {"name": "Category B"},
                                      {"x-csrf-token": outsider_csrf}).json()["category"]
        own = outsider.post("/api/copies", {
            "category_id": other_category["id"], "name": "Outsider Plan",
            "visibility": "private", "default_role": "viewer",
        }, {"x-csrf-token": outsider_csrf}).json()["copy"]
        shared = self.owner.post("/api/copies", {
            "category_id": self.category_id, "name": "Named Private Share",
            "visibility": "private", "default_role": "viewer",
        }, {"x-csrf-token": self.owner_csrf}).json()["copy"]
        hidden = self.owner.post("/api/copies", {
            "category_id": self.category_id, "name": "Hidden Private Control",
            "visibility": "private", "default_role": "viewer",
        }, {"x-csrf-token": self.owner_csrf}).json()["copy"]
        granted = self.owner.put(f"/api/copies/{shared['id']}/shares", {
            "email": "copy-outsider@example.test", "role": "editor",
        }, {"x-csrf-token": self.owner_csrf})
        self.assertEqual(granted.status_code, 200)

        unfiltered = outsider.get("/api/copies").json()["copies"]
        self.assertEqual(
            {(item["id"], item["category_id"], item["my_role"]) for item in unfiltered},
            {(own["id"], other_category["id"], "owner"), (shared["id"], self.category_id, "editor")},
        )
        self.assertNotIn(hidden["id"], {item["id"] for item in unfiltered})
        self.assertEqual([item["id"] for item in outsider.get(
            f"/api/copies?category={self.category_id}").json()["copies"]], [shared["id"]])
        self.assertEqual([item["id"] for item in outsider.get(
            f"/api/copies?category={other_category['id']}").json()["copies"]], [own["id"]])

        revoked = self.owner.put(f"/api/copies/{shared['id']}/shares", {
            "email": "copy-outsider@example.test", "role": None,
        }, {"x-csrf-token": self.owner_csrf})
        self.assertEqual(revoked.status_code, 200)
        self.assertEqual([item["id"] for item in outsider.get("/api/copies").json()["copies"]], [own["id"]])

    def test_malformed_control_fields_return_json_422(self):
        """Unhashable role and visibility values must never escape as TypeError 500."""
        created = self.owner.post("/api/copies", {
            "category_id": self.category_id, "name": "Boundary Plan",
            "visibility": "private", "default_role": "viewer",
        }, {"x-csrf-token": self.owner_csrf}).json()["copy"]
        copy_id = created["id"]
        cases = []
        for malformed in ([], {}):
            cases.extend([
                ("PUT", f"/api/categories/{self.category_id}/members",
                 {"email": "copy-member@example.test", "role": malformed}),
                ("POST", "/api/copies", {"category_id": self.category_id, "name": "Bad",
                                           "visibility": malformed, "default_role": "viewer"}),
                ("POST", "/api/copies", {"category_id": self.category_id, "name": "Bad",
                                           "visibility": "private", "default_role": malformed}),
                ("PATCH", f"/api/copies/{copy_id}", {"visibility": malformed}),
                ("PATCH", f"/api/copies/{copy_id}", {"default_role": malformed}),
                ("PUT", f"/api/copies/{copy_id}/shares",
                 {"email": "copy-member@example.test", "role": malformed}),
            ])
        for method, endpoint, body in cases:
            with self.subTest(method=method, endpoint=endpoint, body=body):
                response = self.owner.request(method, endpoint, body,
                                              {"x-csrf-token": self.owner_csrf})
                self.assertEqual((response.status_code, response.json()["error"]), (422, "invalid"))

    def test_fork_freezes_parent_overrides_and_provenance(self):
        """Later parent edits must not leak into a fork."""
        from pp.ops import OperationService

        created = self.owner.post("/api/copies", {
            "category_id": self.category_id, "name": "Parent",
            "visibility": "category", "default_role": "editor",
        }, {"x-csrf-token": self.owner_csrf}).json()["copy"]
        service = OperationService(self.fixture.database, self.fixture.snapshot)
        service.apply(created["id"], self.owner_user["id"], {
            "op_id": "before-fork", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        })
        fork = self.member.post(f"/api/copies/{created['id']}/fork", {"name": "My Fork"},
                                {"x-csrf-token": self.member_csrf})
        self.assertEqual(fork.status_code, 201)
        fork_copy = fork.json()["copy"]
        service.apply(created["id"], self.owner_user["id"], {
            "op_id": "after-fork", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 15, "expected_version": 1}],
        })
        with self.fixture.database.connect() as connection:
            copied = connection.execute(
                "SELECT value_json,origin_copy_id,origin_revision FROM pp_overrides WHERE copy_id=?",
                (fork_copy["id"],),
            ).fetchone()
        self.assertEqual((copied["value_json"], copied["origin_copy_id"], copied["origin_revision"]),
                         ("12", created["id"], 1))
        self.assertEqual(fork_copy["parent"], {"id": created["id"], "name": "Parent", "revision": 1})

    def test_fork_preserves_attribution_history_and_clears_origin_after_local_ops(self):
        """Restamping inherited edits or retaining inherited markers after local work must fail."""
        alice, alice_csrf, alice_user = self.fixture.signup("alice@example.test", "Alice")
        bob, bob_csrf, _ = self.fixture.signup("bob@example.test", "Bob")
        category = alice.post("/api/categories", {"name": "Alice Plans"}, {
            "x-csrf-token": alice_csrf,
        }).json()["category"]
        alice.put(f"/api/categories/{category['id']}/members", {
            "email": "bob@example.test", "role": "member",
        }, {"x-csrf-token": alice_csrf})
        parent = alice.post("/api/copies", {
            "category_id": category["id"], "name": "Alice Parent",
            "visibility": "category", "default_role": "editor",
        }, {"x-csrf-token": alice_csrf}).json()["copy"]
        edited = alice.post(f"/api/copies/{parent['id']}/ops", {
            "op_id": "alice-before-fork", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        }, {"x-csrf-token": alice_csrf})
        self.assertEqual(edited.status_code, 200)

        def override(copy_id):
            with self.fixture.database.connect() as connection:
                return dict(connection.execute(
                    "SELECT * FROM pp_overrides WHERE copy_id=? AND sheet=0 AND row=1 AND col=1",
                    (copy_id,),
                ).fetchone())

        original = override(parent["id"])
        fork = bob.post(f"/api/copies/{parent['id']}/fork", {"name": "Bob Fork"}, {
            "x-csrf-token": bob_csrf,
        }).json()["copy"]
        inherited = override(fork["id"])
        row = bob.get(f"/api/copies/{fork['id']}/sheets/0/rows?start=1&count=1").json()["rows"]["1"]["1"]
        self.assertEqual((inherited["updated_by"], inherited["updated_at"]),
                         (original["updated_by"], original["updated_at"]))
        self.assertEqual((row["by"], row["at"]),
                         ({"id": alice_user["id"], "name": "Alice"}, original["updated_at"]))

        history_url = f"/api/copies/{fork['id']}/history?sheet=0&r=1&c=1"
        inherited_history = bob.get(history_url).json()["entries"]
        self.assertEqual([(entry["op_id"], entry["by"]["name"], entry.get("origin"))
                          for entry in inherited_history], [
            ("alice-before-fork", "Alice", {
                "copy_id": parent["id"], "name": "Alice Parent", "revision": 1,
            }),
        ])
        alice.post(f"/api/copies/{parent['id']}/ops", {
            "op_id": "alice-after-fork", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 15, "expected_version": 1}],
        }, {"x-csrf-token": alice_csrf})
        self.assertEqual(bob.get(history_url).json()["entries"], inherited_history)

        local = bob.post(f"/api/copies/{fork['id']}/ops", {
            "op_id": "bob-local", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 14, "expected_version": 1}],
        }, {"x-csrf-token": bob_csrf})
        self.assertEqual(local.status_code, 200)
        self.assertEqual((override(fork["id"])["origin_copy_id"],
                          override(fork["id"])["origin_revision"]), (None, None))
        undone = bob.post(f"/api/copies/{fork['id']}/ops", {
            "op_id": "bob-undo", "action": "undo",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 2}],
        }, {"x-csrf-token": bob_csrf})
        self.assertEqual(undone.status_code, 200)
        self.assertEqual((override(fork["id"])["origin_copy_id"],
                          override(fork["id"])["origin_revision"]), (None, None))
        history = bob.get(history_url).json()["entries"]
        self.assertEqual([(entry["op_id"], entry["by"]["name"], entry.get("origin")) for entry in history], [
            ("bob-undo", "Bob", None), ("bob-local", "Bob", None),
            ("alice-before-fork", "Alice", {
                "copy_id": parent["id"], "name": "Alice Parent", "revision": 1,
            }),
        ])

    def test_copy_service_creates_durable_pending_revision_zero_without_pool(self):
        """A copy without a calculator must not claim its blank formula cache is current."""
        from pp.copies import CopyService

        copy = CopyService(self.fixture.database, self.fixture.snapshot.hash).create(
            self.owner_user["id"], {
                "category_id": self.category_id, "name": "Pending Plan",
                "visibility": "private", "default_role": "viewer",
            },
        )

        with self.fixture.database.connect() as connection:
            job = connection.execute(
                "SELECT revision,status FROM pp_calc_jobs WHERE copy_id=?", (copy["id"],)
            ).fetchone()
        self.assertEqual((copy["input_revision"], copy["calculated_revision"]), (0, -1))
        self.assertEqual((job["revision"], job["status"]), (0, "pending"))

    def test_create_and_fork_calculate_revision_zero_before_success(self):
        """Successful pooled create and fork responses must expose calculated formulas."""
        created = self.owner.post("/api/copies", {
            "category_id": self.category_id, "name": "Calculated Parent",
            "visibility": "category", "default_role": "editor",
        }, {"x-csrf-token": self.owner_csrf})
        self.assertEqual(created.status_code, 201)
        parent = created.json()["copy"]
        baseline = self.owner.get(f"/api/copies/{parent['id']}/sheets/2/cell?r=1&c=1")
        self.assertEqual((parent["input_revision"], parent["calculated_revision"]), (0, 0))
        self.assertEqual(baseline.json()["v"], 20)

        edited = self.owner.post(f"/api/copies/{parent['id']}/ops", {
            "op_id": "fork-baseline", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        }, {"x-csrf-token": self.owner_csrf})
        self.assertEqual(edited.status_code, 200)
        forked = self.member.post(f"/api/copies/{parent['id']}/fork", {"name": "Calculated Fork"},
                                  {"x-csrf-token": self.member_csrf})
        self.assertEqual(forked.status_code, 201)
        fork = forked.json()["copy"]
        inherited = self.member.get(f"/api/copies/{fork['id']}/sheets/2/cell?r=1&c=1")
        self.assertEqual((fork["input_revision"], fork["calculated_revision"]), (0, 0))
        self.assertEqual(inherited.json()["v"], 24)

    def test_baseline_creation_cannot_be_overtaken_by_revision_one(self):
        """A newer successful calculation must not turn copy creation into a 500."""
        entered, revision_one, release = threading.Event(), threading.Event(), threading.Event()
        original = self.fixture.pool.apply
        responses = {}

        def controlled(request):
            if request.revision == 0:
                entered.set()
                release.wait(timeout=5)
            else:
                revision_one.set()
            return original(request)

        def create_copy():
            responses["create"] = self.owner.post("/api/copies", {
                "category_id": self.category_id, "name": "Racing Baseline",
                "visibility": "category", "default_role": "editor",
            }, {"x-csrf-token": self.owner_csrf})

        with mock.patch.object(self.fixture.pool, "apply", side_effect=controlled):
            creator = threading.Thread(target=create_copy)
            creator.start()
            self.assertTrue(entered.wait(timeout=5))
            with self.fixture.database.connect() as connection:
                copy_id = connection.execute(
                    "SELECT id FROM pp_copies WHERE name='Racing Baseline'"
                ).fetchone()["id"]
            editor = threading.Thread(target=lambda: responses.setdefault(
                "edit", self.member.post(f"/api/copies/{copy_id}/ops", {
                    "op_id": "overtake", "action": "edit",
                    "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 13, "expected_version": 0}],
                }, {"x-csrf-token": self.member_csrf})
            ))
            editor.start()
            revision_one.wait(timeout=0.5)
            release.set()
            creator.join(timeout=5)
            editor.join(timeout=5)

        self.assertFalse(creator.is_alive())
        self.assertFalse(editor.is_alive())
        self.assertEqual((responses["create"].status_code, responses["edit"].status_code), (201, 200))
        copy = self.owner.get(f"/api/copies/{copy_id}/workbook").json()
        result = self.owner.get(f"/api/copies/{copy_id}/sheets/2/cell?r=1&c=1").json()
        self.assertEqual((copy["input_revision"], copy["calculated_revision"]), (1, 1))
        self.assertEqual(result["v"], 26)

    def test_live_subscriber_receives_baseline_calc_and_state(self):
        """A subscriber joining while baseline runs must leave calculating state."""
        from tests.v2.ws_client import WebSocketClient

        entered, release = threading.Event(), threading.Event()
        original = self.fixture.pool.apply
        responses = {}

        def controlled(request):
            if request.revision == 0:
                entered.set()
                release.wait(timeout=5)
            return original(request)

        def create_copy():
            responses["create"] = self.owner.post("/api/copies", {
                "category_id": self.category_id, "name": "Live Baseline",
                "visibility": "category", "default_role": "viewer",
            }, {"x-csrf-token": self.owner_csrf})

        with mock.patch.object(self.fixture.pool, "apply", side_effect=controlled):
            creator = threading.Thread(target=create_copy)
            creator.start()
            self.assertTrue(entered.wait(timeout=5))
            with self.fixture.database.connect() as connection:
                copy_id = connection.execute(
                    "SELECT id FROM pp_copies WHERE name='Live Baseline'"
                ).fetchone()["id"]
            socket = WebSocketClient(
                self.fixture.app, f"/api/copies/{copy_id}/ws", self.member.cookies
            )
            self.addCleanup(socket.close)
            self.assertTrue(socket.accepted)
            socket.send_json({"type": "hello", "last_revision": 0})
            initial = socket.receive_json()
            self.assertEqual((initial["type"], initial["calculated_revision"]), ("calc_state", -1))
            release.set()
            creator.join(timeout=5)
            calc = socket.receive_json()
            state = socket.receive_json()

        self.assertFalse(creator.is_alive())
        self.assertEqual(responses["create"].status_code, 201)
        self.assertEqual((calc["type"], calc["calculated_revision"]), ("calc", 0))
        self.assertEqual((state["type"], state["calculated_revision"]), ("calc_state", 0))

    def test_fork_revision_reconstruction_keeps_immutable_inherited_baseline(self):
        """Later fork edits must not rewrite inherited values in earlier revisions."""
        from pp.ops import OperationService

        parent = self.owner.post("/api/copies", {
            "category_id": self.category_id, "name": "Revision Parent",
            "visibility": "category", "default_role": "editor",
        }, {"x-csrf-token": self.owner_csrf}).json()["copy"]
        operations = OperationService(self.fixture.database, self.fixture.snapshot)
        operations.apply(parent["id"], self.owner_user["id"], {
            "op_id": "parent-a12", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        })
        fork = self.member.post(f"/api/copies/{parent['id']}/fork", {"name": "Revision Fork"},
                                {"x-csrf-token": self.member_csrf}).json()["copy"]
        operations.apply(fork["id"], self.member_user["id"], {
            "op_id": "fork-b7", "action": "edit",
            "cells": [{"sheet": 0, "r": 2, "c": 1, "v": 7, "expected_version": 0}],
        })
        operations.apply(fork["id"], self.member_user["id"], {
            "op_id": "fork-a99", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 99, "expected_version": 1}],
        })

        self.assertEqual(operations.revision_changes(fork["id"], 0), [(0, 1, 1, 12)])
        self.assertEqual(operations.revision_changes(fork["id"], 1),
                         [(0, 1, 1, 12), (0, 2, 1, 7)])
        self.assertEqual(operations.revision_changes(fork["id"], 2),
                         [(0, 1, 1, 99), (0, 2, 1, 7)])


if __name__ == "__main__":
    unittest.main()
