import queue
import asyncio
import threading
import unittest
from unittest import mock

from pp.db import token_hash
from tests.v2.api_support import APIFixture
from tests.v2.ws_client import WebSocketClient


class PipelineSocket:
    def __init__(self):
        self.messages = []
        self.presence_entered = threading.Event()
        self.release_presence = threading.Event()

    async def send_json(self, payload):
        if payload["type"] == "presence":
            self.presence_entered.set()
            await asyncio.to_thread(self.release_presence.wait)
        if payload["type"] in {"op", "calc"}:
            revision = payload.get("input_revision", payload.get("calculated_revision"))
            self.messages.append((payload["type"], revision))


class WebSocketContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture()
        self.owner, self.owner_csrf, self.owner_user = self.fixture.signup("ws-owner@example.test", "WS Owner")
        self.member, self.member_csrf, self.member_user = self.fixture.signup("ws-member@example.test", "WS Member")
        self.category, self.copy = self.fixture.make_copy(
            self.owner, self.owner_csrf, visibility="category", default_role="editor"
        )
        self.owner.put(f"/api/categories/{self.category['id']}/members",
                       {"email": "ws-member@example.test", "role": "member"},
                       {"x-csrf-token": self.owner_csrf})
        self.sockets = []

    def tearDown(self):
        for socket in self.sockets:
            socket.close()
        self.fixture.close()

    def connect(self, client, origin="http://testserver"):
        socket = WebSocketClient(self.fixture.app, f"/api/copies/{self.copy['id']}/ws",
                                 client.cookies, origin)
        self.sockets.append(socket)
        return socket

    def hello(self, socket, last_revision=0):
        socket.send_json({"type": "hello", "last_revision": last_revision})
        messages = []
        while not messages or messages[-1]["type"] != "calc_state":
            messages.append(socket.receive_json())
        return messages

    def identity(self, client, user):
        return {**user, "session_hash": token_hash(client.cookies["pp_session"])}

    def assert_closed(self, socket, code=4403):
        try:
            actual = socket.wait_closed(timeout=1)
        except queue.Empty:
            self.fail("WebSocket remained open")
        self.assertEqual(actual, code)

    def test_op_precedes_calc_and_reconnect_replays_both(self):
        """Broadcasting calc before its committed op must fail."""
        socket = self.connect(self.member)
        self.assertTrue(socket.accepted)
        initial = self.hello(socket)
        self.assertEqual([message["type"] for message in initial], ["calc", "calc_state"])
        self.assertEqual(initial[0]["calculated_revision"], 0)

        self.owner.post(f"/api/copies/{self.copy['id']}/ops", {
            "op_id": "ws-op", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        }, {"x-csrf-token": self.owner_csrf})
        messages = [socket.receive_json() for _ in range(3)]
        types = [message["type"] for message in messages]
        self.assertLess(types.index("op"), types.index("calc"))

        replay = self.connect(self.member)
        replayed = self.hello(replay)
        replay_types = [message["type"] for message in replayed]
        self.assertEqual(replay_types, ["calc", "op", "calc", "calc_state"])
        self.assertEqual([message["calculated_revision"] for message in replayed
                          if message["type"] == "calc"], [0, 1])

    def test_revocation_closes_only_after_all_effective_roles_are_removed(self):
        """Removing one grant must keep a socket when category access remains."""
        self.owner.put(f"/api/copies/{self.copy['id']}/shares",
                       {"email": "ws-member@example.test", "role": "viewer"},
                       {"x-csrf-token": self.owner_csrf})
        socket = self.connect(self.member)
        self.hello(socket)
        self.owner.put(f"/api/copies/{self.copy['id']}/shares",
                       {"email": "ws-member@example.test", "role": None},
                       {"x-csrf-token": self.owner_csrf})
        with self.assertRaises(queue.Empty):
            socket.outbound.get(timeout=0.2)
        self.owner.put(f"/api/categories/{self.category['id']}/members",
                       {"email": "ws-member@example.test", "role": None},
                       {"x-csrf-token": self.owner_csrf})
        self.assert_closed(socket)

    def test_origin_and_session_are_required(self):
        """A WebSocket without both trusted Origin and session must not be accepted."""
        bad_origin = self.connect(self.member, "http://evil.test")
        self.assertFalse(bad_origin.accepted)
        self.assertEqual(bad_origin.close_code, 4403)
        anonymous = WebSocketClient(self.fixture.app, f"/api/copies/{self.copy['id']}/ws")
        self.sockets.append(anonymous)
        self.assertFalse(anonymous.accepted)
        self.assertEqual(anonymous.close_code, 4403)

    def test_malformed_messages_close_with_a_protocol_error(self):
        """Invalid JSON, non-objects, and invalid hello cursors must close explicitly."""
        messages = ["{", "[]", '{"type":"hello","last_revision":"bad"}']
        for raw in messages:
            with self.subTest(raw=raw):
                socket = self.connect(self.member)
                socket.send_text(raw)
                self.assert_closed(socket, 1008)

    def test_stale_socket_does_not_block_other_broadcasts(self):
        """One disconnected client must not make a committed edit return 500."""
        from pp.live import Connection

        class BrokenSocket:
            async def send_json(self, payload):
                raise RuntimeError("closed")

        class RecordingSocket:
            def __init__(self):
                self.messages = []

            async def send_json(self, payload):
                self.messages.append(payload)

        identity = self.identity(self.member, self.member_user)
        broken = Connection(BrokenSocket(), self.copy["id"], identity)
        recording_socket = RecordingSocket()
        recording = Connection(recording_socket, self.copy["id"], identity)
        broken.ready = recording.ready = True
        hub = self.fixture.app.state.services.hub
        hub.add(broken)
        hub.add(recording)

        asyncio.run(hub.broadcast(self.copy["id"], {"type": "calc_state"}))

        self.assertEqual(recording_socket.messages, [{"type": "calc_state"}])

    def test_private_visibility_patch_immediately_revokes_category_socket(self):
        """Changing a category copy to private must close inherited live access."""
        socket = self.connect(self.member)
        self.hello(socket)

        response = self.owner.patch(f"/api/copies/{self.copy['id']}", {"visibility": "private"},
                                    {"x-csrf-token": self.owner_csrf})

        self.assertEqual(response.status_code, 200)
        self.assert_closed(socket)

    def test_stalled_socket_is_bounded_and_does_not_block_other_clients(self):
        """A non-reading client must be closed 1013 without stalling the copy pipeline."""
        from pp.live import Connection

        class StalledSocket:
            def __init__(self):
                self.close_codes = []

            async def send_json(self, payload):
                await asyncio.Event().wait()

            async def close(self, code):
                self.close_codes.append(code)
                await asyncio.Event().wait()

        class RecordingSocket:
            def __init__(self):
                self.messages = []

            async def send_json(self, payload):
                self.messages.append(payload)

        identity = self.identity(self.member, self.member_user)
        stalled_socket, recording_socket = StalledSocket(), RecordingSocket()
        stalled = Connection(stalled_socket, self.copy["id"], identity)
        recording = Connection(recording_socket, self.copy["id"], identity)
        stalled.ready = recording.ready = True
        hub = self.fixture.app.state.services.hub
        hub.add(stalled)
        hub.add(recording)
        with mock.patch("pp.live.SEND_TIMEOUT", 0.01), mock.patch("pp.live.CLOSE_TIMEOUT", 0.01):
            asyncio.run(asyncio.wait_for(
                hub.broadcast(self.copy["id"], {"type": "calc_state"}), timeout=0.1
            ))
            asyncio.run(hub.broadcast(self.copy["id"], {"type": "calc_state", "again": True}))

        self.assertEqual(stalled_socket.close_codes, [1013])
        self.assertEqual(recording_socket.messages, [
            {"type": "calc_state"}, {"type": "calc_state", "again": True},
        ])

    def test_logout_closes_only_the_socket_for_that_session(self):
        """Logging out one browser session must leave another session for the same user live."""
        first = self.connect(self.member)
        second_client = self.fixture.client()
        login = second_client.post("/api/auth/login", {
            "email": "ws-member@example.test", "password": "correct-horse-battery",
        })
        second = self.connect(second_client)

        logout = self.member.post("/api/auth/logout", headers={"x-csrf-token": self.member_csrf})

        self.assertEqual(logout.status_code, 204)
        self.assert_closed(first)
        self.assertEqual(self.hello(second)[-1]["type"], "calc_state")
        self.assertIn("csrf", login.json())

    def test_expired_session_is_rejected_on_inbound_and_outbound_traffic(self):
        """An open socket must not outlive its persisted session expiry."""
        inbound = self.connect(self.member)
        session_hash = token_hash(self.member.cookies["pp_session"])
        with self.fixture.database.transaction() as connection:
            connection.execute("UPDATE pp_sessions SET expires_at=? WHERE token_hash=?",
                               ("2000-01-01T00:00:00Z", session_hash))
        inbound.send_json({"type": "hello", "last_revision": 0})
        self.assert_closed(inbound)

        active = self.fixture.client()
        active_login = active.post("/api/auth/login", {
            "email": "ws-member@example.test", "password": "correct-horse-battery",
        })
        outbound = self.connect(active)
        self.assertEqual(self.hello(outbound)[-1]["type"], "calc_state")
        with self.fixture.database.transaction() as connection:
            connection.execute("UPDATE pp_sessions SET expires_at=? WHERE token_hash=?",
                               ("2000-01-01T00:00:00Z", token_hash(active.cookies["pp_session"])))
        asyncio.run(self.fixture.app.state.services.hub.broadcast(self.copy["id"], {"type": "calc_state"}))
        self.assert_closed(outbound)
        self.assertIn("csrf", active_login.json())

    def test_concurrent_broadcasts_are_serialized_per_copy(self):
        """A paused revision one send must not let revision two arrive first."""
        from pp.live import Connection

        class PausingSocket:
            def __init__(self):
                self.messages = []
                self.entered = threading.Event()
                self.release = threading.Event()

            async def send_json(self, payload):
                if payload["calculated_revision"] == 1:
                    self.entered.set()
                    await asyncio.to_thread(self.release.wait)
                self.messages.append(payload["calculated_revision"])

        socket = PausingSocket()
        connection = Connection(socket, self.copy["id"], self.identity(self.member, self.member_user))
        connection.ready = True
        hub = self.fixture.app.state.services.hub
        hub.add(connection)
        first = threading.Thread(target=lambda: asyncio.run(
            hub.broadcast(self.copy["id"], {"type": "calc", "calculated_revision": 1, "changes": []})
        ))
        second = threading.Thread(target=lambda: asyncio.run(
            hub.broadcast(self.copy["id"], {"type": "calc", "calculated_revision": 2, "changes": []})
        ))

        first.start()
        self.assertTrue(socket.entered.wait(timeout=2))
        second.start()
        socket.release.set()
        first.join(timeout=2)
        second.join(timeout=2)

        self.assertEqual(socket.messages, [1, 2])

    def test_reconnect_replays_calc_at_the_client_input_revision(self):
        """An input cursor must not hide a same-revision calculation completed offline."""
        response = self.owner.post(f"/api/copies/{self.copy['id']}/ops", {
            "op_id": "offline-calc", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        }, {"x-csrf-token": self.owner_csrf})
        self.assertEqual(response.status_code, 200)
        reconnect = self.connect(self.member)

        reconnect.send_json({"type": "hello", "last_revision": 1})

        calc = reconnect.receive_json()
        self.assertEqual((calc["type"], calc["calculated_revision"]), ("calc", 1))
        self.assertEqual(reconnect.receive_json()["type"], "calc_state")

    def test_cursor_zero_replays_baseline_calc_missed_after_pending_fetch(self):
        """A pending baseline fetch followed by hello(0) must recover missed values."""
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
                "category_id": self.category["id"], "name": "Missed Baseline",
                "visibility": "category", "default_role": "viewer",
            }, {"x-csrf-token": self.owner_csrf})

        with mock.patch.object(self.fixture.pool, "apply", side_effect=controlled):
            creator = threading.Thread(target=create_copy)
            creator.start()
            try:
                self.assertTrue(entered.wait(timeout=5))
                with self.fixture.database.connect() as connection:
                    copy_id = connection.execute(
                        "SELECT id FROM pp_copies WHERE name='Missed Baseline'"
                    ).fetchone()["id"]
                rows = self.member.get(
                    f"/api/copies/{copy_id}/sheets/2/rows?start=1&count=1"
                ).json()
                self.assertEqual((rows["input_revision"], rows["calculated_revision"]), (0, -1))
            finally:
                release.set()
            creator.join(timeout=5)

        self.assertFalse(creator.is_alive())
        self.assertEqual(responses["create"].status_code, 201)
        socket = WebSocketClient(self.fixture.app, f"/api/copies/{copy_id}/ws", self.member.cookies)
        self.sockets.append(socket)
        socket.send_json({"type": "hello", "last_revision": 0})
        calc = socket.receive_json()
        self.assertEqual((calc["type"], calc["calculated_revision"]), ("calc", 0))
        values = {(item["sheet"], item["r"], item["c"]): item["v"] for item in calc["changes"]}
        self.assertEqual(values[(2, 1, 1)], 20)
        state = socket.receive_json()
        self.assertEqual((state["type"], state["calculated_revision"]), ("calc_state", 0))

    def test_cursor_zero_resyncs_if_completed_baseline_event_is_missing(self):
        """A completed baseline without its calc event must not claim replay completeness."""
        with self.fixture.database.transaction() as connection:
            connection.execute(
                "DELETE FROM pp_events WHERE copy_id=? AND event_type='calc' AND input_revision=0",
                (self.copy["id"],),
            )
        socket = self.connect(self.member)

        messages = self.hello(socket)

        self.assertEqual([message["type"] for message in messages], ["resync", "calc_state"])
        self.assertEqual(messages[-1]["calculated_revision"], 0)

    def test_simultaneous_edits_publish_complete_revision_pipelines_in_order(self):
        """Locking only individual broadcasts must let revision two overtake revision one."""
        from pp.live import Connection

        socket = PipelineSocket()
        hub = self.fixture.app.state.services.hub
        connection = Connection(socket, self.copy["id"], self.identity(self.owner, self.owner_user))
        connection.ready = True
        hub.add(connection)
        original = hub.broadcast

        async def scenario():
            first_paused, allow_first = asyncio.Event(), asyncio.Event()

            async def delayed(copy_id, payload):
                if payload.get("type") == "op" and payload.get("input_revision") == 1:
                    first_paused.set()
                    await allow_first.wait()
                await original(copy_id, payload)

            hub.broadcast = delayed
            presence = asyncio.create_task(original(self.copy["id"], {"type": "presence", "users": []}))
            self.assertTrue(await asyncio.to_thread(socket.presence_entered.wait, 2))
            first = asyncio.create_task(self._async_edit("pipeline-1", 12, 0))
            await asyncio.wait_for(first_paused.wait(), 2)
            second = asyncio.create_task(self._async_edit("pipeline-2", 13, 1))
            await asyncio.sleep(0.05)
            socket.release_presence.set()
            await asyncio.sleep(0.05)
            allow_first.set()
            return await asyncio.gather(first, second, presence)

        try:
            first, second, _ = asyncio.run(scenario())
        finally:
            hub.broadcast = original
            hub.discard(connection)
        self.assertEqual((first.status_code, second.status_code), (200, 200))
        self.assertEqual(socket.messages, [("op", 1), ("calc", 1), ("op", 2), ("calc", 2)])

    async def _async_edit(self, op_id, value, version):
        return await self.owner._request("POST", f"/api/copies/{self.copy['id']}/ops", {
            "op_id": op_id, "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": value, "expected_version": version}],
        }, {"x-csrf-token": self.owner_csrf})


if __name__ == "__main__":
    unittest.main()
