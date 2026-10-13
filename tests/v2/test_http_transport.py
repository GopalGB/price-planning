"""Exercise HTTP and WebSocket protocols through a real loopback Uvicorn server."""

import http.cookiejar
import json
from pathlib import Path
import socket
import subprocess
import sys
import tempfile
import time
import unittest
import urllib.error
import urllib.request

from websockets.exceptions import ConnectionClosed
from websockets.sync.client import connect

ROOT = Path(__file__).resolve().parents[2]
TMP = ROOT / "results-private" / "tmp"
ORIGIN = "http://testserver"


class HTTPClient:
    def __init__(self, base):
        self.base, self.csrf = base, None
        self.cookies = http.cookiejar.CookieJar()
        self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.cookies))

    def send(self, method, path, body=None):
        headers = {"Origin": ORIGIN, "Content-Type": "application/json"}
        if self.csrf:
            headers["X-CSRF-Token"] = self.csrf
        data = json.dumps(body).encode() if body is not None else None
        request = urllib.request.Request(self.base + path, data=data, headers=headers, method=method)
        try:
            response = self.opener.open(request, timeout=15)
        except urllib.error.HTTPError as error:
            response = error
        with response:
            raw = response.read()
            try:
                value = json.loads(raw) if raw else None
            except json.JSONDecodeError:
                value = {"error": "non_json"}
            return response.status, value

    def signup(self, email):
        status, _ = self.send("POST", "/api/auth/signup", {
            "email": email, "name": "Transport Tester", "password": "synthetic-password-only"
        })
        if status != 201:
            raise AssertionError(f"signup HTTP status {status}")
        status, me = self.send("GET", "/api/me")
        if status != 200:
            raise AssertionError(f"me HTTP status {status}")
        self.csrf = me["csrf"]

    def cookie_header(self):
        return "; ".join(f"{cookie.name}={cookie.value}" for cookie in self.cookies)


class HTTPTransportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        TMP.mkdir(parents=True, exist_ok=True)
        cls.temporary = tempfile.TemporaryDirectory(dir=TMP)
        cls.addClassCleanup(cls.temporary.cleanup)
        with socket.socket() as reservation:
            reservation.bind(("127.0.0.1", 0))
            cls.port = reservation.getsockname()[1]
        cls.base = f"http://127.0.0.1:{cls.port}"
        cls.log = (Path(cls.temporary.name) / "server.log").open("w")
        cls.addClassCleanup(cls.log.close)
        cls.process = subprocess.Popen([sys.executable, str(ROOT / "tests/v2/http_process.py"),
                                        str(cls.port)], cwd=ROOT, stdout=cls.log, stderr=cls.log)
        cls.addClassCleanup(cls.stop_server)
        deadline = time.monotonic() + 20
        while time.monotonic() < deadline:
            if cls.process.poll() is not None:
                raise AssertionError("synthetic Uvicorn server exited during startup")
            try:
                urllib.request.urlopen(cls.base + "/api/me", timeout=1).close()
            except urllib.error.HTTPError as error:
                if error.code == 401:
                    return
            except (OSError, urllib.error.URLError):
                time.sleep(0.05)
        raise AssertionError("synthetic Uvicorn server did not become ready")

    @classmethod
    def stop_server(cls):
        cls.process.terminate()
        try:
            cls.process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            cls.process.kill()
            cls.process.wait(timeout=5)

    def make_copy(self, client):
        status, result = client.send("POST", "/api/categories", {"name": "Network Category"})
        self.assertEqual(status, 201)
        status, result = client.send("POST", "/api/copies", {
            "category_id": result["category"]["id"], "name": "Network Copy",
            "visibility": "private", "default_role": "viewer"
        })
        self.assertEqual(status, 201)
        return result["copy"]["id"]

    def test_http_edit_recalculates_and_websocket_delivers_ordered_batches(self):
        owner = HTTPClient(self.base)
        owner.signup("transport-owner@example.test")
        copy_id = self.make_copy(owner)
        path = f"/api/copies/{copy_id}"
        operation = {"op_id": "transport-edit", "action": "edit", "cells": [
            {"sheet": 0, "r": 1, "c": 1, "v": 15, "expected_version": 0}
        ]}
        url = self.base.replace("http://", "ws://") + path + "/ws"
        with connect(url, origin=ORIGIN, additional_headers={"Cookie": owner.cookie_header()}) as live:
            live.send(json.dumps({"type": "hello", "last_revision": 0}))
            status, ack = owner.send("POST", path + "/ops", operation)
            self.assertEqual(status, 200)
            self.assertEqual(ack["input_revision"], 1)
            messages = self.receive_revision(live)
            self.assertLess(messages.index(("op", 1)), messages.index(("calc", 1)))
            calculated = [revision for kind, revision in messages if kind == "calc"]
            self.assertEqual(calculated, sorted(calculated))
            self.assertEqual(owner.send("POST", path + "/ops", operation), (200, ack))
        status, cell = owner.send("GET", path + "/sheets/2/cell?r=1&c=1")
        self.assertEqual((status, cell["v"]), (200, 30))
        status, workbook = owner.send("GET", path + "/workbook")
        self.assertEqual((status, workbook["calculated_revision"]), (200, 1))
        self.check_viewer_revocation(owner, path)

    def receive_revision(self, live):
        messages = []
        for _ in range(20):
            event = json.loads(live.recv(timeout=15))
            if event.get("type") in {"op", "calc"}:
                revision = event.get("input_revision") if event["type"] == "op" else event.get("calculated_revision")
                messages.append((event["type"], revision))
            if event.get("type") == "calc" and event.get("calculated_revision") == 1:
                return messages
        self.fail("revision 1 was not published over the real WebSocket")

    def test_fresh_copy_baseline_is_calculated_before_it_is_current(self):
        owner = HTTPClient(self.base)
        owner.signup("transport-baseline@example.test")
        copy_id = self.make_copy(owner)
        path = f"/api/copies/{copy_id}"
        status, cell = owner.send("GET", path + "/sheets/2/cell?r=1&c=1")
        self.assertEqual((status, cell["v"]), (200, 20))
        status, workbook = owner.send("GET", path + "/workbook")
        self.assertEqual(status, 200)
        self.assertEqual(workbook["input_revision"], workbook["calculated_revision"])

    def test_same_value_edit_preserves_the_complete_formula_snapshot(self):
        owner = HTTPClient(self.base)
        owner.signup("transport-repeat@example.test")
        path = f"/api/copies/{self.make_copy(owner)}"
        for version in range(2):
            status, _ = owner.send("POST", path + "/ops", {
                "op_id": f"repeat-{version}", "action": "edit", "cells": [
                    {"sheet": 0, "r": 1, "c": 1, "v": 15, "expected_version": version}
                ]
            })
            self.assertEqual(status, 200)
            self.assertEqual(owner.send("GET", path + "/sheets/2/cell?r=1&c=1")[1]["v"], 30)
            self.assertEqual(owner.send("GET", path + "/sheets/2/cell?r=16&c=1")[1]["v"], "x b")

    def test_api_error_value_recalculates_to_a_typed_formula_error(self):
        owner = HTTPClient(self.base)
        owner.signup("transport-error-value@example.test")
        path = f"/api/copies/{self.make_copy(owner)}"
        error = {"error": "#DIV/0!"}
        self.http_edit(owner, path, (error, 0))
        status, cell = owner.send("GET", path + "/sheets/2/cell?r=1&c=1")
        self.assertEqual((status, cell["v"]), (200, error))
        self.assertEqual(owner.send("GET", path + "/workbook")[1]["calculated_revision"], 1)

    def test_logout_closes_an_existing_real_websocket(self):
        owner = HTTPClient(self.base)
        owner.signup("transport-logout@example.test")
        path = f"/api/copies/{self.make_copy(owner)}"
        url = self.base.replace("http://", "ws://") + path + "/ws"
        with connect(url, origin=ORIGIN, additional_headers={"Cookie": owner.cookie_header()}) as live:
            live.send(json.dumps({"type": "hello", "last_revision": 0}))
            self.assertEqual(owner.send("POST", "/api/auth/logout")[0], 204)
            self.assert_socket_revoked(live)

    def test_refused_socket_reaches_the_browser_as_4403_not_a_handshake_error(self):
        # A handshake 403 shows up in browsers as code 1006, so the grid retried every second forever.
        url = self.base.replace("http://", "ws://") + "/api/copies/missing/ws"
        with connect(url, origin=ORIGIN) as live, self.assertRaises(ConnectionClosed) as closed:
            live.recv(timeout=5)
        self.assertEqual(closed.exception.rcvd.code, 4403)

    def test_http_conflicts_locked_cells_and_formula_text_are_rejected_atomically(self):
        owner = HTTPClient(self.base)
        owner.signup("transport-validation@example.test")
        path = f"/api/copies/{self.make_copy(owner)}"
        cases = [(0, 0, 1, 12, 0, "locked", 422),
                 (0, 1, 1, "=1+1", 0, "invalid", 422),
                 (0, 1, 1, 12, 5, "conflict", 409)]
        for sheet, row, col, value, version, error, status in cases:
            actual, body = owner.send("POST", path + "/ops", {
                "op_id": f"invalid-{error}", "action": "edit", "cells": [
                    {"sheet": sheet, "r": row, "c": col, "v": value, "expected_version": version}
                ]
            })
            self.assertEqual((actual, body["error"]), (status, error))
        self.assertEqual(owner.send("GET", path + "/workbook")[1]["input_revision"], 0)

    def test_http_category_role_inheritance_and_fork_calculation_isolation(self):
        owner, member = HTTPClient(self.base), HTTPClient(self.base)
        owner.signup("transport-fork-owner@example.test")
        member.signup("transport-fork-member@example.test")
        _, category = owner.send("POST", "/api/categories", {"name": "Fork Category"})
        category_id = category["category"]["id"]
        self.assertEqual(owner.send("PUT", f"/api/categories/{category_id}/members", {
            "email": "transport-fork-member@example.test", "role": "member"
        })[0], 200)
        status, body = owner.send("POST", "/api/copies", {
            "category_id": category_id, "name": "Shared Parent",
            "visibility": "category", "default_role": "editor"
        })
        self.assertEqual(status, 201)
        path = f"/api/copies/{body['copy']['id']}"
        self.assertEqual(member.send("GET", "/api/copies")[1]["copies"][0]["my_role"], "editor")
        self.http_edit(member, path, (12, 0))
        status, fork = member.send("POST", path + "/fork", {"name": "Frozen Fork"})
        self.assertEqual(status, 201)
        self.http_edit(owner, path, (15, 1))
        fork_path = f"/api/copies/{fork['copy']['id']}"
        self.assertEqual(member.send("GET", fork_path + "/sheets/2/cell?r=1&c=1")[1]["v"], 24)
        self.assertEqual(owner.send("GET", path + "/sheets/2/cell?r=1&c=1")[1]["v"], 30)

    def http_edit(self, client, path, change):
        value, version = change
        status, _ = client.send("POST", path + "/ops", {
            "op_id": f"edit-{value}-{version}", "action": "edit", "cells": [
                {"sheet": 0, "r": 1, "c": 1, "v": value, "expected_version": version}
            ]
        })
        self.assertEqual(status, 200)

    def assert_socket_revoked(self, live):
        try:
            for _ in range(20):
                live.recv(timeout=3)
        except ConnectionClosed as closed:
            self.assertEqual(closed.rcvd.code, 4403)
            return
        self.fail("revoked session retained a real WebSocket")

    def check_viewer_revocation(self, owner, path):
        viewer = HTTPClient(self.base)
        email = "transport-viewer@example.test"
        viewer.signup(email)
        status, result = viewer.send("GET", "/api/copies")
        self.assertEqual((status, result["copies"]), (200, []))
        self.assertEqual(owner.send("PUT", path + "/shares", {"email": email, "role": "viewer"})[0], 200)
        status, _ = viewer.send("POST", path + "/ops", {
            "op_id": "viewer-write", "action": "edit", "cells": [
                {"sheet": 0, "r": 1, "c": 1, "v": 99, "expected_version": 1}
            ]
        })
        self.assertEqual(status, 403)
        url = self.base.replace("http://", "ws://") + path + "/ws"
        with connect(url, origin=ORIGIN, additional_headers={"Cookie": viewer.cookie_header()}) as live:
            live.send(json.dumps({"type": "hello", "last_revision": 1}))
            self.assertEqual(owner.send("PUT", path + "/shares", {"email": email, "role": None})[0], 200)
            try:
                for _ in range(20):
                    live.recv(timeout=10)
            except ConnectionClosed as closed:
                self.assertEqual(closed.rcvd.code, 4403)
                return
            self.fail("revoked viewer retained the real WebSocket connection")
