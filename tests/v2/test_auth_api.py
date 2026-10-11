import asyncio
from contextlib import contextmanager
from dataclasses import dataclass
import json
import threading
import unittest
from unittest import mock
from urllib.parse import urlsplit

from pp.db import token_hash
from pp.errors import AppError
from tests.v2.asgi_client import ASGIResponse
from tests.v2.api_support import APIFixture


MUTATION_TABLES = (
    "pp_categories", "pp_category_members", "pp_copies", "pp_copy_shares",
    "pp_copy_baselines", "pp_overrides", "pp_ops", "pp_history",
    "pp_calc_jobs", "pp_results", "pp_events",
)


@dataclass(frozen=True)
class _MutationContext:
    fixture: APIFixture
    client: object
    csrf: str


@dataclass(frozen=True)
class _RequestSpec:
    method: str
    path: str
    body: dict
    copy_id: str | None = None


class _ObservedPipelineLock:
    def __init__(self):
        self._lock = threading.Lock()
        self.attempted = threading.Event()

    def acquire(self, blocking=True):
        if not blocking:
            self.attempted.set()
        return self._lock.acquire(blocking=blocking)

    def release(self):
        self._lock.release()

    def locked(self):
        return self._lock.locked()


class _PausedRequest:
    def __init__(self, context):
        self.context = context

    async def run(self, spec, during_pause):
        raw = json.dumps(spec.body).encode()
        entered, release = threading.Event(), threading.Event()
        messages = []
        receive = _split_receiver(raw, entered, release)

        async def send(message):
            messages.append(message)

        task = asyncio.create_task(self.context.fixture.app(
            _request_scope(self.context, spec, raw), receive, send,
        ))
        if not await asyncio.to_thread(entered.wait, 5):
            task.cancel()
            raise AssertionError("request body did not reach the pause barrier")
        try:
            await asyncio.to_thread(during_pause)
        finally:
            release.set()
        await task
        return _asgi_response(messages)


def _split_receiver(raw, entered, release):
    state = 0
    pivot = max(1, len(raw) // 2)

    async def receive():
        nonlocal state
        if state == 0:
            state = 1
            return {"type": "http.request", "body": raw[:pivot], "more_body": True}
        if state == 1:
            state = 2
            entered.set()
            await asyncio.to_thread(release.wait)
            return {"type": "http.request", "body": raw[pivot:], "more_body": False}
        return {"type": "http.disconnect"}

    return receive


def _request_scope(context, spec, raw):
    parsed = urlsplit(spec.path)
    headers = {
        "origin": context.client.origin, "host": "testserver", "content-type": "application/json",
        "content-length": str(len(raw)), "cookie": "; ".join(
            f"{key}={value}" for key, value in context.client.cookies.items()),
        "x-csrf-token": context.csrf,
    }
    return {
        "type": "http", "asgi": {"version": "3.0"}, "http_version": "1.1",
        "method": spec.method, "scheme": "http", "path": parsed.path,
        "raw_path": parsed.path.encode(), "query_string": parsed.query.encode(),
        "headers": [(key.encode(), value.encode()) for key, value in headers.items()],
        "client": ("127.0.0.1", 1234), "server": ("testserver", 80),
    }


def _asgi_response(messages):
    start = next(message for message in messages if message["type"] == "http.response.start")
    body = b"".join(message.get("body", b"") for message in messages
                    if message["type"] == "http.response.body")
    return ASGIResponse(start["status"], start["headers"], body)


def _mutation_state(fixture):
    with fixture.database.connect() as connection:
        return {
            table: tuple(sorted((tuple(row) for row in connection.execute(f"SELECT * FROM {table}")), key=repr))
            for table in MUTATION_TABLES
        }


def _delete_session(context):
    with context.fixture.database.transaction() as connection:
        changed = connection.execute(
            "DELETE FROM pp_sessions WHERE token_hash=?",
            (token_hash(context.client.cookies["pp_session"]),),
        ).rowcount
    if changed != 1:
        raise AssertionError(f"expected one persisted session deletion, got {changed}")


def _expire_session(context):
    with context.fixture.database.transaction() as connection:
        changed = connection.execute(
            "UPDATE pp_sessions SET expires_at='2000-01-01T00:00:00Z' WHERE token_hash=?",
            (token_hash(context.client.cookies["pp_session"]),),
        ).rowcount
    if changed != 1:
        raise AssertionError(f"expected one persisted session expiry, got {changed}")


def _require_json(response, status, key):
    if response.status_code != status:
        raise AssertionError(response.body)
    return response.json()[key]


def _create_category(context):
    response = context.client.post(
        "/api/categories", {"name": "Existing Category"}, {"x-csrf-token": context.csrf},
    )
    return _require_json(response, 201, "category")


def _create_copy(context):
    category = _create_category(context)
    response = context.client.post("/api/copies", {
        "category_id": category["id"], "name": "Existing Copy",
        "visibility": "private", "default_role": "viewer",
    }, {"x-csrf-token": context.csrf})
    return category, _require_json(response, 201, "copy")


def _category_create_case(_context):
    return _RequestSpec("POST", "/api/categories", {"name": "Blocked Category"})


def _category_member_case(context):
    category = _create_category(context)
    return _RequestSpec("PUT", f"/api/categories/{category['id']}/members", {
        "email": "editor@example.test", "role": "member",
    })


def _copy_create_case(context):
    category = _create_category(context)
    return _RequestSpec("POST", "/api/copies", {
        "category_id": category["id"], "name": "Blocked Copy",
        "visibility": "private", "default_role": "viewer",
    })


def _copy_fork_case(context):
    _, copy = _create_copy(context)
    return _RequestSpec("POST", f"/api/copies/{copy['id']}/fork", {"name": "Blocked Fork"})


def _copy_patch_case(context):
    _, copy = _create_copy(context)
    return _RequestSpec("PATCH", f"/api/copies/{copy['id']}", {"name": "Blocked Rename"})


def _copy_share_case(context):
    _, copy = _create_copy(context)
    return _RequestSpec("PUT", f"/api/copies/{copy['id']}/shares", {
        "email": "editor@example.test", "role": "viewer",
    })


def _edit_case(context):
    _, copy = _create_copy(context)
    return _RequestSpec("POST", f"/api/copies/{copy['id']}/ops", {
        "op_id": "blocked-edit", "action": "edit",
        "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
    }, copy["id"])


def _undo_case(context):
    _, copy = _create_copy(context)
    response = context.client.post(f"/api/copies/{copy['id']}/ops", {
        "op_id": "setup-edit", "action": "edit",
        "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
    }, {"x-csrf-token": context.csrf})
    _require_json(response, 200, "op_id")
    return _RequestSpec("POST", f"/api/copies/{copy['id']}/ops", {
        "op_id": "blocked-undo", "action": "undo",
        "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 10, "expected_version": 1}],
    }, copy["id"])


MUTATION_CASES = (
    ("category create", _category_create_case),
    ("category member update", _category_member_case),
    ("copy create", _copy_create_case),
    ("copy fork", _copy_fork_case),
    ("copy patch", _copy_patch_case),
    ("copy share", _copy_share_case),
    ("operation edit", _edit_case),
    ("operation undo", _undo_case),
)


async def _pipeline_logout(context, spec, gate):
    hub = context.fixture.app.state.services.hub
    hub._pipeline_locks[spec.copy_id] = gate
    gate.acquire()
    try:
        task = asyncio.create_task(context.client._request(
            spec.method, spec.path, spec.body, {"x-csrf-token": context.csrf},
        ))
        if not await asyncio.to_thread(gate.attempted.wait, 5):
            task.cancel()
            raise AssertionError("operation did not wait for the pipeline lock")
        logout = await asyncio.to_thread(
            context.client.post, "/api/auth/logout", None, {"x-csrf-token": context.csrf},
        )
        gate.release()
        return logout, await task
    finally:
        if gate.locked():
            gate.release()


class AuthAPIContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture()

    def tearDown(self):
        self.fixture.close()

    def test_signup_session_me_csrf_and_logout_contract(self):
        """Weak cookies, missing Origin checks, or missing CSRF must fail."""
        client = self.fixture.client()

        signup = client.post("/api/auth/signup", {
            "email": "new@example.test", "name": "New User", "password": "correct-horse-battery"
        })

        self.assertEqual(signup.status_code, 201)
        self.assertEqual(set(signup.json()), {"user"})
        cookie = signup.headers["set-cookie"]
        self.assertIn("pp_session=", cookie)
        self.assertIn("httponly", cookie.lower())
        self.assertIn("samesite=strict", cookie.lower())
        me = client.get("/api/me")
        self.assertEqual(me.status_code, 200)
        csrf = me.json()["csrf"]
        self.assertEqual(client.post("/api/auth/logout").status_code, 403)
        logout = client.post("/api/auth/logout", headers={"x-csrf-token": csrf})
        self.assertEqual(logout.status_code, 204)
        self.assertEqual(client.get("/api/me").status_code, 401)

    def test_signup_is_demo_only_and_origin_validated(self):
        """Unauthenticated bootstrap writes still require a trusted Origin."""
        bad_origin = self.fixture.client()
        bad_origin.origin = "http://evil.test"
        response = bad_origin.post("/api/auth/signup", {
            "email": "bad@example.test", "name": "Bad", "password": "correct-horse-battery"
        })
        self.assertEqual((response.status_code, response.json()["error"]), (403, "origin"))
        short = self.fixture.client().post("/api/auth/signup", {
            "email": "short@example.test", "name": "Short", "password": "too-short"
        })
        self.assertEqual((short.status_code, short.json()["error"]), (422, "invalid"))

    def test_login_rate_limit_is_five_failures_per_minute(self):
        """Allowing a sixth password attempt in one minute must fail."""
        client, csrf, _ = self.fixture.signup("rate@example.test")
        client.post("/api/auth/logout", headers={"x-csrf-token": csrf})
        for attempt in range(5):
            response = client.post("/api/auth/login", {
                "email": "rate@example.test", "password": f"wrong-password-{attempt}"
            })
            self.assertEqual(response.status_code, 401)
        limited = client.post("/api/auth/login", {
            "email": "rate@example.test", "password": "wrong-password-6"
        })
        self.assertEqual((limited.status_code, limited.json()["error"]), (429, "rate_limited"))

    def test_login_rate_limit_covers_one_ip_rotating_email_addresses(self):
        """Changing the email must not bypass the five-attempt IP bucket."""
        client = self.fixture.client()
        for attempt in range(5):
            response = client.post("/api/auth/login", {
                "email": f"missing-{attempt}@example.test", "password": "wrong-password",
            })
            self.assertEqual(response.status_code, 401)
        limited = client.post("/api/auth/login", {
            "email": "another@example.test", "password": "wrong-password",
        })
        self.assertEqual((limited.status_code, limited.json()["error"]), (429, "rate_limited"))

    def test_password_hashing_runs_off_loop_and_outside_write_transaction(self):
        """Scrypt must not hold the ASGI loop or a SQLite BEGIN IMMEDIATE lock."""
        client, csrf, _ = self.fixture.signup("threaded-login@example.test")
        client.post("/api/auth/logout", headers={"x-csrf-token": csrf})
        database = self.fixture.database
        original_transaction = database.transaction
        state = threading.local()
        observations = []

        @contextmanager
        def tracked_transaction():
            state.in_transaction = True
            try:
                with original_transaction() as connection:
                    yield connection
            finally:
                state.in_transaction = False

        def inspect_hash(password, encoded):
            observations.append((threading.get_ident(), getattr(state, "in_transaction", False)))
            return False

        with mock.patch.object(database, "transaction", tracked_transaction):
            with mock.patch("pp.auth.verify_password", side_effect=inspect_hash):
                response = client.post("/api/auth/login", {
                    "email": "threaded-login@example.test", "password": "wrong-password",
                })
        self.assertEqual(response.status_code, 401)
        self.assertNotEqual(observations[0][0], threading.get_ident())
        self.assertIs(observations[0][1], False)

    def test_inflight_hashes_reserve_atomic_ip_rate_limit_slots(self):
        """Concurrent slow hashes must not all pass the pre-hash failure check."""
        auth, attempts = self.fixture.app.state.services.auth, []
        try:
            for index in range(5):
                attempts.append(auth.begin_login(
                    f"pending-{index}@example.test", "wrong-password", "same-ip"
                ))
            with self.assertRaises(AppError) as caught:
                auth.begin_login("pending-six@example.test", "wrong-password", "same-ip")
            self.assertEqual((caught.exception.status, caught.exception.code), (429, "rate_limited"))
        finally:
            for attempt in attempts:
                auth.cancel_login(attempt)

    def test_paused_authenticated_mutations_reject_revoked_persisted_sessions(self):
        """Removing the transaction recheck must let a revoked request change persisted state."""
        for name, build_case in MUTATION_CASES:
            with self.subTest(route=name):
                fixture = APIFixture()
                try:
                    client, csrf, _ = fixture.signup("paused@example.test", "Paused Writer")
                    context = _MutationContext(fixture, client, csrf)
                    spec = build_case(context)
                    before = _mutation_state(fixture)
                    response = asyncio.run(_PausedRequest(context).run(spec, lambda: _delete_session(context)))
                    self.assertEqual(response.status_code, 401)
                    self.assertEqual(response.json()["error"], "unauthorized")
                    self.assertEqual(_mutation_state(fixture), before)
                finally:
                    fixture.close()

    def test_paused_mutation_rejects_a_persisted_expired_session(self):
        """Checking only cookie presence must not authorize an expired persisted session."""
        client, csrf, _ = self.fixture.signup("expired@example.test", "Expired Writer")
        context = _MutationContext(self.fixture, client, csrf)
        spec = _category_create_case(context)
        before = _mutation_state(self.fixture)

        response = asyncio.run(_PausedRequest(context).run(spec, lambda: _expire_session(context)))

        self.assertEqual(response.status_code, 401)
        self.assertEqual(response.json()["error"], "unauthorized")
        self.assertEqual(_mutation_state(self.fixture), before)

    def test_operation_rechecks_session_after_waiting_for_the_pipeline_lock(self):
        """Checking before the real pipeline wait must not permit an operation after logout."""
        client, csrf, _ = self.fixture.signup("pipeline@example.test", "Pipeline Writer")
        context = _MutationContext(self.fixture, client, csrf)
        spec = _edit_case(context)
        before = _mutation_state(self.fixture)

        logout, response = asyncio.run(_pipeline_logout(context, spec, _ObservedPipelineLock()))

        self.assertEqual(logout.status_code, 204)
        self.assertEqual(response.status_code, 401)
        self.assertEqual(response.json()["error"], "unauthorized")
        self.assertEqual(_mutation_state(self.fixture), before)


if __name__ == "__main__":
    unittest.main()
