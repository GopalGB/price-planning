from contextlib import contextmanager
import threading
import unittest
from unittest import mock

from pp.errors import AppError
from tests.v2.api_support import APIFixture


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


if __name__ == "__main__":
    unittest.main()
