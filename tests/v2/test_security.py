import asyncio
import json
import unittest

from tests.v2.api_support import APIFixture


class SecurityContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture()

    def tearDown(self):
        self.fixture.close()

    def test_security_headers_and_json_api_errors_are_consistent(self):
        """Success and error paths must carry browser defenses without internal detail."""
        client, _, _ = self.fixture.signup("security@example.test", "Security")
        success = client.get("/api/me")
        for name, value in {
            "x-content-type-options": "nosniff", "x-frame-options": "DENY",
            "referrer-policy": "same-origin", "cache-control": "no-store",
        }.items():
            self.assertEqual(success.headers[name], value)
        missing = client.get("/api/not-a-route")
        self.assertEqual((missing.status_code, missing.json()["error"]), (404, "not_found"))

    def test_content_security_policy_allows_required_app_assets_only(self):
        client, _, _ = self.fixture.signup("csp@example.test", "CSP")
        policy = client.get("/api/me").headers["content-security-policy"]
        self.assertIn("default-src 'self'", policy)
        self.assertIn("style-src 'self' 'unsafe-inline'", policy)
        self.assertIn("img-src 'self' data:", policy)
        self.assertIn("frame-ancestors 'none'", policy)
        self.assertIn("object-src 'none'", policy)
        self.assertIn("base-uri 'none'", policy)

    def test_actual_body_limit_and_session_expiry_are_enforced(self):
        """Trusting Content-Length or accepting an expired cookie must fail."""
        client, csrf, user = self.fixture.signup("expiry@example.test", "Expiry")
        large = "x" * (2 * 1024 * 1024 + 1)
        response = client.post("/api/categories", {"name": large}, {"x-csrf-token": csrf,
                                                                  "content-length": "1"})
        self.assertEqual((response.status_code, response.json()["error"]), (413, "too_large"))
        with self.fixture.database.transaction() as connection:
            connection.execute("UPDATE pp_sessions SET expires_at='2000-01-01T00:00:00Z' WHERE user_id=?",
                               (user["id"],))
        self.assertEqual(self.fixture.client().get("/api/me").status_code, 401)
        self.assertEqual(client.get("/api/me").status_code, 401)

    def test_streaming_body_limit_stops_before_later_chunks(self):
        """Buffering request.body must consume data after the request is already too large."""
        async def exercise():
            chunks = [
                {"type": "http.request", "body": b"x" * (2 * 1024 * 1024 + 1), "more_body": True},
                {"type": "http.request", "body": b"later", "more_body": False},
            ]
            sent, consumed = [], 0

            async def receive():
                nonlocal consumed
                message = chunks[consumed]
                consumed += 1
                return message

            async def send(message):
                sent.append(message)

            scope = {
                "type": "http", "asgi": {"version": "3.0"}, "http_version": "1.1",
                "method": "POST", "scheme": "http", "path": "/api/auth/signup",
                "raw_path": b"/api/auth/signup", "query_string": b"",
                "headers": [(b"origin", b"http://testserver"), (b"host", b"testserver")],
                "client": ("127.0.0.1", 1234), "server": ("testserver", 80),
            }
            await self.fixture.app(scope, receive, send)
            start = next(message for message in sent if message["type"] == "http.response.start")
            body = b"".join(message.get("body", b"") for message in sent
                            if message["type"] == "http.response.body")
            return consumed, start["status"], json.loads(body)

        consumed, status, payload = asyncio.run(exercise())
        self.assertEqual((status, payload["error"]), (413, "too_large"))
        self.assertEqual(consumed, 1)

    def test_internal_exception_is_sanitized(self):
        """Raw exception messages and stack details must never reach a caller."""
        client, _, _ = self.fixture.signup("internal@example.test", "Internal")
        original = self.fixture.app.state.services.copies.list

        def fail(*args, **kwargs):
            raise RuntimeError("SECRET_INTERNAL_DETAIL")

        self.fixture.app.state.services.copies.list = fail
        try:
            response = client.get("/api/copies")
        finally:
            self.fixture.app.state.services.copies.list = original
        self.assertEqual(response.status_code, 500)
        self.assertNotIn(b"SECRET_INTERNAL_DETAIL", response.body)
        self.assertEqual(set(response.json()), {"error", "message"})


if __name__ == "__main__":
    unittest.main()
