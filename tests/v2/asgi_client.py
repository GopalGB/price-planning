import asyncio
import json
from urllib.parse import urlsplit


class ASGIResponse:
    def __init__(self, status, headers, body):
        self.status_code = status
        self.headers = {key.decode().lower(): value.decode() for key, value in headers}
        self.body = body

    def json(self):
        return json.loads(self.body)


class ASGIClient:
    def __init__(self, app, origin="http://testserver"):
        self.app = app
        self.origin = origin
        self.cookies = {}

    def request(self, method, path, json_body=None, headers=None, raw_body=None):
        response = asyncio.run(self._request(method, path, json_body, headers or {}, raw_body))
        cookie = response.headers.get("set-cookie")
        if cookie:
            name, value = cookie.split(";", 1)[0].split("=", 1)
            if value:
                self.cookies[name] = value
            else:
                self.cookies.pop(name, None)
        return response

    async def _request(self, method, path, json_body, headers, raw_body=None):
        parsed = urlsplit(path)
        body = raw_body if raw_body is not None else b"" if json_body is None else json.dumps(json_body).encode()
        request_headers = {"origin": self.origin, "host": "testserver", **headers}
        if body and raw_body is None:
            request_headers["content-type"] = "application/json"
        if self.cookies:
            request_headers["cookie"] = "; ".join(f"{key}={value}" for key, value in self.cookies.items())
        scope = {
            "type": "http", "asgi": {"version": "3.0"}, "http_version": "1.1",
            "method": method, "scheme": "http", "path": parsed.path,
            "raw_path": parsed.path.encode(), "query_string": parsed.query.encode(),
            "headers": [(key.lower().encode(), value.encode()) for key, value in request_headers.items()],
            "client": ("127.0.0.1", 1234), "server": ("testserver", 80),
        }
        messages, sent = [], False

        async def receive():
            nonlocal sent
            if sent:
                return {"type": "http.disconnect"}
            sent = True
            return {"type": "http.request", "body": body, "more_body": False}

        async def send(message):
            messages.append(message)

        await self.app(scope, receive, send)
        start = next(message for message in messages if message["type"] == "http.response.start")
        content = b"".join(message.get("body", b"") for message in messages
                           if message["type"] == "http.response.body")
        return ASGIResponse(start["status"], start["headers"], content)

    def get(self, path, headers=None):
        return self.request("GET", path, headers=headers)

    def post(self, path, body=None, headers=None):
        return self.request("POST", path, body, headers)

    def put(self, path, body=None, headers=None):
        return self.request("PUT", path, body, headers)

    def patch(self, path, body=None, headers=None):
        return self.request("PATCH", path, body, headers)

    def raw(self, method, path, body, headers=None):
        return self.request(method, path, headers=headers, raw_body=body)
