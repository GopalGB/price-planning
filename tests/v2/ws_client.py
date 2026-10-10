import asyncio
import json
import queue
import threading
import time
from urllib.parse import urlsplit


class WebSocketClient:
    def __init__(self, app, path, cookies=None, origin="http://testserver"):
        self.app, self.path, self.cookies, self.origin = app, path, cookies or {}, origin
        self.inbound, self.outbound = queue.Queue(), queue.Queue()
        self.thread = threading.Thread(target=self._thread_main, daemon=True)
        self.thread.start()
        self.inbound.put({"type": "websocket.connect"})
        first = self.outbound.get(timeout=5)
        self.accepted = first["type"] == "websocket.accept"
        self.close_code = first.get("code") if first["type"] == "websocket.close" else None

    def _thread_main(self):
        asyncio.run(self._run())

    async def _run(self):
        parsed = urlsplit(self.path)
        headers = [(b"origin", self.origin.encode()), (b"host", b"testserver")]
        if self.cookies:
            cookie = "; ".join(f"{key}={value}" for key, value in self.cookies.items())
            headers.append((b"cookie", cookie.encode()))
        scope = {"type": "websocket", "asgi": {"version": "3.0"}, "scheme": "ws",
                 "path": parsed.path, "raw_path": parsed.path.encode(),
                 "query_string": parsed.query.encode(), "headers": headers,
                 "client": ("127.0.0.1", 1234), "server": ("testserver", 80),
                 "subprotocols": []}

        async def receive():
            return await asyncio.to_thread(self.inbound.get)

        async def send(message):
            self.outbound.put(message)

        await self.app(scope, receive, send)

    def send_json(self, payload):
        self.send_text(json.dumps(payload))

    def send_text(self, text):
        self.inbound.put({"type": "websocket.receive", "text": text})

    def receive_json(self, timeout=5):
        message = self.outbound.get(timeout=timeout)
        if message["type"] == "websocket.close":
            self.close_code = message.get("code")
            return {"type": "closed", "code": self.close_code}
        return json.loads(message["text"])

    def wait_closed(self, timeout=5):
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            message = self.outbound.get(timeout=max(0.01, deadline - time.monotonic()))
            if message["type"] == "websocket.close":
                self.close_code = message.get("code")
                return self.close_code
        raise TimeoutError("WebSocket did not close")

    def close(self):
        if self.thread.is_alive():
            self.inbound.put({"type": "websocket.disconnect", "code": 1000})
            self.thread.join(timeout=5)
