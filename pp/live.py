"""Authorized WebSocket replay, broadcast, revocation, and presence."""

from __future__ import annotations

import asyncio
from contextlib import asynccontextmanager
from datetime import datetime, timezone
import json
import threading

from starlette.websockets import WebSocketDisconnect

from .authorization import effective_role
from .errors import AppError


SEND_TIMEOUT = 2
CLOSE_TIMEOUT = 2


class LiveProtocolError(Exception):
    def __init__(self, code):
        super().__init__(code)
        self.code = code


class Connection:
    def __init__(self, websocket, copy_id, user):
        self.websocket = websocket
        self.copy_id = copy_id
        self.user = user
        self.presence = None
        self.ready = False
        self.seen = set()


class LiveHub:
    def __init__(self, database, replay_window: int):
        self.database = database
        self.replay_window = replay_window
        self._lock = threading.RLock()
        self._connections: set[Connection] = set()
        self._publication_locks = {}
        self._pipeline_locks = {}

    def add(self, connection):
        with self._lock:
            self._connections.add(connection)

    def discard(self, connection):
        with self._lock:
            self._connections.discard(connection)

    async def broadcast(self, copy_id: str, payload: dict):
        async with _serialized(self._publication_lock(copy_id)):
            with self._lock:
                targets = [item for item in self._connections
                           if item.copy_id == copy_id and item.ready]
            for target in targets:
                await self._send(target, payload)

    def pipeline(self, copy_id):
        return _serialized(self._pipeline_lock(copy_id))

    async def recheck_user(self, user_id: str):
        with self._lock:
            targets = [item for item in self._connections if item.user["id"] == user_id]
        for target in targets:
            if not self.authorized(target):
                await self._close(target)

    async def recheck_copy(self, copy_id: str):
        with self._lock:
            targets = [item for item in self._connections if item.copy_id == copy_id]
        for target in targets:
            if not self.authorized(target):
                await self._close(target)

    async def recheck_session(self, session_hash: str):
        with self._lock:
            targets = [item for item in self._connections
                       if item.user.get("session_hash") == session_hash]
        for target in targets:
            await self._close(target)

    async def hello(self, connection, last_revision):
        if type(last_revision) is not int or last_revision < 0:
            raise AppError(422, "invalid", "last_revision must be a non-negative integer")
        async with _serialized(self._publication_lock(connection.copy_id)):
            copy, events = self._replay(connection.copy_id, last_revision)
            if events is None:
                await self._send(connection, {"type": "resync"})
            else:
                for event in events:
                    payload = json.loads(event["payload_json"])
                    if payload.get("type") == "op" and "at" not in payload:
                        payload["at"] = event["created_at"]
                    await self._send(connection, payload)
            await self._send(connection, {"type": "calc_state", "input_revision": copy["input_revision"],
                                          "calculated_revision": copy["calculated_revision"]})
            connection.ready = True

    async def presence(self, connection, payload):
        sheet, selection = payload.get("sheet"), payload.get("sel")
        if type(sheet) is not int or sheet < 0 or not isinstance(selection, list) or len(selection) != 4:
            raise AppError(422, "invalid", "Presence selection is invalid")
        if any(type(value) is not int or value < 0 for value in selection):
            raise AppError(422, "invalid", "Presence selection is invalid")
        connection.presence = {"id": connection.user["id"], "name": connection.user["name"],
                               "color": _color(connection.user["id"]), "sheet": sheet, "sel": selection}
        with self._lock:
            users = [item.presence for item in self._connections
                     if item.copy_id == connection.copy_id and item.presence is not None]
        await self.broadcast(connection.copy_id, {"type": "presence", "users": users})

    def authorized(self, connection) -> bool:
        with self.database.connect() as database:
            session = database.execute(
                "SELECT 1 FROM pp_sessions WHERE token_hash=? AND user_id=? AND expires_at>?",
                (connection.user.get("session_hash"), connection.user["id"], _now()),
            ).fetchone()
            role = effective_role(database, connection.copy_id, connection.user["id"])
        return session is not None and role is not None

    def _publication_lock(self, copy_id):
        with self._lock:
            return self._publication_locks.setdefault(copy_id, threading.Lock())

    def _pipeline_lock(self, copy_id):
        with self._lock:
            return self._pipeline_locks.setdefault(copy_id, threading.Lock())

    def _replay(self, copy_id, last_revision):
        with self.database.connect() as database:
            copy = database.execute("SELECT * FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
            if last_revision > copy["input_revision"] or copy["input_revision"] - last_revision > self.replay_window:
                return copy, None
            events = _replay_events(database, copy_id, last_revision)
        if copy["calculated_revision"] >= last_revision:
            if not any(item["event_type"] == "calc" and item["input_revision"] == last_revision for item in events):
                return copy, None
        return copy, events

    async def _send(self, connection, payload):
        if not self.authorized(connection):
            await self._close(connection)
            return False
        key = _event_key(payload)
        if key is not None and key in connection.seen:
            return True
        try:
            await asyncio.wait_for(connection.websocket.send_json(payload), SEND_TIMEOUT)
            if key is not None:
                connection.seen.add(key)
            return True
        except asyncio.TimeoutError:
            await self._close(connection, 1013)
            return False
        except Exception:
            await self._close(connection, 1013)
            return False

    async def _close(self, connection, code=4403):
        await _bounded_close(connection.websocket, code)
        self.discard(connection)


async def serve_live(websocket, services, hub):
    try:
        token, user = _socket_user(websocket, services)
    except AppError:
        await _bounded_close(websocket, 4403)
        return
    await websocket.accept()
    from .db import token_hash

    connection = Connection(websocket, websocket.path_params["copy_id"],
                            {**user, "session_hash": token_hash(token)})
    hub.add(connection)
    try:
        while True:
            payload = await _receive_payload(websocket)
            await _dispatch(connection, payload, hub)
    except WebSocketDisconnect:
        pass
    except LiveProtocolError as exc:
        await _bounded_close(websocket, exc.code)
    except AppError:
        await _bounded_close(websocket, 1008)
    finally:
        hub.discard(connection)


def _socket_user(websocket, services):
    services.auth_origin(websocket)
    token = websocket.cookies.get("pp_session")
    user, _ = services.auth.authenticate(token)
    with services.database.connect() as database:
        role = effective_role(database, websocket.path_params["copy_id"], user["id"])
    if role is None:
        raise AppError(403, "forbidden", "Copy access is required")
    return token, user


async def _receive_payload(websocket):
    text = await websocket.receive_text()
    if len(text) > 65_536:
        raise LiveProtocolError(1009)
    try:
        payload = json.loads(text)
    except json.JSONDecodeError as exc:
        raise LiveProtocolError(1008) from exc
    if not isinstance(payload, dict):
        raise LiveProtocolError(1008)
    return payload


async def _dispatch(connection, payload, hub):
    if not hub.authorized(connection):
        raise LiveProtocolError(4403)
    if payload.get("type") == "hello":
        await hub.hello(connection, payload.get("last_revision"))
    elif payload.get("type") == "presence":
        await hub.presence(connection, payload)
    else:
        raise LiveProtocolError(1008)


def _color(user_id):
    palette = ["#2563eb", "#dc2626", "#16a34a", "#9333ea", "#d97706", "#0891b2"]
    return palette[sum(user_id.encode()) % len(palette)]


def _replay_events(database, copy_id, last_revision):
    return database.execute(
        "SELECT event_type,input_revision,payload_json,created_at FROM pp_events WHERE copy_id=? "
        "AND (input_revision>? OR (event_type='calc' AND input_revision=?)) ORDER BY sequence",
        (copy_id, last_revision, last_revision),
    ).fetchall()


def _event_key(payload):
    if payload.get("type") == "op":
        return "op", payload.get("input_revision"), payload.get("op_id")
    if payload.get("type") == "calc":
        return "calc", payload.get("calculated_revision")
    return None


async def _bounded_close(websocket, code):
    try:
        await asyncio.wait_for(websocket.close(code=code), CLOSE_TIMEOUT)
    except Exception:
        pass


@asynccontextmanager
async def _serialized(lock):
    while not lock.acquire(blocking=False):
        await asyncio.sleep(0.001)
    try:
        yield
    finally:
        lock.release()


def _now():
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
