"""Starlette HTTP and WebSocket boundary for Price Planning v2."""

from __future__ import annotations

import asyncio
from dataclasses import dataclass
import json
import logging
from pathlib import Path
from uuid import uuid4
import threading
import time

from starlette.applications import Starlette
from starlette.background import BackgroundTask
from starlette.exceptions import HTTPException
from starlette.requests import Request
from starlette.responses import FileResponse, JSONResponse, Response
from starlette.routing import Route, WebSocketRoute

from .auth import AuthService, require_csrf, require_origin
from .dashboard_views import DashboardViewQuery, DashboardViewService
from .categories import CategoryService, MembershipChange
from .calc_service import CalculationPublication, CalculationService
from .copies import CopyService, ShareChange
from .db import token_hash
from .engine.pool import CalculationRequest
from .errors import AppError
from .export.xlsm import export_xlsm
from .live import LiveHub, serve_live
from .ops import OperationService
from .pivot import PivotQuery, PivotService
from .workbook_service import CellQuery, HistoryQuery, RowQuery, TemplateCatalog, WorkbookService


LOGGER = logging.getLogger(__name__)
BODY_LIMIT = 2 * 1024 * 1024


@dataclass(frozen=True)
class ServiceContext:
    database: object
    snapshot: object
    config: object
    engine_pool: object = None


class Services:
    def __init__(self, context: ServiceContext):
        self.database = context.database
        self.snapshot = context.snapshot
        self.config = context.config
        templates = TemplateCatalog(context.database, context.snapshot)
        self.auth = AuthService(context.database, context.config)
        self.categories = CategoryService(context.database)
        self.copies = CopyService(context.database, context.snapshot.hash)
        self.workbooks = WorkbookService(context.database, templates)
        self.dashboard_views = DashboardViewService(context.database, templates)
        self.pivots = PivotService(context.database, context.snapshot)
        self.operations = OperationService(context.database, templates)
        self.calculations = CalculationService(context.database, templates)
        self.engine_pool = context.engine_pool
        self.hub = LiveHub(context.database, context.config.replay_window)
        self.writes = WriteLimiter()

    def auth_origin(self, connection):
        require_origin(connection.headers.get("origin"), self.config.allowed_origins)


def create_app(context: ServiceContext) -> Starlette:
    services = Services(context)
    routes = [
        Route("/api/auth/signup", _signup, methods=["POST"]),
        Route("/api/auth/login", _login, methods=["POST"]),
        Route("/api/auth/logout", _logout, methods=["POST"]),
        Route("/api/me", _me, methods=["GET"]),
        Route("/api/categories", _categories, methods=["GET", "POST"]),
        Route("/api/categories/{category_id}/members", _category_members, methods=["PUT"]),
        Route("/api/copies", _copies, methods=["GET", "POST"]),
        Route("/api/copies/{copy_id}/fork", _fork, methods=["POST"]),
        Route("/api/copies/{copy_id}", _copy_patch, methods=["PATCH"]),
        Route("/api/copies/{copy_id}/shares", _copy_shares, methods=["PUT"]),
        Route("/api/copies/{copy_id}/workbook", _workbook, methods=["GET"]),
        Route("/api/copies/{copy_id}/sheets/{sheet:int}/rows", _rows, methods=["GET"]),
        Route("/api/copies/{copy_id}/sheets/{sheet:int}/cell", _cell, methods=["GET"]),
        Route("/api/copies/{copy_id}/sheets/{sheet:int}/pivots/{name}/refresh",
              _pivot_refresh, methods=["POST"]),
        Route("/api/copies/{copy_id}/sheets/{sheet:int}/pivots/{name}/view",
              _dashboard_view, methods=["GET", "PUT"]),
        Route("/api/copies/{copy_id}/sheets/{sheet:int}/pivots/{name}/category/refresh",
              _category_pivot_refresh, methods=["POST"]),
        Route("/api/copies/{copy_id}/sheets/{sheet:int}/pivots/{name}/category/view",
              _category_dashboard_view, methods=["GET", "PUT"]),
        Route("/api/copies/{copy_id}/history", _history, methods=["GET"]),
        Route("/api/copies/{copy_id}/ops", _ops, methods=["POST"]),
        Route("/api/copies/{copy_id}/export.xlsm", _export, methods=["GET"]),
        WebSocketRoute("/api/copies/{copy_id}/ws", _websocket),
        Route("/api/{path:path}", _api_not_found),
        Route("/{path:path}", _static),
    ]
    app = Starlette(routes=routes, exception_handlers={AppError: _app_error, HTTPException: _http_error})
    app.state.services = services
    app.state.engine_pool = context.engine_pool
    app.add_middleware(SecurityBoundaryMiddleware)
    return app


async def _signup(request: Request):
    services = _services(request)
    require_origin(request.headers.get("origin"), services.config.allowed_origins)
    body = await _body(request)
    user, token, _ = services.auth.signup(body.get("email"), body.get("name"), body.get("password"))
    response = JSONResponse({"user": user}, status_code=201)
    _set_session(response, token, services.config.cookie_secure)
    return response


async def _login(request: Request):
    services = _services(request)
    require_origin(request.headers.get("origin"), services.config.allowed_origins)
    body = await _body(request)
    remote = request.client.host if request.client else "unknown"
    user, token, csrf = await asyncio.to_thread(
        services.auth.login, body.get("email"), body.get("password"), remote
    )
    response = JSONResponse({"user": user, "csrf": csrf})
    _set_session(response, token, services.config.cookie_secure)
    return response


async def _logout(request: Request):
    services, _, _ = _write_auth(request)
    token = request.cookies.get("pp_session", "")
    services.auth.logout(token)
    await services.hub.recheck_session(token_hash(token))
    response = Response(status_code=204)
    response.delete_cookie("pp_session", path="/")
    return response


async def _me(request: Request):
    services = _services(request)
    user, csrf = services.auth.authenticate(request.cookies.get("pp_session"))
    return JSONResponse({"user": user, "csrf": csrf})


async def _categories(request: Request):
    services = _services(request)
    if request.method == "GET":
        user, _ = services.auth.authenticate(request.cookies.get("pp_session"))
        return JSONResponse({"categories": services.categories.list(user["id"])})
    services, user, _ = _write_auth(request)
    body = await _body(request)
    category = services.categories.create(user["id"], body.get("name"), _session_check(request, services))
    return JSONResponse({"category": category}, status_code=201)


async def _category_members(request: Request):
    services, user, _ = _write_auth(request)
    body = await _body(request)
    change = MembershipChange(user["id"], request.path_params["category_id"],
                              body.get("email"), body.get("role"))
    category = services.categories.update_member(change, _session_check(request, services))
    await services.hub.recheck_user(_user_id(services.database, body.get("email")))
    return JSONResponse({"category": category})


async def _copies(request: Request):
    services = _services(request)
    if request.method == "GET":
        user, _ = services.auth.authenticate(request.cookies.get("pp_session"))
        copies = services.copies.list(user["id"], request.query_params.get("category"))
        occupants = services.hub.occupants({copy["id"] for copy in copies})
        for copy in copies:
            copy["occupants"] = occupants.get(copy["id"], [])
        return JSONResponse({"copies": copies})
    services, user, _ = _write_auth(request)
    body = await _body(request)
    copy = services.copies.create(user["id"], body, _session_check(request, services))
    copy = await _calculate_baseline(services, copy, user["id"])
    return JSONResponse({"copy": copy}, status_code=201)


async def _fork(request: Request):
    services, user, _ = _write_auth(request)
    body = await _body(request)
    copy = services.copies.fork(user["id"], request.path_params["copy_id"], body.get("name"),
                                _session_check(request, services))
    copy = await _calculate_baseline(services, copy, user["id"])
    return JSONResponse({"copy": copy}, status_code=201)


async def _copy_patch(request: Request):
    services, user, _ = _write_auth(request)
    body = await _body(request)
    copy = services.copies.patch(user["id"], request.path_params["copy_id"], body, _session_check(request, services))
    await services.hub.recheck_copy(copy["id"])
    return JSONResponse({"copy": copy})


async def _copy_shares(request: Request):
    services, user, _ = _write_auth(request)
    body = await _body(request)
    change = ShareChange(user["id"], request.path_params["copy_id"],
                         body.get("email"), body.get("role"))
    copy = services.copies.share(change, _session_check(request, services))
    await services.hub.recheck_user(_user_id(services.database, body.get("email")))
    return JSONResponse({"copy": copy})


async def _workbook(request: Request):
    services, user = _read_auth(request)
    return JSONResponse(services.workbooks.workbook(request.path_params["copy_id"], user["id"]))


async def _rows(request: Request):
    services, user = _read_auth(request)
    start = _integer(request.query_params.get("start"), "start")
    count = _integer(request.query_params.get("count"), "count")
    query = RowQuery(request.path_params["copy_id"], user["id"],
                     request.path_params["sheet"], start, count)
    body = services.workbooks.rows(query)
    return JSONResponse(body)


async def _cell(request: Request):
    services, user = _read_auth(request)
    row = _integer(request.query_params.get("r"), "r")
    col = _integer(request.query_params.get("c"), "c")
    query = CellQuery(request.path_params["copy_id"], user["id"],
                      request.path_params["sheet"], row, col)
    body = services.workbooks.cell(query)
    return JSONResponse(body)


async def _pivot_refresh(request: Request):
    services, user = _read_auth(request)
    require_origin(request.headers.get("origin"), services.config.allowed_origins)
    query = PivotQuery(request.path_params["copy_id"], user["id"],
                       request.path_params["sheet"], request.path_params["name"],
                       await _body(request))
    body = await asyncio.to_thread(services.pivots.refresh, query)
    return JSONResponse(body)


async def _dashboard_view(request: Request):
    if request.method == "GET":
        services, user = _read_auth(request)
    else:
        services, user, _ = _write_auth(request)
    query = DashboardViewQuery(request.path_params["copy_id"], user["id"],
                               request.path_params["sheet"], request.path_params["name"])
    if request.method == "GET":
        return JSONResponse(services.dashboard_views.get(query))
    body = services.dashboard_views.update(
        query, await _body(request), _session_check(request, services)
    )
    return JSONResponse(body)


async def _category_pivot_refresh(request: Request):
    services, user = _read_auth(request)
    require_origin(request.headers.get("origin"), services.config.allowed_origins)
    query = PivotQuery(request.path_params["copy_id"], user["id"],
                       request.path_params["sheet"], request.path_params["name"],
                       await _body(request))
    body = await asyncio.to_thread(services.pivots.refresh_category, query)
    return JSONResponse(body)


async def _category_dashboard_view(request: Request):
    if request.method == "GET":
        services, user = _read_auth(request)
    else:
        services, user, _ = _write_auth(request)
    query = DashboardViewQuery(request.path_params["copy_id"], user["id"],
                               request.path_params["sheet"], request.path_params["name"])
    if request.method == "GET":
        return JSONResponse(services.dashboard_views.get_category(query))
    body = services.dashboard_views.update_category(
        query, await _body(request), _session_check(request, services)
    )
    return JSONResponse(body)


async def _history(request: Request):
    services, user = _read_auth(request)
    values = {name: _optional_integer(request.query_params.get(name), name)
              for name in ("sheet", "r", "c")}
    limit = _optional_integer(request.query_params.get("limit"), "limit") or 50
    query = HistoryQuery(request.path_params["copy_id"], user["id"],
                         values["sheet"], values["r"], values["c"], limit)
    body = services.workbooks.history(query)
    return JSONResponse(body)


async def _ops(request: Request):
    services, user, _ = _write_auth(request)
    copy_id, payload = request.path_params["copy_id"], await _body(request)
    async with services.hub.pipeline(copy_id):
        ack, created = services.operations.apply_with_status(copy_id, user["id"], payload,
                                                              _session_check(request, services))
        if created:
            await _calculate_and_publish(services, copy_id, ack["input_revision"])
    return JSONResponse(ack)


async def _calculate_and_publish(services, copy_id, revision):
    event = _event(services.database, copy_id, ("op", revision))
    await services.hub.broadcast(copy_id, event)
    state = _calc_state(services.database, copy_id)
    await services.hub.broadcast(copy_id, {"type": "calc_state", **state})
    if services.engine_pool is None:
        return
    try:
        source = services.workbooks.calculation_source(copy_id)
        changes = services.operations.revision_changes(copy_id, revision)
        request = CalculationRequest(copy_id, revision, source, changes)
        result = await asyncio.to_thread(services.engine_pool.apply, request)
        publication = CalculationPublication(copy_id, revision, result["results"])
        published = await asyncio.to_thread(services.calculations.complete, publication)
        if published:
            await services.hub.broadcast(copy_id, _event(services.database, copy_id, ("calc", revision)))
    except Exception as exc:
        await asyncio.to_thread(services.calculations.failed, copy_id, revision, type(exc).__name__)


async def _calculate_baseline(services, copy, user_id):
    if services.engine_pool is None:
        return copy
    async with services.hub.pipeline(copy["id"]):
        try:
            source = services.workbooks.calculation_source(copy["id"])
            changes = services.operations.revision_changes(copy["id"], 0)
            request = CalculationRequest(copy["id"], 0, source, changes)
            result = await asyncio.to_thread(services.engine_pool.apply, request)
            publication = CalculationPublication(copy["id"], 0, result["results"])
            published = await asyncio.to_thread(services.calculations.complete, publication)
            if published:
                event = _event(services.database, copy["id"], ("calc", 0))
                await services.hub.broadcast(copy["id"], event)
                state = _calc_state(services.database, copy["id"])
                await services.hub.broadcast(copy["id"], {"type": "calc_state", **state})
        except Exception as exc:
            await asyncio.to_thread(services.calculations.failed, copy["id"], 0, type(exc).__name__)
            raise
    return services.copies.get(user_id, copy["id"])


async def _export(request: Request):
    services, user = _read_auth(request)
    source, overrides = services.workbooks.source_and_overrides(request.path_params["copy_id"], user["id"])
    output_dir = Path(services.config.data_dir) / "exports"
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f"price-plan-{uuid4().hex}.xlsm"
    cells = [(*coordinate, value) for coordinate, value in sorted(overrides.items())]
    try:
        await asyncio.to_thread(export_xlsm, source, cells, path)
    except BaseException:
        path.unlink(missing_ok=True)
        raise
    return FileResponse(path, media_type="application/vnd.ms-excel.sheet.macroenabled.12",
                        filename="price-plan.xlsm", background=BackgroundTask(path.unlink, missing_ok=True))


async def _websocket(websocket):
    await serve_live(websocket, websocket.app.state.services, websocket.app.state.services.hub)


def _services(request):
    return request.app.state.services


def _read_auth(request):
    services = _services(request)
    user, _ = services.auth.authenticate(request.cookies.get("pp_session"))
    return services, user


def _session_check(request, services):
    token = request.cookies.get("pp_session")
    return lambda connection: services.auth.authenticate_in_transaction(connection, token)


def _write_auth(request):
    services = _services(request)
    require_origin(request.headers.get("origin"), services.config.allowed_origins)
    user, csrf = services.auth.authenticate(request.cookies.get("pp_session"))
    require_csrf(request.headers.get("x-csrf-token"), csrf)
    services.writes.check(user["id"])
    return services, user, csrf


async def _body(request: Request) -> dict:
    length = int(request.headers.get("content-length", "0") or 0)
    if length > BODY_LIMIT:
        raise AppError(413, "too_large", "Request body is too large")
    try:
        raw = await _bounded_body(request)
        value = json.loads(raw)
    except AppError:
        raise
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise AppError(400, "invalid_json", "Request body must be JSON") from exc
    if not isinstance(value, dict):
        raise AppError(422, "invalid", "Request body must be an object")
    return value


async def _bounded_body(request):
    chunks, size = [], 0
    async for chunk in request.stream():
        size += len(chunk)
        if size > BODY_LIMIT:
            raise AppError(413, "too_large", "Request body is too large")
        chunks.append(chunk)
    return b"".join(chunks)


def _set_session(response, token, secure):
    response.set_cookie("pp_session", token, max_age=8 * 60 * 60, path="/", secure=secure,
                        httponly=True, samesite="strict")


def _integer(value, name):
    try:
        result = int(value)
    except (TypeError, ValueError) as exc:
        raise AppError(422, "invalid", f"{name} must be an integer") from exc
    return result


def _optional_integer(value, name):
    return None if value is None else _integer(value, name)


def _user_id(database, email):
    with database.connect() as connection:
        row = connection.execute("SELECT id FROM pp_users WHERE email=?", (str(email).strip().lower(),)).fetchone()
    return row["id"] if row else ""


def _event(database, copy_id, key):
    event_type, revision = key
    with database.connect() as connection:
        row = connection.execute(
            "SELECT payload_json FROM pp_events WHERE copy_id=? AND event_type=? AND input_revision=? ORDER BY sequence DESC LIMIT 1",
            (copy_id, event_type, revision),
        ).fetchone()
    return json.loads(row["payload_json"])


def _calc_state(database, copy_id):
    with database.connect() as connection:
        row = connection.execute("SELECT input_revision,calculated_revision FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
    return dict(row)


async def _app_error(request, exc: AppError):
    return JSONResponse({"error": exc.code, "message": exc.message, **exc.extra}, status_code=exc.status)


async def _http_error(request, exc: HTTPException):
    code = "not_found" if exc.status_code == 404 else "method_not_allowed" if exc.status_code == 405 else "http_error"
    return JSONResponse({"error": code, "message": str(exc.detail)}, status_code=exc.status_code)


async def _api_not_found(request):
    raise AppError(404, "not_found", "Endpoint not found")


async def _static(request):
    root = Path(__file__).resolve().parents[1] / "static" / "dist"
    relative = request.path_params.get("path", "") or "index.html"
    candidate = (root / relative).resolve()
    if root.is_dir() and candidate.is_file() and root.resolve() in candidate.parents:
        return FileResponse(candidate)
    if relative == "index.html":
        return Response("<!doctype html><title>Price Planning</title><h1>Price Planning v2 backend is running</h1>",
                        media_type="text/html")
    raise AppError(404, "not_found", "File not found")


class WriteLimiter:
    def __init__(self, limit=240):
        self.limit = limit
        self._entries = {}
        self._lock = threading.Lock()

    def check(self, key):
        now = time.monotonic()
        with self._lock:
            entries = [stamp for stamp in self._entries.get(key, []) if now - stamp < 60]
            if len(entries) >= self.limit:
                raise AppError(429, "rate_limited", "Write rate limit exceeded")
            entries.append(now)
            self._entries[key] = entries


class SecurityBoundaryMiddleware:
    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return
        started = False

        async def secure_send(message):
            nonlocal started
            if message["type"] == "http.response.start":
                started = True
                message["headers"] = _security_headers(message.get("headers", []))
            await send(message)

        try:
            await self.app(scope, receive, secure_send)
        except Exception as exc:
            LOGGER.error("Unhandled v2 request error type=%s", type(exc).__name__)
            if not started:
                response = JSONResponse({"error": "internal", "message": "The server could not complete the request"},
                                        status_code=500)
                await response(scope, receive, secure_send)


def _security_headers(headers):
    remove = {b"x-content-type-options", b"x-frame-options", b"referrer-policy", b"cache-control",
              b"content-security-policy"}
    result = [(key, value) for key, value in headers if key.lower() not in remove]
    result.extend([
        (b"x-content-type-options", b"nosniff"), (b"x-frame-options", b"DENY"),
        (b"referrer-policy", b"same-origin"), (b"cache-control", b"no-store"),
        (b"content-security-policy", b"default-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; frame-ancestors 'none'; object-src 'none'; base-uri 'none'"),
    ])
    return result
