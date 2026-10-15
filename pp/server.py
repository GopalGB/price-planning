"""Starlette HTTP and WebSocket boundary for Price Planning v2."""

from __future__ import annotations

import asyncio
from contextlib import asynccontextmanager
from dataclasses import dataclass, replace
import gzip
import io
import json
import logging
from pathlib import Path
import shutil
import tempfile
from urllib.parse import unquote
from uuid import uuid4
import threading
import time

from starlette.applications import Starlette
from starlette.datastructures import Headers, MutableHeaders
from starlette.exceptions import HTTPException
from starlette.requests import Request
from starlette.responses import FileResponse, JSONResponse, Response
from starlette.routing import Route, WebSocketRoute

from .auth import AuthService, require_csrf, require_origin
from .dashboard_views import DashboardViewQuery, DashboardViewService
from .categories import CategoryService, MembershipChange
from .calc_service import (
    CalculationPublication, CalculationService, PublicationBaseMismatch,
)
from .copies import CopyService, ShareChange
from .db import token_hash
from .engine.excel import REFRESH_ALL, convert_to_xlsm
from .engine.formualizer import EngineError, MacroTimeout
from .engine.pool import CalculationRequest, MacroRequest, SlicerRequest
from .pivot_state import with_refreshed, with_selection
from .config import BRAND_TYPES
from .errors import AppError
from .export.transfer import TransferRequest, export_transfer
from .jobs import JobStore
from .live import LiveHub, serve_live
from .ops import OperationService
from .pivot import PivotQuery, PivotService
from .upload import ConfirmRequest, PreviewRequest, UploadDependencies, UploadService
from .workbook_service import (CellQuery, HistoryQuery, RowQuery, TemplateCatalog, WorkbookService,
                               _calculated_revision, calc_error)


LOGGER = logging.getLogger(__name__)
BODY_LIMIT = 2 * 1024 * 1024
GZIP_MINIMUM_BYTES = 1000
GZIP_LEVEL = 6
JSON_MEDIA_TYPE = "application/json"
MAX_PREPARED_DOWNLOADS = 32
MAX_PREPARED_DOWNLOAD_BYTES = 512 * 1024 * 1024
MAX_ACTIVE_TRANSFERS = 2
MAX_PENDING_TRANSFERS = 4
PREPARED_DOWNLOAD_SECONDS = 2 * 60


@dataclass(frozen=True)
class ServiceContext:
    database: object
    snapshot: object
    config: object
    engine_pool: object = None


@dataclass(frozen=True)
class ExportJob:
    snapshot: dict
    copy_id: str
    fmt: str
    selected: set[str] | None
    path: Path


@dataclass(frozen=True)
class PreparedArtifact:
    path: Path
    user_id: str
    session_hash: str
    copy_id: str
    filename: str
    expires: float


class Services:
    def __init__(self, context: ServiceContext):
        self.database = context.database
        self.snapshot = context.snapshot
        self.config = context.config
        _sweep_transfer_files(context.config.data_dir)
        templates = TemplateCatalog(context.database, context.snapshot)
        self.auth = AuthService(context.database, context.config)
        self.categories = CategoryService(context.database)
        self.copies = CopyService(context.database, context.snapshot.hash)
        self.workbooks = WorkbookService(context.database, templates)
        self.dashboard_views = DashboardViewService(context.database, templates)
        self.pivots = PivotService(context.database, context.snapshot)
        self.operations = OperationService(context.database, templates)
        self.uploads = UploadService(UploadDependencies(
            context.database, templates, self.copies, self.operations,
            Path(context.config.data_dir) / "uploads" / "previews",
            convert_to_xlsm if getattr(context.config, "engine", None) == "excel" else None,
        ))
        self.calculations = CalculationService(context.database, templates)
        self.engine_pool = context.engine_pool
        self.hub = LiveHub(context.database, context.config.replay_window)
        self.writes = WriteLimiter()
        self.downloads = PreparedDownloadStore()
        self.jobs = JobStore()
        self.transfer_limiter = threading.BoundedSemaphore(MAX_ACTIVE_TRANSFERS)
        self.transfer_admission = threading.BoundedSemaphore(MAX_PENDING_TRANSFERS)

    def close(self):
        self.downloads.close()
        self.uploads.close()

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
        Route("/api/copies/{copy_id}/macros", _macros, methods=["POST"]),
        Route("/api/copies/{copy_id}/slicers", _slicers, methods=["POST"]),
        Route("/api/copies/{copy_id}/refresh", _refresh, methods=["POST"]),
        Route("/api/copies/{copy_id}/export.xlsm", _export, methods=["GET"]),
        Route("/api/copies/{copy_id}/export.xlsx", _export, methods=["GET"]),
        Route("/api/copies/{copy_id}/downloads", _prepare_download, methods=["POST"]),
        Route("/api/downloads/{token}", _prepared_download, methods=["GET"]),
        Route("/api/uploads/preview", _upload_preview, methods=["POST"]),
        Route("/api/copies/{copy_id}/uploads/preview", _upload_preview, methods=["POST"]),
        Route("/api/uploads/{token}/confirm", _upload_confirm, methods=["POST"]),
        Route("/api/jobs/{token}", _job_status, methods=["GET"]),
        WebSocketRoute("/api/copies/{copy_id}/ws", _websocket),
        Route("/api/brand", _brand, methods=["GET"]),
        Route("/api/{path:path}", _api_not_found),
        Route("/brand/{asset}", _brand_asset, methods=["GET"]),
        Route("/{path:path}", _static),
    ]
    app = Starlette(
        routes=routes,
        exception_handlers={AppError: _app_error, HTTPException: _http_error},
        lifespan=_lifespan,
    )
    app.state.services = services
    app.state.engine_pool = context.engine_pool
    app.add_middleware(JsonGZipMiddleware)
    app.add_middleware(SecurityBoundaryMiddleware)
    return app


@asynccontextmanager
async def _lifespan(app):
    try:
        yield
    finally:
        app.state.services.close()


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
    snapshot = services.snapshot
    template = {"name": snapshot.name, "sheets": len(snapshot.sheets),
                "visible": sum(not sheet["hidden"] for sheet in snapshot.sheets),
                "calculation": getattr(snapshot, "results", "engine")}
    return JSONResponse({"user": user, "csrf": csrf, "template": template})


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
    body = await asyncio.to_thread(services.workbooks.rows, query)
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
            await _announce(services, copy_id, ack["input_revision"])
    if created and services.engine_pool is not None:
        # the edit is saved; Excel can take a while, so the reply never waits for it
        _in_background(_calculate_latest(services, copy_id, ack["input_revision"]))
    return JSONResponse(ack)


async def _macros(request: Request):
    services, user, _ = _write_auth(request)
    copy_id, payload = request.path_params["copy_id"], await _body(request)
    button, op_id = payload.get("button"), payload.get("op_id")
    engine_kind = services.engine_pool.kind if services.engine_pool is not None else None
    macro = services.workbooks.macro_request(copy_id, user["id"], button, op_id, engine_kind)
    body = await _run_excel_action(
        services, copy_id, user["id"], {"op_id": op_id, "button": button}, "macro", _MACRO_FAILURES,
        lambda revision, source, changes, base, stored: services.engine_pool.run_macro(MacroRequest(
            copy_id, revision, source, changes, base, macro, pivot_state=stored)))
    return JSONResponse(body)


async def _refresh(request: Request):
    services, user, _ = _write_auth(request)
    copy_id, payload = request.path_params["copy_id"], await _body(request)
    op_id = payload.get("op_id")
    engine_kind = services.engine_pool.kind if services.engine_pool is not None else None
    services.workbooks.refresh_request(copy_id, user["id"], op_id, engine_kind)
    body = await _run_excel_action(
        services, copy_id, user["id"], {"op_id": op_id, "action": "refresh"}, "refresh", _REFRESH_FAILURES,
        lambda revision, source, changes, base, stored: services.engine_pool.run_macro(MacroRequest(
            copy_id, revision, source, changes, base, REFRESH_ALL, pivot_state=stored)),
        plan=lambda stored: None if stored["refreshed"] else with_refreshed(stored))
    return JSONResponse(body)


async def _slicers(request: Request):
    services, user, _ = _write_auth(request)
    copy_id, payload = request.path_params["copy_id"], await _body(request)
    op_id, cache = payload.get("op_id"), payload.get("cache")
    engine_kind = services.engine_pool.kind if services.engine_pool is not None else None
    selected = services.workbooks.slicer_request(copy_id, user["id"], op_id, cache, payload.get("selected"),
                                                 engine_kind)
    body = await _run_excel_action(
        services, copy_id, user["id"], {"op_id": op_id, "cache": cache, "selected": selected}, "slicer",
        _SLICER_FAILURES,
        lambda revision, source, changes, base, stored: services.engine_pool.run_slicer(SlicerRequest(
            copy_id, revision, source, changes, base, cache, selected, pivot_state=stored)),
        plan=lambda stored: with_selection(stored, cache, selected), skip_unchanged=True)
    return JSONResponse({**body, "selected": selected})


_MACRO_FAILURES = ("The macro ran too long and was stopped.", "The macro did not finish in Excel.")
_REFRESH_FAILURES = ("The refresh ran too long and was stopped.", "The refresh did not finish in Excel.")
_SLICER_FAILURES = ("The slicer change ran too long and was stopped.", "The slicer change did not finish in Excel.")


async def _run_excel_action(services, copy_id, user_id, operation, action, failures, call, plan=None,
                            skip_unchanged=False):
    """Run one Excel action (a macro, a refresh or a slicer change) and commit what it wrote as one revision.

    ``call(revision, source, changes, base, stored)`` runs it on the copy's engine with the stored pivot state.
    ``plan(stored)`` returns the pivot state the action stores, or None when it stores none. With
    ``skip_unchanged``, an action whose plan changes nothing is answered without running Excel.
    """
    async with services.hub.pipeline(copy_id):
        replay = services.operations.macro_replay(copy_id, user_id, operation)
        if replay is not None:
            return _macro_response(replay)
        state = services.operations.revisions(copy_id)
        revision = state["input_revision"] + 1
        source = services.workbooks.calculation_source(copy_id)
        changes = services.operations.revision_changes(copy_id, state["input_revision"])
        stored = services.workbooks.pivot_state(copy_id)
        target = plan(stored) if plan is not None else None
        if target == stored:
            target = None
        unchanged = {"messages": [], "changed": 0, "input_revision": state["input_revision"]}
        if skip_unchanged and target is None:
            return unchanged
        try:
            result = await asyncio.to_thread(call, revision, source, changes, state["calculated_revision"], stored)
        except MacroTimeout as exc:
            raise AppError(502, "macro_failed", failures[0]) from exc
        except EngineError as exc:
            raise AppError(502, "macro_failed", failures[1]) from exc
        ack = None
        if result["contents"] or target is not None:
            ack = await asyncio.to_thread(services.operations.apply_macro, copy_id, user_id, operation,
                                          result["messages"], result["contents"], revision, action, target)
        if ack is None:
            return {"messages": result["messages"], "changed": 0, "input_revision": state["input_revision"]}
        await _publish_macro(services, copy_id, revision, source, result, target or stored)
    return _macro_response(ack)


def _macro_response(ack):
    return {"messages": ack["messages"], "changed": ack["accepted"], "input_revision": ack["input_revision"]}


async def _publish_macro(services, copy_id, revision, source, result, pivot_state):
    """Publish a committed macro like an edit; a failure here leaves the calc job for the normal replay."""
    await services.hub.broadcast(copy_id, _event(services.database, copy_id, ("op", revision)))
    await services.hub.broadcast(copy_id, {"type": "calc_state", **_calc_state(services, copy_id)})
    try:
        request = CalculationRequest(copy_id, revision, source, services.operations.revision_changes(copy_id, revision),
                                     pivot_state=pivot_state)
        if await _publish_result(services, request, result):
            await services.hub.broadcast(copy_id, _event(services.database, copy_id, ("calc", revision)))
    except Exception as exc:
        logging.getLogger(__name__).error("Calculating copy %s revision %s failed: %s", copy_id, revision, exc)
        await asyncio.to_thread(services.calculations.failed, copy_id, revision, type(exc).__name__)


async def _announce(services, copy_id, revision):
    """Tell browsers about a saved edit; a values-only template is then already current."""
    event = _event(services.database, copy_id, ("op", revision))
    await services.hub.broadcast(copy_id, event)
    state = _calc_state(services, copy_id)
    await services.hub.broadcast(copy_id, {"type": "calc_state", **state})
    if services.engine_pool is None and state["calculated_revision"] >= revision:
        await services.hub.broadcast(copy_id, {"type": "calc", "from_calculated_revision": revision - 1,
                                               "calculated_revision": revision, "changes": []})


def _failure(exc):
    """The reason a person sees: the error type and its first line, never a traceback."""
    detail = str(exc).strip().splitlines()[0] if str(exc).strip() else ""
    return f"{type(exc).__name__}: {detail}"[:200] if detail else type(exc).__name__


_CALC_LOCKS = {}  # copy id -> the lock its background calculations take turns on


async def _calculate_latest(services, copy_id, revision):
    """Calculate the copy's newest saved state once; edits saved meanwhile are folded into the next run.

    Edits keep saving while it runs: only publishing the result takes the copy's pipeline lock.
    """
    lock = _CALC_LOCKS.setdefault(copy_id, asyncio.Lock())
    async with lock:
        state = await asyncio.to_thread(_calc_state, services, copy_id)
        if state["calculated_revision"] >= revision:
            return  # an earlier run already calculated a state that includes this edit
        latest = state["input_revision"]
        try:
            source = services.workbooks.calculation_source(copy_id)
            changes = services.operations.revision_changes(copy_id, latest)
            request = CalculationRequest(copy_id, latest, source, changes, state["calculated_revision"],
                                         pivot_state=services.workbooks.pivot_state(copy_id))
            result = await asyncio.to_thread(services.engine_pool.apply, request)
            async with services.hub.pipeline(copy_id):
                if await _publish_result(services, request, result):
                    await services.hub.broadcast(copy_id, _event(services.database, copy_id, ("calc", latest)))
        except Exception as exc:
            logging.getLogger(__name__).exception("Calculating copy %s revision %s failed", copy_id, latest)
            await asyncio.to_thread(services.calculations.failed, copy_id, latest, _failure(exc))
            await services.hub.broadcast(copy_id, {"type": "calc_state", **_calc_state(services, copy_id)})


async def _calculate_baseline(services, copy, user_id):
    if services.engine_pool is None:
        return copy
    async with services.hub.pipeline(copy["id"]):
        try:
            source = services.workbooks.calculation_source(copy["id"])
            changes = services.operations.revision_changes(copy["id"], 0)
            request = CalculationRequest(
                copy["id"], 0, source, changes, copy["calculated_revision"],
                pivot_state=services.workbooks.pivot_state(copy["id"]),
            )
            published = await _run_calculation(services, request)
            if published:
                event = _event(services.database, copy["id"], ("calc", 0))
                await services.hub.broadcast(copy["id"], event)
                state = _calc_state(services, copy["id"])
                await services.hub.broadcast(copy["id"], {"type": "calc_state", **state})
        except Exception as exc:
            logging.getLogger(__name__).exception("Calculating copy %s failed", copy["id"])
            await asyncio.to_thread(services.calculations.failed, copy["id"], 0, _failure(exc))
            await services.hub.broadcast(copy["id"], {"type": "calc_state", **_calc_state(services, copy["id"])})
            raise
    return services.copies.get(user_id, copy["id"])


async def _run_calculation(services, request):
    result = await asyncio.to_thread(services.engine_pool.apply, request)
    return await _publish_result(services, request, result)


async def _publish_result(services, request, result):
    publication = CalculationPublication.from_result(
        request.copy_id, request.revision, result
    )
    try:
        return await asyncio.to_thread(services.calculations.complete, publication)
    except PublicationBaseMismatch:
        replay = replace(request)
        result = await asyncio.to_thread(services.engine_pool.apply, replay)
        publication = CalculationPublication.from_result(
            request.copy_id, request.revision, result
        )
        return await asyncio.to_thread(services.calculations.complete, publication)


async def _export(request: Request):
    _reject_head(request)
    services, user = _read_auth(request)
    admission = _admit_transfer(services)
    try:
        copy_id = request.path_params["copy_id"]
        snapshot = services.workbooks.export_snapshot(copy_id, user["id"])
        selected = _selected_sheets(request.query_params, snapshot["sheets"])
        _require_export_ready(snapshot, selected)
        if request.query_params.get("check") == "1":
            return Response(status_code=204)
        fmt = "xlsx" if request.url.path.endswith(".xlsx") else "xlsm"
        path = _export_path(services, fmt, False)
        await _build_export(ExportJob(
            snapshot, copy_id, fmt, selected, path,
        ), services.transfer_limiter)
        return DisposableFileResponse(
            path, media_type=_export_media(fmt), filename=f"price-plan.{fmt}",
        )
    finally:
        admission.release()


async def _prepare_download(request: Request):
    services, user, _ = _write_auth(request)
    admission = _admit_transfer(services)
    handed_off = False
    try:
        payload = await _body(request)
        fmt = _download_format(payload.get("format"))
        copy_id = request.path_params["copy_id"]
        snapshot = services.workbooks.export_snapshot(copy_id, user["id"])
        selected = _selected_values(payload.get("sheets"), snapshot["sheets"])
        _require_export_ready(snapshot, selected)
        path = _export_path(services, fmt, True)
        session = token_hash(request.cookies.get("pp_session", ""))
        job = ExportJob(snapshot, copy_id, fmt, selected, path)

        def publish():
            artifact = PreparedArtifact(
                path, user["id"], session,
                copy_id, f"price-plan.{fmt}", time.monotonic() + PREPARED_DOWNLOAD_SECONDS,
            )
            return {"download_url": f"/api/downloads/{services.downloads.add(artifact)}"}

        if request.query_params.get("async") == "1":
            def work(progress):
                progress.progress("Preparing the workbook")
                try:
                    _call_bounded(services.transfer_limiter, _write_export_job, (job, threading.Event()))
                except ValueError as exc:
                    raise _export_error(exc) from exc
                except BaseException:
                    path.unlink(missing_ok=True)
                    raise
                return publish()

            def cleanup():
                admission.release()

            token = services.jobs.start(user["id"], session, work, cleanup)
            handed_off = True
            return JSONResponse({"job": token}, status_code=202)
        await _build_export(job, services.transfer_limiter)
        return JSONResponse(publish(), status_code=201)
    finally:
        if not handed_off:
            admission.release()


async def _prepared_download(request: Request):
    _reject_head(request)
    services = _services(request)
    token = request.path_params["token"]
    artifact = None
    try:
        user, _ = services.auth.authenticate(request.cookies.get("pp_session"))
        artifact = services.downloads.take(
            token, user["id"], token_hash(request.cookies.get("pp_session", "")),
        )
        services.copies.get(user["id"], artifact.copy_id)
        return DisposableFileResponse(
            artifact.path, media_type=_export_media(artifact.path.suffix[1:]),
            filename=artifact.filename,
        )
    except BaseException:
        if artifact is None:
            services.downloads.discard(token)
        else:
            artifact.path.unlink(missing_ok=True)
        raise


async def _build_export(job, limiter=None):
    cancelled = threading.Event()
    try:
        await _run_bounded(
            limiter or threading.BoundedSemaphore(1), _write_export_job, (job, cancelled),
        )
    except ValueError as exc:
        job.path.unlink(missing_ok=True)
        raise _export_error(exc) from exc
    except BaseException:
        cancelled.set()
        job.path.unlink(missing_ok=True)
        raise


def _export_error(exc):
    if str(exc).startswith("current calculation result"):
        return AppError(409, "calculation_required", str(exc))
    return exc


def _write_export(job):
    cells = [(*coordinate, value)
             for coordinate, value in sorted(job.snapshot["overrides"].items())]
    export_transfer(TransferRequest(
        template_path=job.snapshot["source"], cells=cells, out_path=job.path,
        template_hash=job.snapshot["template_id"], copy_id=job.copy_id,
        input_revision=job.snapshot["input_revision"], fmt=job.fmt,
        selected=job.selected, results=job.snapshot["results"], extras_id=job.snapshot.get("extras_id") or "",
    ))
    if job.snapshot.get("order"):
        from .export.extras_graft import reorder_sheets
        reorder_sheets(job.path, job.snapshot["order"])  # the copy's own tabs go back where the user put them


def _write_export_job(payload):
    job, cancelled = payload
    try:
        _write_export(job)
    finally:
        if cancelled.is_set():
            job.path.unlink(missing_ok=True)


def _call_bounded(limiter, function, value):
    with limiter:
        return function(value)


async def _run_bounded(limiter, function, value):
    worker = asyncio.create_task(asyncio.to_thread(_call_bounded, limiter, function, value))
    try:
        return await asyncio.shield(worker)
    except asyncio.CancelledError:
        try:
            await worker
        except BaseException:
            pass
        raise


def _require_export_ready(snapshot, selected):
    if (selected is not None and len(selected) < len(snapshot["sheets"])
            and snapshot["calculated_revision"] != snapshot["input_revision"]):
        raise AppError(
            409, "calculation_required",
            "Workbook calculation is pending; wait for current results before downloading a sheet subset",
        )


def _download_format(value):
    if value not in {"xlsx", "xlsm"}:
        raise AppError(422, "invalid_selection", "Download format is invalid")
    return value


def _export_path(services, fmt, prepared):
    directory = Path(services.config.data_dir) / "exports"
    if prepared:
        directory /= "prepared"
    directory.mkdir(parents=True, exist_ok=True)
    return directory / f"price-plan-{uuid4().hex}.{fmt}"


def _export_media(fmt):
    return ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if fmt == "xlsx"
            else "application/vnd.ms-excel.sheet.macroenabled.12")


def _declared_length(request) -> int:
    raw = request.headers.get("content-length", "")
    if not raw:
        return 0
    try:
        length = int(raw)
    except ValueError as exc:
        raise AppError(400, "invalid_length", "Content-Length header is not a number") from exc
    return max(length, 0)


def _reject_head(request) -> None:
    if request.method == "HEAD":
        raise AppError(405, "method_not_allowed", "Use GET to download this file")


async def _upload_preview(request: Request):
    services, user, _ = _write_auth(request)
    admission = _admit_transfer(services)
    handoff = []
    try:
        return await _admitted_upload_preview(request, services, user, handoff)
    finally:
        if not handoff:
            admission.release()


async def _job_status(request: Request):
    services = _services(request)
    user, _ = services.auth.authenticate(request.cookies.get("pp_session"))
    return JSONResponse(services.jobs.get(
        request.path_params["token"], user["id"], token_hash(request.cookies.get("pp_session", "")),
    ))


async def _admitted_upload_preview(request, services, user, handoff):
    limit = getattr(services.config, "upload_max_mb", 100) * 1024 * 1024
    length = _declared_length(request)
    if length > limit:
        raise AppError(413, "too_large", "Upload exceeds the configured size limit")
    directory = Path(services.config.data_dir) / "uploads"
    directory.mkdir(parents=True, exist_ok=True)
    handle = tempfile.NamedTemporaryFile(prefix="preview-", suffix=".xlsx", dir=directory, delete=False)
    path = Path(handle.name)
    try:
        size = 0
        with handle:
            async for chunk in request.stream():
                size += len(chunk)
                if size > limit:
                    raise AppError(413, "too_large", "Upload exceeds the configured size limit")
                handle.write(chunk)
        if size == 0:
            raise AppError(422, "invalid_workbook", "Upload is empty")
        preview = PreviewRequest(
            path=path,
            filename=unquote(request.headers.get("x-upload-filename", "")),
            user_id=user["id"],
            session_hash=token_hash(request.cookies.get("pp_session", "")),
            target_copy_id=(request.path_params.get("copy_id")
                            or request.query_params.get("copy_id")),
            category_id=request.query_params.get("category_id"),
        )
        if request.query_params.get("async") == "1":
            def work(job):
                return _call_bounded(services.transfer_limiter, services.uploads.preview,
                                     replace(preview, progress=job.progress))

            def cleanup():
                path.unlink(missing_ok=True)
                services.transfer_admission.release()

            token = services.jobs.start(user["id"], preview.session_hash, work, cleanup)
            handoff.append(token)
            return JSONResponse({"job": token}, status_code=202)
        body = await _run_bounded(services.transfer_limiter, services.uploads.preview, preview)
        return JSONResponse(body)
    finally:
        if not handoff:
            path.unlink(missing_ok=True)


def _admit_transfer(services):
    if not services.transfer_admission.acquire(blocking=False):
        raise AppError(503, "transfer_busy", "Workbook transfers are busy; try again shortly")
    return services.transfer_admission


async def _upload_confirm(request: Request):
    services, user, _ = _write_auth(request)
    payload = await _body(request)
    mode = payload.get("mode")
    confirmation = ConfirmRequest(
        token=request.path_params["token"], payload=payload, user_id=user["id"],
        session_hash=token_hash(request.cookies.get("pp_session", "")),
        session_check=_session_check(request, services),
    )
    body = await asyncio.to_thread(services.uploads.confirm, confirmation)
    # The upload is saved now; recalculating a whole uploaded workbook can take Excel minutes, so it runs
    # after the reply and browsers follow it through the copy's calc_state and calc messages.
    if body.get("accepted"):
        copy_id = body.get("copy_id") or body["copy"]["id"]
        async with services.hub.pipeline(copy_id):
            await _announce(services, copy_id, body["input_revision"])
        if services.engine_pool is not None:
            _in_background(_calculate_latest(services, copy_id, body["input_revision"]))
    elif mode == "new_copy":
        _in_background(_calculate_baseline(services, body["copy"], user["id"]))
    return JSONResponse(body, status_code=201 if mode == "new_copy" else 200)


_BACKGROUND = set()  # running background calculations, held so they are not garbage collected


def _in_background(coroutine):
    task = asyncio.get_running_loop().create_task(coroutine)
    _BACKGROUND.add(task)

    def done(finished):
        _BACKGROUND.discard(finished)
        if not finished.cancelled() and finished.exception() is not None:
            logging.getLogger(__name__).error("Background calculation failed: %s", finished.exception())

    task.add_done_callback(done)
    return task


def _selected_sheets(query, sheets):
    values = query.getlist("sheets")
    if not values:
        return None
    known = {sheet["name"] for sheet in sheets}
    selected = []
    try:
        for value in values:
            if value.startswith("["):
                decoded = json.loads(value)
                if not isinstance(decoded, list) or not all(isinstance(item, str) for item in decoded):
                    raise ValueError
                selected.extend(decoded)
            elif value in known:
                selected.append(value)
            else:
                selected.extend(value.split(","))
    except (TypeError, ValueError, json.JSONDecodeError) as exc:
        raise AppError(422, "invalid_selection", "Sheet selection is invalid") from exc
    return _validate_selection(selected, sheets)


def _selected_values(values, sheets):
    if not isinstance(values, list) or not all(isinstance(item, str) for item in values):
        raise AppError(422, "invalid_selection", "Sheet selection is invalid")
    return _validate_selection(values, sheets)


def _validate_selection(values, sheets):
    selected = set(values)
    known = {sheet["name"] for sheet in sheets}
    if not selected or selected - known:
        raise AppError(422, "invalid_selection", "Select known workbook sheets")
    chosen = [sheet for sheet in sheets if sheet["name"] in selected]
    if not any(sheet["hidden"] not in {"hidden", "veryHidden"} for sheet in chosen):
        raise AppError(422, "invalid_selection", "Select at least one visible sheet")
    return selected


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
    length = _declared_length(request)
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


def _calc_state(services, copy_id):
    with services.database.connect() as connection:
        row = connection.execute("SELECT input_revision,calculated_revision,template_id FROM pp_copies WHERE id=?",
                                 (copy_id,)).fetchone()
        snapshot = services.workbooks.templates.get(connection, row["template_id"])
        error = calc_error(connection, copy_id, row["input_revision"])
    # same rule as rows and workbook reads: a values-only template is always current
    calculated = _calculated_revision(snapshot, row)
    return {"input_revision": row["input_revision"], "calculated_revision": calculated,
            "error": error if calculated < row["input_revision"] else None}


async def _app_error(request, exc: AppError):
    return JSONResponse({"error": exc.code, "message": exc.message, **exc.extra}, status_code=exc.status)


async def _http_error(request, exc: HTTPException):
    code = "not_found" if exc.status_code == 404 else "method_not_allowed" if exc.status_code == 405 else "http_error"
    return JSONResponse({"error": code, "message": str(exc.detail)}, status_code=exc.status_code)


async def _api_not_found(request):
    raise AppError(404, "not_found", "Endpoint not found")


async def _brand(request):
    brand = getattr(_services(request).config, "brand", None)
    if not brand:
        return JSONResponse({})
    return JSONResponse({"name": brand["name"], "colors": brand["colors"],
                         "assets": sorted(brand["assets"])})


async def _brand_asset(request):
    brand = getattr(_services(request).config, "brand", None) or {}
    path = brand.get("assets", {}).get(request.path_params["asset"])
    if path is None:
        raise AppError(404, "not_found", "File not found")
    return FileResponse(path, media_type=BRAND_TYPES[path.suffix.lower()])


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


def _sweep_transfer_files(data_dir):
    root = Path(data_dir)
    for directory in (root / "exports", root / "uploads" / "previews"):
        if not directory.is_dir():
            continue
        for path in directory.rglob("*"):
            if path.is_file():
                path.unlink(missing_ok=True)
    for path in (root / "uploads" / "previews").glob("extras-*"):
        shutil.rmtree(path, ignore_errors=True)  # an unconfirmed preview's own tabs and columns
    uploads = root / "uploads"
    if uploads.is_dir():
        for path in uploads.glob("preview-*.xlsx"):
            path.unlink(missing_ok=True)


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


class PreparedDownloadStore:
    def __init__(self):
        self._items = {}
        self._lock = threading.Lock()

    def add(self, artifact):
        token = uuid4().hex
        try:
            size = artifact.path.stat().st_size
            if size > MAX_PREPARED_DOWNLOAD_BYTES:
                raise AppError(
                    413, "download_too_large", "Prepared download exceeds the safe disk limit"
                )
            with self._lock:
                self._purge()
                while (len(self._items) >= MAX_PREPARED_DOWNLOADS
                       or self._stored_bytes() + size > MAX_PREPARED_DOWNLOAD_BYTES):
                    removed = self._items.pop(next(iter(self._items)))
                    removed.path.unlink(missing_ok=True)
                self._items[token] = artifact
        except BaseException:
            artifact.path.unlink(missing_ok=True)
            raise
        return token

    def take(self, token, user_id, session_hash):
        with self._lock:
            self._purge()
            artifact = self._items.pop(token, None)
        if artifact is None:
            raise AppError(404, "download_expired", "Download expired; prepare it again")
        if artifact.user_id != user_id or artifact.session_hash != session_hash:
            artifact.path.unlink(missing_ok=True)
            raise AppError(404, "download_expired", "Download expired; prepare it again")
        return artifact

    def close(self):
        with self._lock:
            artifacts = tuple(self._items.values())
            self._items.clear()
        for artifact in artifacts:
            artifact.path.unlink(missing_ok=True)

    def discard(self, token):
        with self._lock:
            artifact = self._items.pop(token, None)
        if artifact is not None:
            artifact.path.unlink(missing_ok=True)

    def _purge(self):
        now = time.monotonic()
        for token, artifact in tuple(self._items.items()):
            if artifact.expires <= now:
                del self._items[token]
                artifact.path.unlink(missing_ok=True)

    def _stored_bytes(self):
        return sum(artifact.path.stat().st_size for artifact in self._items.values())


class DisposableFileResponse(FileResponse):
    def __init__(self, path, **kwargs):
        self.cleanup_path = Path(path)
        super().__init__(path, **kwargs)

    async def __call__(self, scope, receive, send):
        try:
            await super().__call__(scope, receive, send)
        finally:
            self.cleanup_path.unlink(missing_ok=True)


class JsonGZipMiddleware:
    """Gzip application/json responses only.

    Starlette's own middleware compresses anything large enough, which would gzip the .xlsm and
    .xlsx downloads: they are already zip containers, so that costs time and CPU and saves
    nothing. A response head is held back until its first body chunk, which is where the content
    type is known, and streaming responses are passed through untouched.
    """

    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http" or "gzip" not in Headers(scope=scope).get("accept-encoding", ""):
            await self.app(scope, receive, send)
            return
        await _JsonGZipResponder(self.app)(scope, receive, send)


class _JsonGZipResponder:
    def __init__(self, app):
        self.app = app
        self.head = None
        self.compress = False

    async def __call__(self, scope, receive, send):
        self.send = send
        await self.app(scope, receive, self.relay)

    async def relay(self, message):
        if message["type"] == "http.response.start":
            self.head = message
            headers = Headers(raw=message["headers"])
            self.compress = (headers.get("content-type", "").split(";")[0].strip() == JSON_MEDIA_TYPE
                             and "content-encoding" not in headers)
            return
        head = self.head
        if head is None or message["type"] != "http.response.body":
            await self.send(message)
            return
        self.head = None
        body = message.get("body", b"")
        if not self.compress or message.get("more_body", False) or len(body) < GZIP_MINIMUM_BYTES:
            await self.send(head)
            await self.send(message)
            return
        buffer = io.BytesIO()
        with gzip.GzipFile(mode="wb", fileobj=buffer, compresslevel=GZIP_LEVEL) as archive:
            archive.write(body)
        compressed = buffer.getvalue()
        headers = MutableHeaders(raw=head["headers"])
        headers.add_vary_header("Accept-Encoding")
        headers["Content-Encoding"] = "gzip"
        headers["Content-Length"] = str(len(compressed))
        await self.send(head)
        await self.send({**message, "body": compressed})


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
