"""Bounded workbook upload inspection, preview, and confirmation."""

from __future__ import annotations

from dataclasses import dataclass, field, replace
import hashlib
from io import StringIO
from concurrent.futures import ProcessPoolExecutor, as_completed
from concurrent.futures.process import BrokenProcessPool
import json
import logging
import multiprocessing
import os
from pathlib import Path, PurePosixPath
import re
import secrets
import shutil
import tempfile
import threading
import time
from uuid import uuid4
import zipfile
import xml.etree.ElementTree as ET
import defusedxml.ElementTree as SafeET
import xml.parsers.expat as expat

from openpyxl.formula.translate import Translator
from openpyxl.utils.escape import unescape as unescape_ooxml

from .authorization import effective_role
from .errors import AppError
from .extras import ExtrasPlan, attach, build_extras, current_plan, edited_since
from .export.transfer import (
    REL_ATTRIBUTE_NAMESPACES,
    LinkScan,
    _local,
    _relationship_graph,
    _relationship_references,
    _sheet_records,
    open_check,
    validate_package,
)
from .ops import EDIT_ROLES, UploadOperationRequest
from .upload_align import FORMULA, MASK, align_rows, formula_column_hash
from .values import exactly_equal, validate_value

LOGGER = logging.getLogger(__name__)


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XLSX_MAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
XLSM_MAIN = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
MAX_MEMBERS = 10_000
# The Spring 2027 template unpacks to ~650 MB (one sheet is 191 MB of XML); the ratio guard catches bombs.
MAX_UNCOMPRESSED = 2 * 1024 * 1024 * 1024
MAX_RATIO = 200
RATIO_MIN_BYTES = 1024 * 1024
# Workbooks whose sheets unpack past this are read one sheet per worker process, all cores at once
PARALLEL_MIN_BYTES = 32 * 1024 * 1024
MAX_SCAN_WORKERS = 4
MAX_TREE_XML = 8 * 1024 * 1024
MAX_RELATIONSHIP_XML = 4 * 1024 * 1024
MAX_RELATIONSHIP_XML_TOTAL = 16 * 1024 * 1024
MAX_CUSTOM_PROPERTIES_XML = 1024 * 1024
MAX_STYLES_XML = 16 * 1024 * 1024
MAX_SHARED_STRINGS_XML = 64 * 1024 * 1024
MAX_THEME_XML = 4 * 1024 * 1024
MAX_SHARED_STRING_COUNT = 500_000
MAX_SHARED_STRING_BYTES = 32 * 1024 * 1024
MAX_STRUCTURE_RECORDS = 1_000_000  # the real template has ~157k styled rows and columns
MAX_STRUCTURE_BYTES = 64 * 1024 * 1024
MAX_XML_DEPTH = 256
MAX_XML_TOKEN_BYTES = 8 * 1024 * 1024
PREVIEW_SECONDS = 15 * 60
MAX_ACTIVE_PREVIEWS = 128
MAX_PREVIEW_CACHE_BYTES = 1024 * 1024 * 1024
XML_RELATIONSHIP_ROLES = {
    "officeDocument", "worksheet", "styles", "sharedStrings", "theme", "custom-properties",
    "chart", "chartsheet", "chartSheet", "drawing",
}


@dataclass(frozen=True)
class WorkbookModel:
    sheets: tuple[str, ...]
    cells: dict[str, dict[str, dict]]
    structures: dict[str, tuple]
    fingerprint: str
    provenance: dict | None
    ignored: tuple[dict, ...]
    displaced: frozenset[str]
    # Highest row and column holding a cell, per sheet, measured the same way for template and upload
    extents: dict[str, tuple[int, int]] = field(default_factory=dict)
    # Per sheet {row: sum of formula_column_hash over the row's formula cells}: equal maps mean no
    # formula moved, so rows were only filled or added below
    formula_rows: dict[str, dict[int, int]] = field(default_factory=dict)
    parts: dict[str, str] = field(default_factory=dict)  # sheet name -> worksheet part, for row matching
    strings_part: str | None = None
    filled_cols: dict[str, int] = field(default_factory=dict)  # one past the last filled column
    suspect: frozenset[str] = frozenset()  # sheets with values in locked empty cells
    digests: dict[str, str] = field(default_factory=dict)  # per-sheet formula hash, when fingerprinted
    columns: dict[str, dict[int, int]] = field(default_factory=dict)  # filled cells per column
    headers: dict[str, dict[int, str]] = field(default_factory=dict)  # a name for each column
    filled_rows: dict[str, int] = field(default_factory=dict)
    # Per template sheet, the first column past the template's own: cells there are the user's extras
    limits: dict[str, int] = field(default_factory=dict)


@dataclass(frozen=True)
class UploadDependencies:
    database: object
    templates: object
    copies: object
    operations: object
    preview_dir: Path
    converter: object = None  # converter(source, target): Excel's Save As .xlsm, on the Excel engine only


@dataclass(frozen=True)
class PreviewRequest:
    path: Path
    filename: str
    user_id: str
    session_hash: str
    target_copy_id: str | None = None
    category_id: str | None = None
    progress: object = None  # progress(stage, done, total), for a background upload job


@dataclass(frozen=True)
class ConfirmRequest:
    token: str
    payload: dict
    user_id: str
    session_hash: str
    session_check: object = None


@dataclass(frozen=True)
class PreviewContext:
    connection: object
    model: WorkbookModel
    template_model: WorkbookModel
    snapshot: object
    copy: object
    user_id: str
    revision: int | None
    path: Path | None = None  # the uploaded file, read again for sheets whose rows moved


@dataclass(frozen=True)
class SheetContext:
    snapshot: object
    sheet: int
    name: str
    overrides: dict
    ignored: list
    limit: int | None = None  # first extra column; the copy's values there belong to its extras


@dataclass(frozen=True)
class InputValues:
    old: object
    new: object
    template: object
    version: int


@dataclass(frozen=True)
class CellIdentity:
    sheet_name: str
    reference: str
    coordinate: tuple


@dataclass(frozen=True)
class HistoryRequest:
    copy_id: str
    revision: int | None


@dataclass(frozen=True)
class PreviewEntry:
    path: Path
    user_id: str
    session_hash: str
    expires: float
    size: int
    extras: str | None = None  # the extras folder built for this preview, removed with it


@dataclass(frozen=True)
class SheetScan:
    archive: object
    part: str
    strings: list
    snapshot: object
    index: int | None
    name: str
    all_formulas: bool
    keep: frozenset = frozenset()  # (row, col) of the copy's saved edits on this sheet
    moved: bool = False  # rows were matched to template rows: shifted formulas are not differences
    limit: int | None = None  # first extra column: cells from here on are the user's own, not inputs


@dataclass(frozen=True)
class SheetData:
    cells: dict
    formulas: dict
    ignored_count: int
    displaced: bool
    extent: tuple[int, int] = (0, 0)
    formula_rows: dict = field(default_factory=dict)
    filled_cols: int = 0
    suspect: bool = False  # a locked empty cell holds a value: rows or columns may have moved
    columns: dict = field(default_factory=dict)
    headers: dict = field(default_factory=dict)
    filled_rows: int = 0


@dataclass
class DisplacementDetector:
    snapshot: object
    sheet: int | None
    rows: int = 0
    cols: int = 0
    locked_layout_changed: bool = False
    formula_rows: dict = field(default_factory=dict)
    filled_cols: int = 0  # one past the last column holding a value or formula
    locked_filled: bool = False
    columns: dict = field(default_factory=dict)  # col -> filled cells, for the file/template comparison
    headers: dict = field(default_factory=dict)  # col -> first text near the top, to name the column
    filled_rows: int = 0
    last_filled_row: int = -1

    def observe(self, record, extra=False):
        self.rows = max(self.rows, record["row"] + 1)
        self.cols = max(self.cols, record["col"] + 1)
        if record["formula"] is not None or record["value"] not in (None, ""):
            col = record["col"]
            self.filled_cols = max(self.filled_cols, col + 1)
            self.columns[col] = self.columns.get(col, 0) + 1
            if record["row"] != self.last_filled_row:
                self.filled_rows, self.last_filled_row = self.filled_rows + 1, record["row"]
            if (col not in self.headers and record["formula"] is None and isinstance(record["value"], str)
                    and record["row"] < HEADER_ROWS):
                self.headers[col] = record["value"][:60]
        if record["formula"] is not None and not extra:
            row = record["row"]
            self.formula_rows[row] = (self.formula_rows.get(row, 0) + formula_column_hash(record["col"])) & MASK

    def locked_value_changed(self, record, expected):
        """A changed value in a locked, formula-free cell means rows or columns moved under it. A value
        in a cell the template leaves empty only may mean that (rows added below, or a column inserted),
        so it sends the sheet to row matching instead."""
        if record["formula"] is None and expected.get("formula") is None:
            if expected.get("v") in (None, ""):
                self.locked_filled = True
            else:
                self.locked_layout_changed = True

    def detected(self):
        # The extent is compared with the template's own cell extent in _preview_details: the
        # snapshot's sheet size also counts formatted empty columns, so it never matched a real file
        return self.snapshot is not None and self.sheet is not None and self.locked_layout_changed


class PreviewStore:
    def __init__(self, directory):
        self.directory = Path(directory)
        self.directory.mkdir(parents=True, exist_ok=True)
        self._items = {}
        self._lock = threading.Lock()

    def put(self, value):
        token = secrets.token_urlsafe(24)
        path = self._write(token, value)
        size = path.stat().st_size
        if size > MAX_PREVIEW_CACHE_BYTES:
            path.unlink(missing_ok=True)
            raise AppError(413, "preview_too_large", "Upload preview exceeds the safe limit")
        with self._lock:
            self._purge()
            self._make_room(value["user_id"], value["session_hash"], size)
            self._items[token] = PreviewEntry(
                path, value["user_id"], value["session_hash"],
                time.monotonic() + PREVIEW_SECONDS, size, value.get("extras_dir"),
            )
        return token

    def take(self, token, user_id, session_hash):
        with self._lock:
            self._purge()
            value = self._items.pop(token, None)
        if value is None:
            raise AppError(404, "preview_expired", "Upload preview expired; preview the file again")
        try:
            if value.user_id != user_id or value.session_hash != session_hash:
                _remove_extras(value)
                raise AppError(404, "preview_expired", "Upload preview expired; preview the file again")
            return json.loads(value.path.read_text(encoding="utf-8"))  # the caller now owns the extras
        finally:
            value.path.unlink(missing_ok=True)

    def close(self):
        with self._lock:
            for token in tuple(self._items):
                self._delete(token)

    def _write(self, token, value):
        handle = tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", prefix=f"{token}-", suffix=".json",
            dir=self.directory, delete=False,
        )
        try:
            with handle:
                json.dump(value, handle, separators=(",", ":"), allow_nan=False)
            return Path(handle.name)
        except BaseException:
            Path(handle.name).unlink(missing_ok=True)
            raise

    def _purge(self):
        now = time.monotonic()
        for token, value in tuple(self._items.items()):
            if value.expires <= now:
                self._delete(token)

    def _make_room(self, user_id, session_hash, size):
        for token, value in tuple(self._items.items()):
            if value.user_id == user_id and value.session_hash == session_hash:
                self._delete(token)
        while (len(self._items) >= MAX_ACTIVE_PREVIEWS
               or sum(item.size for item in self._items.values()) + size > MAX_PREVIEW_CACHE_BYTES):
            if not self._items:
                break
            self._delete(next(iter(self._items)))

    def _delete(self, token):
        value = self._items.pop(token, None)
        if value is not None:
            value.path.unlink(missing_ok=True)
            _remove_extras(value)


def _remove_extras(entry):
    if entry.extras:
        shutil.rmtree(entry.extras, ignore_errors=True)


class UploadService:
    def __init__(self, dependencies):
        self.database = dependencies.database
        self.templates = dependencies.templates
        self.copies = dependencies.copies
        self.operations = dependencies.operations
        self.previews = PreviewStore(dependencies.preview_dir)
        self.converter = dependencies.converter
        self.data_dir = Path(dependencies.preview_dir).parents[1]
        self._template_models = {}
        self._input_cols = {}  # template hash -> {sheet index: one past its rightmost input column}

    def preview(self, request):
        path = self._readable(request)
        try:
            return self._preview(replace(request, path=path))
        finally:
            if path != request.path:
                path.unlink(missing_ok=True)

    def _readable(self, request):
        """The upload as an .xlsx/.xlsm package: .xlsb and .xls go through Excel's Save As when the app
        runs on Excel; otherwise the message says what the file is and how to save it."""
        kind = _workbook_format(request.path)
        if kind == "ooxml":
            return request.path
        if kind in {"xlsb", "ole"} and self.converter is not None:
            (request.progress or _no_progress)("Converting the file with Excel")
            source = request.path.with_name(request.path.stem + (".xlsb" if kind == "xlsb" else ".xls"))
            target = request.path.with_name(request.path.stem + "-converted.xlsm")
            try:
                source.write_bytes(request.path.read_bytes())  # Excel trusts the extension
                self.converter(source, target)
                return target
            except Exception as exc:
                LOGGER.warning("Excel could not convert an uploaded %s file: %s", kind, type(exc).__name__)
                target.unlink(missing_ok=True)
            finally:
                source.unlink(missing_ok=True)
        raise AppError(422, "unsupported_workbook", UNSUPPORTED_MESSAGES[kind])

    def _preview(self, request):
        progress = request.progress or _no_progress
        filename = _filename(request.filename)
        progress("Checking the file is safe")
        deferred = _deferred_parts(request.path)
        _validate_archive(request.path, deferred)
        # Only the provenance tag picks the copy and template; the sheets are read once, below
        provenance = _read_provenance(request.path)
        with self.database.read_transaction() as connection:
            copy, role = self._target_copy(connection, provenance, request)
            snapshot = (self.templates.get(connection, copy["template_id"]) if copy is not None
                        else self._new_copy_snapshot(connection, provenance, request))
            edited = frozenset(_overrides(connection, copy["id"])) if copy is not None else frozenset()
            template_model = self._template_model(snapshot)
            inputs = self._input_limits(snapshot)
            limits = {name.casefold(): max(template_model.filled_cols.get(name, 0), inputs.get(index, 0))
                      for index, name in enumerate(template_model.sheets)}
            model = _read_workbook(request.path, snapshot=snapshot, fingerprint=provenance is None,
                                   progress=progress, overrides=edited, deferred=deferred,
                                   column_limits=limits)
            progress("Comparing with the copy")
            identification, revision = self._identify(model, snapshot, copy)
            details = _preview_details(PreviewContext(
                connection, model, self._template_model(snapshot), snapshot,
                copy, request.user_id, revision, request.path,
            ))
            if identification == "structure":
                details["ignored"].extend(_formula_note(model, self._template_model(snapshot)))
            structure = _structure_report(model, self._template_model(snapshot))
            kept = current_plan(connection, copy["id"]) if copy is not None else None
            stale = kept is not None and edited_since(connection, copy["id"], revision)
        plan = ExtrasPlan(tuple(details["extras"]["sheets"]), details["extras"]["columns"], model.sheets)
        built, replace_extras = None, kept is not None
        if stale:
            # The file would undo edits made in the app since it was downloaded: keep the app's version
            details["ignored"].append({"sheet": None, "cell": None, "reason":
                                       "Your own tabs and columns were changed in the app after this file was "
                                       "downloaded, so the copy keeps the app's version of them. Download the "
                                       "copy again to change them in Excel"})
            plan, replace_extras, kept = ExtrasPlan((), {}, model.sheets), False, None
        if not plan.empty():
            progress("Reading your own tabs and columns")
            try:
                built, replace_extras = build_extras(request.path, plan, snapshot, self.previews.directory), True
                lost = json.loads((built / "extras.json").read_text(encoding="utf-8"))["counts"]["dropped_objects"]
                if lost:
                    details["ignored"].append({"sheet": None, "cell": None, "reason":
                                               f"{lost} charts, pictures, comments or tables in your own tabs are "
                                               "not kept; their cells are"})
            except Exception as exc:  # the input changes still apply; the copy keeps what it had
                LOGGER.exception("Reading an upload's own tabs and columns failed")
                details["ignored"].append({"sheet": None, "cell": None, "reason":
                                           "Your own tabs and columns could not be read, so this copy keeps "
                                           f"the ones it had ({type(exc).__name__})"})
                plan, replace_extras = ExtrasPlan((), {}, model.sheets), False
                kept = None
        extras = _extras_summary(plan, kept, model)
        stored = {
            **details,
            "user_id": request.user_id,
            "session_hash": request.session_hash,
            "copy_id": copy["id"] if copy is not None else None,
            "template_id": snapshot.hash,
            "filename": filename,
            "extras_dir": str(built) if built is not None else None,
            "replace_extras": replace_extras,
        }
        try:
            token = self.previews.put(stored)
        except BaseException:
            if built is not None:
                shutil.rmtree(built, ignore_errors=True)
            raise
        return {
            "token": token,
            "target_copy_id": copy["id"] if copy is not None else None,
            "template_id": snapshot.hash,
            "filename": filename,
            "identification": identification,
            "can_update": role in EDIT_ROLES,
            "changes": [_public_change(item) for item in details["changes"]],
            "conflicts": details["conflicts"],
            "ignored": details["ignored"],
            "structure": structure,
            "extras": extras,
        }

    def close(self):
        self.previews.close()

    def confirm(self, request):
        if (not isinstance(request.payload, dict)
                or request.payload.get("mode") not in {"update", "new_copy"}):
            raise AppError(422, "invalid", "Choose update or new copy")
        preview = self.previews.take(request.token, request.user_id, request.session_hash)
        try:
            if request.payload["mode"] == "update":
                return self._update(preview, request)
            return self._new_copy(preview, request)
        finally:
            if preview.get("extras_dir"):  # still here when the confirm failed before attaching it
                shutil.rmtree(preview["extras_dir"], ignore_errors=True)

    def _target_copy(self, connection, provenance, request):
        provenance_id = provenance["copy_id"] if provenance else None
        copy_id = request.target_copy_id or provenance_id
        if copy_id is None:
            return None, None
        copy = connection.execute("SELECT * FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
        role = effective_role(connection, copy_id, request.user_id) if copy is not None else None
        if copy is None or role is None:
            if request.target_copy_id is None:
                return None, None
            raise AppError(404, "not_found", "Choose a copy you can access")
        return copy, role

    def _new_copy_snapshot(self, connection, provenance, request):
        member = connection.execute(
            "SELECT 1 FROM pp_category_members WHERE category_id=? AND user_id=?",
            (request.category_id, request.user_id),
        ).fetchone()
        if member is None:
            raise AppError(404, "not_found", "Choose a copy you can access")
        if provenance is not None:
            template_hash = provenance["template_hash"]
            try:
                return self.templates.get(connection, template_hash)
            except AppError as exc:
                if exc.code != "template_missing":
                    raise
                raise AppError(
                    422, "different_template", "This workbook uses a different template",
                ) from exc
        return self.templates.current(connection)

    def _identify(self, model, snapshot, copy):
        if model.provenance:
            if model.provenance["template_hash"] != snapshot.hash:
                raise AppError(422, "different_template", "This workbook uses a different template")
            revision = (model.provenance["input_revision"]
                        if copy is not None and model.provenance["copy_id"] == copy["id"] else None)
            return "provenance", revision
        expected = self._template_model(snapshot)
        if (model.sheets, model.fingerprint) == (expected.sheets, expected.fingerprint):
            return "structure", None
        names = {name.casefold() for name in model.sheets}
        missing = [name for name in expected.sheets if name.casefold() not in names]
        if missing:
            shown = ", ".join(missing[:8]) + (f" and {len(missing) - 8} more" if len(missing) > 8 else "")
            raise AppError(422, "different_template",
                           f"This file does not match the template: it has no sheet named {shown}")
        return "structure", None  # same sheets, some formulas differ: _formula_note says where

    def _template_model(self, snapshot):
        model = self._template_models.get(snapshot.hash)
        if model is None:
            model = _load_template_identity(snapshot)
            if model is None:
                model = _read_workbook(snapshot.source_path, provenance=False)
                _save_template_identity(snapshot, model)
            self._template_models[snapshot.hash] = model
        return model

    def _input_limits(self, snapshot):
        limits = self._input_cols.get(snapshot.hash)
        if limits is None:
            limits = self._input_cols[snapshot.hash] = _input_limits(snapshot)
        return limits

    def warm(self, snapshot):
        """Build the template's upload identity off the request path; uploads then skip a full template read."""
        try:
            self._template_model(snapshot)
            self._input_limits(snapshot)
        except Exception:
            LOGGER.exception("Upload template identity warm-up failed")

    def _update(self, preview, request):
        if preview["copy_id"] is None:
            raise AppError(403, "forbidden", "Choose a copy you can edit")
        excluded = _excluded_conflicts(request.payload, preview["conflicts"])
        changes = [item for item in preview["changes"] if item["key"] not in excluded]
        if not changes and not preview.get("replace_extras"):
            copy = self.copies.get(request.user_id, preview["copy_id"])
            return {"copy_id": preview["copy_id"], "copy": copy,
                    "accepted": 0, "operation": None}
        ack, operation, moved = {"accepted": 0}, None, []
        try:
            with self.database.transaction() as connection:
                if effective_role(connection, preview["copy_id"], request.user_id) not in EDIT_ROLES:
                    raise AppError(403, "forbidden", "Editor access is required")
                replaced = self._attach_extras(connection, preview["copy_id"], preview, moved)
                if changes:
                    ack, _, prepared = self.operations.apply_upload_in_transaction(
                        UploadOperationRequest(
                            connection, preview["copy_id"], request.user_id,
                            _operation_payload(preview["filename"], changes), request.session_check,
                        )
                    )
                    operation = _operation_receipt(ack, prepared)
        except BaseException:
            for folder in moved:  # the copy never pointed at it: the transaction rolled back
                _drop(folder)
            raise
        _drop(replaced)
        copy = self.copies.get(request.user_id, preview["copy_id"])
        return {**ack, "copy_id": preview["copy_id"], "copy": copy, "operation": operation}

    def _attach_extras(self, connection, copy_id, preview, moved):
        """Store the file's own tabs and columns with the copy; returns the folder they replace. The folder
        moved into the copy's place is added to ``moved``, for the caller to delete if the commit fails."""
        if not preview.get("replace_extras"):
            return None
        built = Path(preview["extras_dir"]) if preview.get("extras_dir") else None
        replaced, final = attach(connection, copy_id, built, self.data_dir)
        preview["extras_dir"] = None  # moved into the copy's folder: no longer the preview's to delete
        if final is not None:
            moved.append(final)
        return replaced

    def _new_copy(self, preview, request):
        body = {
            "category_id": request.payload.get("category_id"),
            "name": request.payload.get("name"),
            "visibility": "private",
            "default_role": "viewer",
        }
        inputs = [{**item, "expected_version": 0} for item in preview["all_inputs"]]
        moved = []
        try:
            with self.database.transaction() as connection:
                if self.copies.template_id != preview["template_id"]:
                    raise AppError(409, "refresh_preview", "The active template changed; preview the upload again")
                copy = self.copies.create_in_transaction(
                    connection, request.user_id, body, request.session_check,
                )
                self._attach_extras(connection, copy["id"], preview, moved)
                if inputs:
                    ack, _, prepared = self.operations.apply_upload_in_transaction(
                        UploadOperationRequest(
                            connection, copy["id"], request.user_id,
                            _operation_payload(preview["filename"], inputs), request.session_check,
                        )
                    )
                    operation = _operation_receipt(ack, prepared)
                else:
                    ack, operation = {"accepted": 0, "input_revision": 0}, None
        except BaseException:
            for folder in moved:
                _drop(folder)
            raise
        copy = self.copies.get(request.user_id, copy["id"])
        return {**ack, "copy": copy, "operation": operation}


def _input_limits(snapshot):
    """Per sheet, one past the rightmost template input cell (stored, masked, or unlocked by a column or
    row style), so a user's own columns never start on an input column the template leaves empty."""
    limits = {}
    for key in snapshot.input_mask:
        sheet, _, col = map(int, key.split(":"))
        limits[sheet] = max(limits.get(sheet, 0), col + 1)
    for info in snapshot.sheets:
        sheet = info["index"]
        for path in (snapshot.path / "sheets" / str(sheet)).glob("*.json"):
            for cells in json.loads(path.read_text(encoding="utf-8")).values():
                for col, record in cells.items():
                    if record["k"] == "i":
                        limits[sheet] = max(limits.get(sheet, 0), int(col) + 1)
        if info.get("protected", False):
            width = min(info["cols"], snapshot._used_cols(sheet))
            unlocked = [min(last, width - 1) + 1 for _, last, style in snapshot._column_spans(sheet)
                        if not snapshot.styles[style].get("locked", True)]
            if any(not snapshot.styles[style].get("locked", True) for style in info.get("row_styles", {}).values()):
                unlocked.append(width)
            limits[sheet] = max([limits.get(sheet, 0), *unlocked])
    return limits


def _drop(folder):
    if folder is not None:
        shutil.rmtree(folder, ignore_errors=True)


def _extras_summary(plan, kept, model):
    """What the copy keeps from the file, and its earlier own tabs and columns the file no longer has."""
    names = {name.casefold(): name for name in model.sheets}
    columns = []
    for sheet, first in plan.columns.items():
        upload = names[sheet.casefold()]
        found = sorted(col for col in model.columns.get(upload, {}) if col >= first)
        columns.append({"sheet": upload, "count": len(found),
                        "columns": [_named_column(col, model.headers.get(upload, {}))
                                    for col in found[:MAX_LISTED_COLUMNS]]})
    removed_sheets, removed_columns = [], []
    if kept:
        now = {name.casefold() for name in plan.sheets}
        removed_sheets = [name for name in kept["sheets"] if name.casefold() not in now]
        now = {name.casefold() for name in plan.columns}
        removed_columns = [name for name in kept["columns"] if name.casefold() not in now]
    return {"sheets": list(plan.sheets), "columns": columns,
            "removed_sheets": removed_sheets, "removed_columns": removed_columns}


def _filename(value):
    if not isinstance(value, str):
        raise AppError(422, "invalid_filename", "Upload filename is required")
    name = Path(value).name.strip()
    if not name or len(name) > 120 or any(ord(char) < 32 for char in name):
        raise AppError(422, "invalid_filename", "Upload filename is invalid")
    return name


def _stored_sizes(infos, total_bytes):
    """Measure each member's real on-disk footprint from its header offsets.

    The compress_size in the zip directory is attacker controlled, so the
    expansion ratio is only trustworthy against measured bytes. A forged
    offset still has to point at a readable header, so this cannot be spoofed
    without breaking the archive read itself.
    """
    ordered = sorted(infos, key=lambda item: item.header_offset)
    stored = {}
    for index, info in enumerate(ordered):
        end = ordered[index + 1].header_offset if index + 1 < len(ordered) else total_bytes
        overhead = 30 + len(info.filename.encode("utf-8")) + len(info.extra)
        stored[info.filename] = max(end - info.header_offset - overhead, 1)
    return stored


def _safe_member(info, stored=1):
    name = info.filename
    path = PurePosixPath(name)
    if (not name or "\\" in name or name.startswith("/") or ".." in path.parts
            or info.flag_bits & 1):
        raise AppError(422, "invalid_workbook", "Workbook archive contains an unsafe member")
    if info.file_size > MAX_UNCOMPRESSED:
        raise AppError(413, "archive_bomb", "Workbook expands beyond the safe limit")
    if info.file_size > RATIO_MIN_BYTES and info.file_size / stored > MAX_RATIO:
        raise AppError(413, "archive_bomb", "Workbook compression ratio is unsafe")


OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
UNSUPPORTED_MESSAGES = {
    "xlsb": "This is an .xlsb (binary) workbook. In Excel choose File > Save As > Excel Macro-Enabled "
            "Workbook (.xlsm), then upload that file.",
    "ole": "This is an old .xls workbook, or it is protected with a password. Open it in Excel, remove any "
           "password, save it as Excel Macro-Enabled Workbook (.xlsm), then upload that file.",
    "unknown": "This file is not an Excel workbook. Upload the .xlsm or .xlsx file downloaded from the app.",
}


def _workbook_format(path):
    """ooxml (.xlsx/.xlsm), xlsb, ole (.xls, or any password-protected workbook) or unknown."""
    with open(path, "rb") as handle:
        head = handle.read(len(OLE_MAGIC))
    if head == OLE_MAGIC:
        return "ole"
    if not zipfile.is_zipfile(path):
        return "unknown"
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
    except (OSError, zipfile.BadZipFile):
        return "unknown"
    return "xlsb" if "xl/workbook.bin" in names else "ooxml"


def _unreadable(where, exc):
    """The rejection names the part and the kind of fault, so a file Excel saved oddly can be diagnosed."""
    LOGGER.warning("Upload rejected while reading %s: %s: %s", where, type(exc).__name__, str(exc)[:200])
    return AppError(422, "invalid_workbook",
                    f"File is not a safe Excel workbook (could not read {where}: {type(exc).__name__})")


def _validate_archive(path, deferred=frozenset()):
    """deferred: worksheet parts whose XML screening and link check run on the worker that reads them."""
    try:
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            names = [item.filename for item in infos]
            if len(infos) > MAX_MEMBERS or len(names) != len(set(names)):
                raise AppError(422, "invalid_workbook", "Workbook archive structure is invalid")
            stored = _stored_sizes(infos, path.stat().st_size)
            for info in infos:
                _safe_member(info, stored[info.filename])
            if sum(item.file_size for item in infos) > MAX_UNCOMPRESSED:
                raise AppError(413, "archive_bomb", "Workbook expands beyond the safe limit")
            graph, linkless = _validate_xml_members(archive, stored, deferred)
            _validate_workbook_content_type(archive)
            for _ in _iter_shared_strings(
                    archive, _workbook_role_part(graph, "sharedStrings", "xl/sharedStrings.xml")):
                pass
        validate_package(path, deferred=deferred, no_links=linkless)
    except AppError:
        raise
    except (KeyError, OSError, zipfile.BadZipFile, ET.ParseError, expat.ExpatError, ValueError) as exc:
        raise _unreadable("the workbook package", exc) from exc


def _validate_xml_members(archive, stored, deferred=frozenset()):
    """Returns the relationship graph and the screened parts with no relationship attribute."""
    xml_members = _xml_members(archive)
    total, relationship_total, linkless = 0, 0, set()
    for info in archive.infolist():
        # A deferred part is still unpacked here for the size and ratio limits; only its parse moves
        parser = (_xml_parser() if info.filename in xml_members and info.filename not in deferred
                  else None)
        tree_limit = _tree_xml_limit(info.filename)
        expanded, links = 0, LinkScan()
        with archive.open(info) as stream:
            while chunk := stream.read(1 << 20):
                expanded += len(chunk)
                total += len(chunk)
                if info.filename.endswith(".rels"):
                    relationship_total += len(chunk)
                    if relationship_total > MAX_RELATIONSHIP_XML_TOTAL:
                        raise AppError(
                            413, "archive_bomb", "Workbook relationships exceed the safe limit",
                        )
                if expanded > MAX_UNCOMPRESSED or total > MAX_UNCOMPRESSED:
                    raise AppError(413, "archive_bomb", "Workbook expands beyond the safe limit")
                if tree_limit is not None and expanded > tree_limit:
                    raise AppError(413, "archive_bomb", "Workbook XML metadata is too large")
                if parser is not None:
                    parser.Parse(chunk, False)
                    _check_xml_token(parser, expanded)
                    links.feed(chunk)
        if expanded > RATIO_MIN_BYTES and expanded / stored[info.filename] > MAX_RATIO:
            raise AppError(413, "archive_bomb", "Workbook compression ratio is unsafe")
        if parser is not None:
            parser.Parse(b"", True)
            if not links.possible:
                linkless.add(info.filename)
    return _validate_xml_roles(archive, xml_members), frozenset(linkless)


def _xml_members(archive):
    content_types = archive.getinfo("[Content_Types].xml")
    _screen_xml(archive, content_types, MAX_TREE_XML)
    with archive.open(content_types) as stream:
        root = SafeET.parse(stream).getroot()
    defaults = {
        node.get("Extension") for node in root if _local(node.tag) == "Default"
        and _is_xml_content_type(node.get("ContentType"))
    }
    members = {
        node.get("PartName", "").lstrip("/") for node in root
        if _local(node.tag) == "Override" and _is_xml_content_type(node.get("ContentType"))
    }
    members.update(
        info.filename for info in archive.infolist()
        if (info.filename.endswith((".xml", ".rels"))
            or info.filename.rsplit(".", 1)[-1] in defaults)
    )
    return members | {"[Content_Types].xml"}


def _is_xml_content_type(value):
    value = (value or "").split(";", 1)[0].strip().casefold()
    return value in {"application/xml", "text/xml"} or value.endswith("+xml")


def _screen_xml(archive, info, limit):
    parser = _xml_parser()
    expanded = 0
    with archive.open(info) as stream:
        while chunk := stream.read(1 << 20):
            expanded += len(chunk)
            if expanded > limit:
                raise AppError(413, "archive_bomb", "Workbook XML metadata is too large")
            parser.Parse(chunk, False)
            _check_xml_token(parser, expanded)
    parser.Parse(b"", True)


def _tree_xml_limit(name):
    if name.endswith(".rels"):
        return MAX_RELATIONSHIP_XML
    if name in {"[Content_Types].xml", "xl/workbook.xml"}:
        return MAX_TREE_XML
    if name == "docProps/custom.xml":
        return MAX_CUSTOM_PROPERTIES_XML
    if name == "xl/styles.xml":
        return MAX_STYLES_XML
    if name == "xl/sharedStrings.xml":
        return MAX_SHARED_STRINGS_XML
    if name.startswith("xl/theme/"):
        return MAX_THEME_XML
    return None


class _XMLGuard:
    def __init__(self):
        self.depth = 0
        self.text_bytes = 0

    def start(self, name, attributes):
        self.depth += 1
        self.text_bytes = 0
        if self.depth > MAX_XML_DEPTH:
            raise AppError(413, "archive_bomb", "Workbook XML nesting exceeds the safe limit")
        size = len(name.encode("utf-8")) + sum(
            len(item.encode("utf-8")) for pair in attributes.items() for item in pair
        )
        if size > MAX_XML_TOKEN_BYTES:
            raise AppError(413, "archive_bomb", "Workbook XML token exceeds the safe limit")

    def end(self, *_args):
        self.depth -= 1
        self.text_bytes = 0

    def characters(self, value):
        self.text_bytes += len(value.encode("utf-8"))
        if self.text_bytes > MAX_XML_TOKEN_BYTES:
            raise AppError(413, "archive_bomb", "Workbook XML token exceeds the safe limit")


def _check_xml_token(parser, consumed):
    if consumed - parser.CurrentByteIndex > MAX_XML_TOKEN_BYTES:
        raise AppError(413, "archive_bomb", "Workbook XML token exceeds the safe limit")


def _xml_parser():
    parser, guard = expat.ParserCreate(), _XMLGuard()

    parser.StartElementHandler = guard.start
    parser.EndElementHandler = guard.end
    parser.CharacterDataHandler = guard.characters
    parser.StartDoctypeDeclHandler = _unsafe_xml
    parser.EntityDeclHandler = _unsafe_xml
    parser.UnparsedEntityDeclHandler = _unsafe_xml
    parser.ExternalEntityRefHandler = _unsafe_xml
    return parser


def _validate_xml_roles(archive, xml_members):
    graph = _relationship_graph(archive)
    for relationships in graph.values():
        for relationship in relationships:
            role = relationship["type"].rsplit("/", 1)[-1]
            target = relationship["target"]
            limit = _role_xml_limit(role)
            if target is not None and limit is not None:
                try:
                    _screen_xml(archive, archive.getinfo(target), limit)
                except KeyError as exc:
                    raise AppError(
                        422, "invalid_workbook",
                        "Workbook XML relationship has a non-XML content type",
                    ) from exc
            if target is not None and role in XML_RELATIONSHIP_ROLES and target not in xml_members:
                raise AppError(
                    422, "invalid_workbook",
                    "Workbook XML relationship has a non-XML content type",
                )
    return graph


def _role_xml_limit(role):
    return {
        "styles": MAX_STYLES_XML,
        "sharedStrings": MAX_SHARED_STRINGS_XML,
        "theme": MAX_THEME_XML,
        "chart": MAX_TREE_XML,
        "chartsheet": MAX_TREE_XML,
        "chartSheet": MAX_TREE_XML,
        "drawing": MAX_TREE_XML,
    }.get(role)


def _workbook_role_part(graph, role, fallback):
    return next((item["target"] for item in graph["xl/workbook.xml"]
                 if item["type"].rsplit("/", 1)[-1] == role), fallback)


def _unsafe_xml(*_args):
    raise AppError(422, "unsafe_xml", "Workbook XML declarations are unsafe")


def _validate_workbook_content_type(archive):
    root = SafeET.fromstring(archive.read("[Content_Types].xml"))
    content = next(
        (node.get("ContentType") for node in root if _local(node.tag) == "Override"
         and node.get("PartName", "").lstrip("/") == "xl/workbook.xml"),
        None,
    )
    if content not in {XLSX_MAIN, XLSM_MAIN}:
        raise AppError(422, "invalid_workbook", "Only .xlsx and .xlsm workbook content is supported")


def _no_progress(*_args):
    pass


def _read_workbook(path, provenance=True, snapshot=None, fingerprint=True, progress=_no_progress,
                   overrides=frozenset(), deferred=frozenset(), column_limits=None):
    """overrides: (sheet, row, col) of the copy's saved edits; deferred: parts _validate_archive left to us;
    column_limits: template sheet name (casefolded) -> first column past the template's own."""
    column_limits = column_limits or {}
    with zipfile.ZipFile(path) as archive:
        graph = _relationship_graph(archive)
        sheets = _sheet_records(archive.read("xl/workbook.xml"), graph)
        names = tuple(item["name"] for item in sheets)
        if len({name.casefold() for name in names}) != len(names):
            raise AppError(422, "invalid_workbook", "Workbook sheet names must be unique")
        strings_part = _workbook_role_part(graph, "sharedStrings", "xl/sharedStrings.xml")
        sizes = {info.filename: info.file_size for info in archive.infolist()}
        props = _provenance(archive) if provenance else None
        indexes = ({item["name"].casefold(): item["index"] for item in snapshot.sheets}
                   if snapshot is not None else {})
        jobs = []
        for item in sheets:
            index = indexes.get(item["name"].casefold())
            jobs.append(PartJob(
                str(path), item["part"], strings_part,
                str(snapshot.path) if snapshot is not None else None,
                index, item["name"], item["hidden"], snapshot is None and not provenance, fingerprint,
                frozenset((row, col) for sheet, row, col in overrides if sheet == index),
                item["part"] in deferred,
                column_limits.get(item["name"].casefold()) if index is not None else None,
            ))
        parallel = (_scan_workers(len(jobs)) > 1
                    and sum(sizes.get(job.part, 0) for job in jobs) >= PARALLEL_MIN_BYTES)
        strings = None if parallel else _shared_strings(archive, strings_part)
    results = _scan_parts(jobs, progress, parallel, snapshot, strings, sizes)
    failed = next((result for result in results.values() if result.error is not None), None)
    if failed is not None:
        raise AppError(*failed.error)
    for job in jobs:
        result = results[job.name]
        if job.screen and result.references - {item["id"] for item in graph.get(job.part, [])}:
            raise AppError(422, "invalid_workbook",
                           f"File is not a safe Excel workbook (sheet {job.name} links to a missing part)")
    _screen_unread(path, deferred - {job.part for job in jobs})
    if deferred:
        _deferred_open_check(path)
    progress("Checking rows and columns")
    structures = _bounded_structures({job.name: results[job.name].structure for job in jobs})
    # A file the app wrote carries provenance, which identifies its template without the formula hash
    fingerprint = (_fingerprint(names, {name: result.digest for name, result in results.items()})
                   if fingerprint else None)
    cells = {name: results[name].data.cells for name in names}
    ignored = tuple(
        {"sheet": name, "cell": None, "count": results[name].data.ignored_count,
         "reason": "formula or locked cells are ignored"}
        for name in names if results[name].data.ignored_count
    )
    displaced = frozenset(name for name in names if results[name].data.displaced)
    extents = {name: results[name].data.extent for name in names}
    formula_rows = {name: results[name].data.formula_rows for name in names}
    parts = {job.name: job.part for job in jobs}
    filled_cols = {name: results[name].data.filled_cols for name in names}
    suspect = frozenset(name for name in names if results[name].data.suspect)
    digests = ({name: result.digest.hex() for name, result in results.items() if result.digest is not None}
               if fingerprint else {})
    return WorkbookModel(names, cells, structures, fingerprint, props, ignored, displaced, extents,
                         formula_rows, parts, strings_part, filled_cols, suspect, digests,
                         {name: results[name].data.columns for name in names},
                         {name: results[name].data.headers for name in names},
                         {name: results[name].data.filled_rows for name in names},
                         {job.name: job.limit for job in jobs if job.limit is not None})


@dataclass(frozen=True)
class PartJob:
    """One worksheet part to read, in this process or on a worker; everything in it pickles."""
    path: str
    part: str
    strings_part: str
    snapshot_path: str | None
    index: int | None
    name: str
    hidden: object
    all_formulas: bool
    fingerprint: bool
    keep: frozenset
    screen: bool  # the archive check deferred this part's XML screening and link check to here
    limit: int | None = None


@dataclass(frozen=True)
class PartResult:
    name: str
    data: SheetData | None
    digest: bytes | None
    structure: tuple | None
    references: frozenset
    error: tuple | None = None


def _scan_workers(parts):
    return min(MAX_SCAN_WORKERS, parts, max(1, (os.cpu_count() or 1) - 1))


def _deferred_parts(path):
    """Worksheet parts that move their XML checks onto the workers, when the workbook is big enough."""
    try:
        with zipfile.ZipFile(path) as archive:
            parts = [info for info in archive.infolist()
                     if info.filename.startswith("xl/worksheets/") and info.filename.endswith(".xml")
                     and "/" not in info.filename[len("xl/worksheets/"):]]
    except (OSError, zipfile.BadZipFile):
        return frozenset()  # _validate_archive reports the real problem
    if _scan_workers(len(parts)) < 2 or sum(info.file_size for info in parts) < PARALLEL_MIN_BYTES:
        return frozenset()
    return frozenset(info.filename for info in parts)


def _scan_parts(jobs, progress, parallel, snapshot, strings, sizes):
    results, total = {}, len(jobs)
    progress("Reading sheets", 0, total)
    if not parallel:
        for number, job in enumerate(jobs, 1):
            results[job.name] = _scan_part(job, snapshot, strings)
            progress("Reading sheets", number, total)
        return results
    # Biggest sheets first, so one large sheet does not start last and hold everyone up
    ordered = sorted(jobs, key=lambda job: sizes.get(job.part, 0), reverse=True)
    context = multiprocessing.get_context("spawn")  # never fork a server that runs threads
    try:
        with ProcessPoolExecutor(max_workers=_scan_workers(total), mp_context=context) as pool:
            futures = [pool.submit(_scan_part, job) for job in ordered]
            for future in as_completed(futures):
                result = future.result()
                results[result.name] = result
                progress("Reading sheets", len(results), total)
                if result.error is not None:
                    for pending in futures:
                        pending.cancel()
                    return results
    except BrokenProcessPool:
        # A reader process died (most likely out of memory); finish the upload here, just slower
        LOGGER.warning("A sheet reader process stopped; reading the remaining sheets in the server")
        for job in ordered:
            if job.name not in results:
                results[job.name] = _scan_part(job, snapshot)
                progress("Reading sheets", len(results), total)
    return results


def _scan_part(job, snapshot=None, strings=None):
    try:
        with zipfile.ZipFile(job.path) as archive:
            if job.screen:
                _screen_part(archive, job.part)
            if strings is None:
                strings = _shared_strings(archive, job.strings_part)
            if snapshot is None and job.snapshot_path is not None:
                from .workbook_service import TemplateSnapshot
                snapshot = TemplateSnapshot(job.snapshot_path)
            found, references = _empty_structure(), set() if job.screen else None
            with archive.open(job.part) as stream:
                data = _sheet_cells(SheetScan(archive, job.part, strings, snapshot, job.index, job.name,
                                              job.all_formulas, job.keep, limit=job.limit),
                                    _iter_sheet(stream, found, references))
            structure = (job.hidden, tuple(sorted(found["merge"])), tuple(found["row"]),
                         tuple(found["column"]), found["protection"])
            references = frozenset(references or ())
        digest = _sheet_digest(data.formulas) if job.fingerprint else None
        # Formulas only feed the digest, and a template read keeps no cells, so neither crosses back
        data = replace(data, formulas={}, cells={} if job.all_formulas else data.cells)
        return PartResult(job.name, data, digest, structure, references)
    except AppError as exc:
        return PartResult(job.name, None, None, None, frozenset(), (exc.status, exc.code, exc.message))
    except (KeyError, OSError, zipfile.BadZipFile, ET.ParseError, expat.ExpatError, ValueError) as exc:
        error = _unreadable(f"sheet {job.name}", exc)
        return PartResult(job.name, None, None, None, frozenset(), (error.status, error.code, error.message))


def _screen_part(archive, part):
    """The archive check's XML screen for one part; its size and ratio limits were already enforced."""
    parser = _xml_parser()
    expanded = 0
    with archive.open(part) as stream:
        while chunk := stream.read(1 << 20):
            expanded += len(chunk)
            parser.Parse(chunk, False)
            _check_xml_token(parser, expanded)
    parser.Parse(b"", True)


def _screen_unread(path, parts):
    """Deferred parts no sheet points at still get the screen and link check before anything trusts the file."""
    if not parts:
        return
    part = "the workbook package"
    try:
        with zipfile.ZipFile(path) as archive:
            graph = _relationship_graph(archive)
            for part in sorted(parts):
                _screen_part(archive, part)
                with archive.open(part) as stream:
                    references = _relationship_references(stream)
                if references - {item["id"] for item in graph.get(part, [])}:
                    raise AppError(422, "invalid_workbook",
                                   f"File is not a safe Excel workbook ({part} links to a missing part)")
    except AppError:
        raise
    except (KeyError, OSError, zipfile.BadZipFile, ET.ParseError, expat.ExpatError, ValueError) as exc:
        raise _unreadable(part, exc) from exc


def _deferred_open_check(path):
    """validate_package's openpyxl open, run only now that every worksheet part has been screened."""
    try:
        open_check(path)
    except (KeyError, OSError, zipfile.BadZipFile, ET.ParseError, expat.ExpatError, ValueError) as exc:
        raise _unreadable("the workbook with openpyxl", exc) from exc


def _bounded_structures(structures):
    usage = [0, 0]
    for structure in structures.values():
        _, merges, rows, columns, protection = structure
        for values in (merges, rows, columns):
            for value in values:
                _bounded_structure(("aggregate", value), usage)
        if protection is not None:
            _bounded_structure(("aggregate", protection), usage)
    return structures


def _structure_record(element):
    local = _local(element.tag)
    if local == "mergeCell":
        return "merge", element.get("ref", "").upper()
    if local == "row":
        values = tuple(sorted(
            (name, value) for name, value in element.attrib.items()
            if name not in {"r", "spans"}
        ))
        return ("row", (element.get("r"), values)) if values else (None, None)
    if local == "col":
        return "column", tuple(sorted(element.attrib.items()))
    if local == "sheetProtection":
        return "protection", tuple(sorted(element.attrib.items()))
    return None, None


def _structure_bytes(value):
    if value is None:
        return 0
    if isinstance(value, str):
        return len(value.encode("utf-8"))
    return sum(_structure_bytes(item) for item in value)


def _bounded_structure(record, usage):
    if record[0] is None:
        return record
    usage[0] += 1
    usage[1] += _structure_bytes(record[1])
    if usage[0] > MAX_STRUCTURE_RECORDS or usage[1] > MAX_STRUCTURE_BYTES:
        raise AppError(413, "archive_bomb", "Workbook structure exceeds the safe limit")
    return record


def _empty_structure():
    return {"merge": [], "row": [], "column": [], "protection": None}


def _iter_sheet(stream, found, references=None):
    """One parse of a worksheet: yields each <c> with its children, fills found with the structure
    records, and, given a set, collects the relationship ids the part points at."""
    stack, usage, cell_depth = [], [0, 0], 0
    for event, element in SafeET.iterparse(stream, events=("start", "end")):
        if event == "start":
            stack.append(element)
            if not cell_depth and _local(element.tag) == "c":
                cell_depth = len(stack)
            continue
        if references is not None:
            for name, value in element.attrib.items():
                if name.startswith("{"):
                    namespace, local = name[1:].split("}", 1)
                    if namespace in REL_ATTRIBUTE_NAMESPACES and local in {"id", "embed", "link"}:
                        references.add(value)
        depth = len(stack)
        stack.pop()
        if depth > cell_depth > 0:
            continue  # a cell's children stay until the cell itself is handed out
        if depth == cell_depth:
            cell_depth = 0
            yield element
        else:
            kind, value = _bounded_structure(_structure_record(element), usage)
            if kind == "protection":
                found[kind] = value
            elif kind is not None:
                found[kind].append(value)
        if stack:
            stack[-1].remove(element)
        element.clear()


def _iter_shared_strings(archive, part):
    if part not in archive.namelist():
        return
    count, decoded, phonetic_depth, buffer, stack = 0, 0, 0, None, []
    with archive.open(part) as stream:
        for event, element in SafeET.iterparse(stream, events=("start", "end")):
            local = _local(element.tag)
            if event == "start":
                stack.append(element)
                if local == "si":
                    count += 1
                    if count > MAX_SHARED_STRING_COUNT:
                        raise AppError(
                            413, "archive_bomb", "Workbook shared strings exceed the safe limit"
                        )
                    buffer = StringIO()
                elif local == "rPh" and buffer is not None:
                    phonetic_depth += 1
                continue
            if local == "t" and buffer is not None and not phonetic_depth:
                fragment = element.text or ""
                decoded += len(fragment.encode("utf-8"))
                if decoded > MAX_SHARED_STRING_BYTES:
                    raise AppError(413, "archive_bomb", "Workbook shared strings exceed the safe limit")
                if fragment:
                    buffer.write(fragment)
            elif local == "si" and buffer is not None:
                yield buffer.getvalue()
                buffer.close()
                buffer = None
            elif local == "rPh" and buffer is not None:
                phonetic_depth -= 1
            if len(stack) > 1:
                stack[-2].remove(element)
            element.clear()
            stack.pop()


def _shared_strings(archive, part):
    return list(_iter_shared_strings(archive, part))


def _sheet_cells(request, stream_cells):
    cells, formulas, shared, ignored = {}, {}, {}, 0
    detector = DisplacementDetector(request.snapshot, request.index)
    for cell in stream_cells:
        reference = (cell.get("r") or "").replace("$", "").upper()
        row, col = _coordinate(reference)
        if request.snapshot is None and not request.all_formulas:
            continue
        if request.snapshot is not None and request.index is None:
            continue
        formula = next((child for child in cell if _local(child.tag) == "f"), None)
        if request.limit is not None and col >= request.limit:
            # The user's own columns: counted for the comparison, kept as extras, never read as inputs
            detector.observe({"row": row, "col": col, "value": _cell_value(cell, request.strings),
                              "formula": "=" if formula is not None else None}, extra=True)
            continue
        coordinate = (None if request.snapshot is None
                      else (request.index, row, col))
        expected = None if coordinate is None else request.snapshot.cell(*coordinate)
        expected_formula = expected is None or expected.get("formula") is not None
        if (expected_formula and formula is not None
                and formula.get("t") == "shared" and formula.text):
            # One translator per shared formula: building it tokenizes the master formula
            shared[formula.get("si")] = Translator("=" + formula.text.lstrip("="), origin=reference)
        formula_value = (_formula_value(formula, reference, shared)
                         if expected_formula else "=" if formula is not None else None)
        record = {
            "row": row, "col": col, "value": _cell_value(cell, request.strings),
            "formula": formula_value,
        }
        detector.observe(record)
        if expected_formula and formula_value is not None:
            formulas[reference] = record
        if request.snapshot is None:
            if formula_value is not None:
                cells[reference] = record
            continue
        # A formula typed into an open (unprotected, non-template-input) cell is dropped like a locked one
        template_input, editable = request.snapshot.input_flags(*coordinate, expected)
        if editable and (formula is None or template_input):
            # An open cell that still holds the template value and was never edited cannot change
            # anything, and unprotected sheets hold millions of them
            if (template_input or formula is not None or (row, col) in request.keep
                    or not exactly_equal(record["value"], expected["v"])):
                cells[reference] = record
        elif _locked_cell_changed(record, expected) and not (
                request.moved and record["formula"] is not None and expected.get("formula") is not None):
            ignored += 1
            if not editable:
                detector.locked_value_changed(record, expected)
    return SheetData(cells, formulas, ignored, detector.detected(), (detector.rows, detector.cols),
                     detector.formula_rows, detector.filled_cols,
                     detector.snapshot is not None and detector.sheet is not None and detector.locked_filled,
                     detector.columns, detector.headers, detector.filled_rows)


def _locked_cell_changed(uploaded, expected):
    uploaded_formula = uploaded["formula"]
    expected_formula = expected.get("formula")
    if uploaded_formula == expected_formula is not None:
        return False  # the common case: the formula came back as the template wrote it
    if uploaded_formula is not None or expected_formula is not None:
        return _normalize_formula(uploaded_formula or "") != _normalize_formula(expected_formula or "")
    return not exactly_equal(uploaded["value"], expected["v"])


def _coordinate(reference):
    match = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)", reference)
    if match is None:
        raise AppError(422, "invalid_workbook", "Workbook contains an invalid cell coordinate")
    col = 0
    for letter in match.group(1):
        col = col * 26 + ord(letter) - 64
    return int(match.group(2)) - 1, col - 1


def _cell_value(cell, strings):
    kind = cell.get("t")
    value = next((child.text for child in cell if _local(child.tag) == "v"), None)
    if kind == "s" and value is not None:
        index = int(value)
        return unescape_ooxml(strings[index]) if 0 <= index < len(strings) else ""
    if kind == "b":
        return value == "1"
    if kind == "inlineStr":
        return "".join(unescape_ooxml(node.text or "") for node in cell.iter() if _local(node.tag) == "t")
    if kind == "e":
        return {"error": value or "#VALUE!"}
    if kind in {"str", "d"}:
        return value
    if value is None:
        return None
    try:
        return float(value) if any(marker in value for marker in ".eE") else int(value)
    except ValueError:
        return value


def _formula_value(node, reference, shared):
    if node is None:
        return None
    text = node.text or ""
    if node.get("t") != "shared" or node.get("si") is None:
        return "=" + text.lstrip("=")
    if text:
        return "=" + text.lstrip("=")
    master = shared.get(node.get("si"))
    if master is None:
        raise AppError(422, "invalid_workbook", "Shared formula master is missing")
    return master.translate_formula(reference)


def _sheet_digest(formulas):
    digest = hashlib.sha256()
    for reference, record in sorted(formulas.items()):
        if record["formula"] is not None:
            digest.update(b"F\0" + reference.encode() + b"\0")
            digest.update(_normalize_formula(record["formula"]).encode("utf-8") + b"\0")
    return digest.digest()


def _formula_note(model, template):
    """A file without download information whose formulas differ: say on which sheets, since the
    template's formulas are what the copy keeps."""
    by_name = {name.casefold(): digest for name, digest in template.digests.items()}
    changed = [name for name, digest in model.digests.items()
               if name.casefold() in by_name and by_name[name.casefold()] != digest]
    if not changed:
        return []
    shown = ", ".join(changed[:8]) + (f" and {len(changed) - 8} more" if len(changed) > 8 else "")
    return [{"sheet": None, "cell": None, "reason":
             f"This file has no download information and its formulas differ from the template on: {shown}. "
             "Values were matched by sheet and row; the template's formulas are kept"}]


def _fingerprint(sheet_names, digests):
    """Each sheet is hashed where it is read, so only 32 bytes per sheet leave a worker."""
    digest = hashlib.sha256()
    for name in sheet_names:
        digest.update(b"S\0" + name.casefold().encode("utf-8") + b"\0" + digests[name])
    return digest.hexdigest()


_FORMULA_SPACE = re.compile(r"\s+")


def _normalize_formula(formula):
    """Drop whitespace and case outside string literals. Even split parts sit outside quotes."""
    parts = formula.split('"')
    parts[::2] = [_FORMULA_SPACE.sub("", part).casefold() for part in parts[::2]]
    return '"'.join(parts)


# Bump when the fingerprint or structure records change shape, so a stale saved identity is rebuilt.
TEMPLATE_IDENTITY_VERSION = 6
TEMPLATE_IDENTITY_FILE = "upload_identity.json"


def _load_template_identity(snapshot):
    path = Path(snapshot.path) / TEMPLATE_IDENTITY_FILE
    try:
        saved = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    if not isinstance(saved, dict) or saved.get("version") != TEMPLATE_IDENTITY_VERSION:
        return None
    structures = {name: _tupled(value) for name, value in saved["structures"].items()}
    extents = {name: tuple(value) for name, value in saved["extents"].items()}
    formula_rows = {name: {int(row): value for row, value in rows.items()}
                    for name, rows in saved["formula_rows"].items()}
    return WorkbookModel(tuple(saved["sheets"]), {}, structures, saved["fingerprint"], None, (), frozenset(),
                         extents, formula_rows, filled_cols=saved["filled_cols"], digests=saved["digests"],
                         columns={name: {int(col): count for col, count in value.items()}
                                  for name, value in saved["columns"].items()},
                         headers={name: {int(col): text for col, text in value.items()}
                                  for name, value in saved["headers"].items()},
                         filled_rows=saved["filled_rows"])


def _save_template_identity(snapshot, model):
    path = Path(snapshot.path) / TEMPLATE_IDENTITY_FILE
    body = {"version": TEMPLATE_IDENTITY_VERSION, "sheets": list(model.sheets),
            "fingerprint": model.fingerprint, "structures": model.structures, "extents": model.extents,
            "formula_rows": model.formula_rows, "filled_cols": model.filled_cols,
            "digests": model.digests, "columns": model.columns, "headers": model.headers,
            "filled_rows": model.filled_rows}
    temporary = path.with_suffix(".tmp")
    try:
        temporary.write_text(json.dumps(body), encoding="utf-8")
        temporary.replace(path)
    except OSError:
        LOGGER.warning("Could not save the upload template identity; it is rebuilt on next start")


def _tupled(value):
    """JSON turns the structure tuples into lists; equality with a freshly read upload needs tuples back."""
    return tuple(_tupled(item) for item in value) if isinstance(value, list) else value


def _read_provenance(path):
    with zipfile.ZipFile(path) as archive:
        return _provenance(archive)


def _provenance(archive):
    if "docProps/custom.xml" not in archive.namelist():
        return None
    root = SafeET.fromstring(archive.read("docProps/custom.xml"))
    values = {}
    for node in root:
        if _local(node.tag) != "property":
            continue
        name = node.get("name")
        if name in values:
            raise AppError(422, "invalid_provenance", "Workbook provenance is duplicated")
        values[name] = next((child.text or "" for child in node), "")
    present = {name for name in ("TemplateHash", "CopyId", "InputRevision") if name in values}
    if not present:
        return None
    if len(present) != 3 or not values["TemplateHash"] or not values["CopyId"]:
        raise AppError(422, "invalid_provenance", "Workbook provenance is incomplete")
    try:
        revision = int(values["InputRevision"])
    except ValueError as exc:
        raise AppError(422, "invalid_provenance", "Workbook revision is invalid") from exc
    if revision < 0:
        raise AppError(422, "invalid_provenance", "Workbook revision is invalid")
    return {"template_hash": values["TemplateHash"], "copy_id": values["CopyId"],
            "input_revision": revision}


def _preview_details(context):
    model, snapshot, copy = context.model, context.snapshot, context.copy
    sheet_indexes = {item["name"].casefold(): item["index"] for item in snapshot.sheets}
    overrides = _overrides(context.connection, copy["id"]) if copy is not None else {}
    history = _history_since(context.connection, HistoryRequest(copy["id"], context.revision)) if copy is not None else {}
    changes, conflicts, ignored, all_inputs = [], [], list(model.ignored), []
    extra_sheets, extra_columns = [], {}
    if copy is not None and context.revision is None:
        ignored.append({"sheet": None, "cell": None, "reason":
                        "This file has no download information, so later edits cannot be checked for conflicts"})
    known_order = tuple(name for name in model.sheets if name.casefold() in sheet_indexes)
    expected_order = tuple(
        name for name in context.template_model.sheets
        if name.casefold() in {item.casefold() for item in known_order}
    )
    if tuple(name.casefold() for name in known_order) != tuple(name.casefold() for name in expected_order):
        ignored.append({"sheet": None, "cell": None, "reason": "sheet order changes are ignored"})
    for sheet_name in model.sheets:
        index = sheet_indexes.get(sheet_name.casefold())
        if index is None:
            extra_sheets.append(sheet_name)  # the user's own tab: kept with the copy
            continue
        expected_name = context.template_model.sheets[index]
        if model.structures[sheet_name] != context.template_model.structures[expected_name]:
            ignored.append({
                "sheet": sheet_name, "cell": None,
                "reason": "row heights, column widths, merged cells, protection and hidden rows stay as "
                          "in the template",
            })
        sheet_cells, skipped_rows = model.cells[sheet_name], frozenset()
        filled, expected_filled = (model.filled_cols.get(sheet_name, 0),
                                   context.template_model.filled_cols.get(expected_name, 0))
        if filled < expected_filled:
            # The last filled column moved left: a deleted column would otherwise read as cleared inputs
            ignored.append({"sheet": sheet_name, "cell": None, "reason":
                            "Input changes are ignored because columns may have moved: the last filled "
                            f"column is {_cell_name(0, filled - 1)[:-1] if filled else 'none'}, the "
                            f"template's is {_cell_name(0, expected_filled - 1)[:-1]}"})
            continue
        limit = model.limits.get(sheet_name)
        added = limit is not None and any(col >= limit for col in model.columns.get(sheet_name, {}))
        moved = added and _moved_heading(model, context.template_model, sheet_name, expected_name)
        if moved:
            # A template column's heading is no longer in its place: a column was inserted, not added
            ignored.append({"sheet": sheet_name, "cell": None, "reason":
                            "Input changes and new columns are ignored because columns may have moved: "
                            f"{moved}. Add new columns after column {_cell_name(0, expected_filled - 1)[:-1]}"})
            continue
        if (sheet_name in model.displaced or sheet_name in model.suspect
                or model.formula_rows.get(sheet_name, {})
                != context.template_model.formula_rows.get(expected_name, {})):
            # Formulas or locked labels are not where the template has them: rows were inserted,
            # deleted or the table was extended, so match the rows before reading any input
            matched = _matched_sheet(context, sheet_name, index, overrides, ignored)
            if matched is None:
                continue
            if added and (matched[0] is not sheet_cells or matched[1]):
                # Rows were matched back to the template's, but the new columns sit at the file's rows
                ignored.append({"sheet": sheet_name, "cell": None, "reason":
                                "Your new columns on this sheet are not kept because rows were added or removed "
                                "inside the table. Add rows at the end instead, then upload again"})
                added = False
            sheet_cells, skipped_rows = matched
        if added:
            span = max(model.columns.get(sheet_name, {})) + 1 - limit
            if span > MAX_EXTRA_COLUMNS:
                ignored.append({"sheet": sheet_name, "cell": None, "reason":
                                f"Your new columns on this sheet are not kept: they span {span} columns, more "
                                f"than the {MAX_EXTRA_COLUMNS} a copy can add to one sheet"})
                added = False
        if added:
            extra_columns[expected_name] = limit
        sheet_context = SheetContext(snapshot, index, sheet_name, overrides, ignored, limit)
        found, seen = _inspect_sheet(sheet_cells, sheet_context)
        _add_missing_inputs(found, seen, sheet_context, skipped_rows)
        for item in sorted(found.values(), key=lambda value: (value["r"], value["c"])):
            if not exactly_equal(item["template"], item["new"]):
                all_inputs.append(item)
            change = _uploaded_change(item, history)
            if change is not None:
                changes.append(change)
                if "conflict" in change:
                    conflicts.append(_public_change(change))
    return {"changes": changes, "conflicts": conflicts, "ignored": ignored, "all_inputs": all_inputs,
            "current_revision": copy["input_revision"] if copy is not None else None,
            "extras": {"sheets": extra_sheets, "columns": extra_columns}}


def _moved_heading(model, template, name, expected_name):
    """Why the template's columns look moved in the file (emptied, or a heading elsewhere); None when in place."""
    filled = model.columns.get(name, {})
    emptied = next((col for col in sorted(template.columns.get(expected_name, {})) if not filled.get(col)), None)
    if emptied is not None:
        # A column the template fills is now empty while new columns appeared: they moved right
        return f"column {_cell_name(0, emptied)[:-1]} has data in the template and none in your file"
    found = model.headers.get(name, {})
    places = {}
    for col, heading in found.items():
        places.setdefault(heading, col)
    for col, heading in sorted(template.headers.get(expected_name, {}).items()):
        # Text typed over a heading is an edit; the same heading in another column is a shift
        if found.get(col) != heading and heading in places:
            return (f"the heading \"{heading}\" of column {_cell_name(0, col)[:-1]} is in column "
                    f"{_cell_name(0, places[heading])[:-1]} of your file")
    return None


MAX_ROW_NOTES = 20
MAX_EXTRA_COLUMNS = 200  # a copy's own columns per sheet: each is rendered for every row of the sheet
HEADER_ROWS = 30  # a column's name is the first text in its top rows
MAX_LISTED_COLUMNS = 40


def _structure_report(model, template):
    """Per sheet: what the file has and the template does not, and the reverse. Columns count as present
    when they hold any value or formula; each is named by its letter and the text at its top."""
    found = {name.casefold(): name for name in model.sheets}
    known = {name.casefold() for name in template.sheets}
    report = []
    for name in template.sheets:
        upload = found.get(name.casefold())
        if upload is None:
            report.append({"sheet": name, "status": "missing"})
            continue
        t_cols, u_cols = template.columns.get(name, {}), model.columns.get(upload, {})
        missing = [_named_column(col, template.headers.get(name, {})) for col in sorted(t_cols) if col not in u_cols]
        extra = [_named_column(col, model.headers.get(upload, {})) for col in sorted(u_cols) if col not in t_cols]
        rows = {"template": template.filled_rows.get(name, 0), "file": model.filled_rows.get(upload, 0)}
        same = not missing and not extra and rows["template"] == rows["file"]
        report.append({"sheet": upload, "status": "same" if same else "different", "rows": rows,
                       "columns": {"template": len(t_cols), "file": len(u_cols)},
                       "missing_columns": missing[:MAX_LISTED_COLUMNS], "missing_count": len(missing),
                       "extra_columns": extra[:MAX_LISTED_COLUMNS], "extra_count": len(extra)})
    report += [{"sheet": name, "status": "extra"} for name in model.sheets if name.casefold() not in known]
    return report


def _named_column(col, headers):
    return {"column": _cell_name(0, col)[:-1], "header": headers.get(col)}


def _matched_sheet(context, name, index, overrides, ignored):
    """Rows of this sheet moved: line them up with the template's, then read the inputs at the matched
    rows. Returns (cells, template rows the file no longer has), or None when nothing can be applied."""
    model = context.model
    limit = model.limits.get(name)
    sheet_overrides = {(row, col): item["value"] for (sheet, row, col), item in overrides.items()
                       if sheet == index and (limit is None or col < limit)}
    alignment = align_rows(_template_rows(context.snapshot, index, sheet_overrides),
                           _sheet_rows(context.path, model.parts[name], model.strings_part, limit),
                           model.extents.get(name, (None, None))[0])
    if alignment.failure is not None:
        ignored.append({"sheet": name, "cell": None, "reason":
                        f"Input changes are ignored because rows or columns may have moved: {alignment.failure}"})
        return None
    if not alignment.moved:
        if name in model.displaced:
            ignored.append({"sheet": name, "cell": None, "reason":
                            "Input changes are ignored because rows, columns, or locked cells may have moved"})
            return None
        if alignment.unsure:
            ignored.append({"sheet": name, "cell": None, "reason":
                            "New rows were read as added at the end of the table. No unchanged value (an item "
                            "number or label) shows where they were added; if you inserted rows in the "
                            "middle, cancel and add them at the end instead"})
        return model.cells[name], frozenset()  # the table only grew: formulas copied into new rows
    data = _remapped_sheet(context, name, index, frozenset(sheet_overrides), alignment)
    if data.displaced:
        ignored.append({"sheet": name, "cell": None, "reason":
                        "Input changes are ignored because locked cells changed even after matching rows"})
        return None
    ignored[:] = [item for item in ignored if not (item["sheet"] == name and "count" in item)]
    if data.ignored_count:
        ignored.append({"sheet": name, "cell": None, "count": data.ignored_count,
                        "reason": "formula or locked cells are ignored"})
    _note_rows(ignored, name, alignment)
    return data.cells, frozenset(alignment.deleted)


def _note_rows(ignored, name, alignment):
    shifted = sum(1 for upload_row, row in alignment.mapping.items() if upload_row != row)
    if shifted:
        ignored.append({"sheet": name, "cell": None, "reason":
                        f"Rows were added or removed in your file; {shifted} rows were matched back to "
                        "their template rows by their contents"})
    notes = [f"Row {row + 1} of your file is a new row inside the table; the template's rows are fixed, "
             "so its values were not added" for row in alignment.inserted]
    notes += [f"Row {row + 1} of the template is missing from your file; its values were left unchanged"
              for row in alignment.deleted]
    for note in notes[:MAX_ROW_NOTES]:
        ignored.append({"sheet": name, "cell": None, "reason": note})
    if len(notes) > MAX_ROW_NOTES:
        ignored.append({"sheet": name, "cell": None,
                        "reason": f"{len(notes) - MAX_ROW_NOTES} more added or missing rows"})


def _template_rows(snapshot, index, overrides):
    """The copy as downloaded, by row: template values with the copy's edits, FORMULA for formulas."""
    rows = {}
    for block in range(-(-snapshot.sheets[index].get("rows", 0) // 256)):
        for offset, cells in snapshot.block(index, block).items():
            row = block * 256 + int(offset)
            for col, cell in cells.items():
                value = FORMULA if cell.get("k") == "f" else cell.get("v")
                if value is not None and value != "":
                    rows.setdefault(row, {})[int(col)] = value
    for (row, col), value in overrides.items():
        if value is None or value == "":
            rows.get(row, {}).pop(col, None)
        else:
            rows.setdefault(row, {})[col] = value
    return {row: cells for row, cells in rows.items() if cells}


def _sheet_rows(path, part, strings_part, limit=None):
    """Every filled cell of one uploaded sheet left of ``limit`` (the user's own columns) by row: its value,
    or FORMULA."""
    rows = {}
    with zipfile.ZipFile(path) as archive:
        strings = _shared_strings(archive, strings_part)
        with archive.open(part) as stream:
            for cell in _iter_sheet(stream, _empty_structure()):
                row, col = _coordinate((cell.get("r") or "").replace("$", "").upper())
                if limit is not None and col >= limit:
                    continue
                if any(_local(child.tag) == "f" for child in cell):
                    rows.setdefault(row, {})[col] = FORMULA
                    continue
                value = _cell_value(cell, strings)
                if value is not None and value != "":
                    rows.setdefault(row, {})[col] = value
    return rows


def _remapped_sheet(context, name, index, keep, alignment):
    """The sheet read again with each row placed at its matched template row; unmatched rows dropped."""
    model = context.model
    with zipfile.ZipFile(context.path) as archive:
        strings = _shared_strings(archive, model.strings_part)
        with archive.open(model.parts[name]) as stream:
            scan = SheetScan(archive, model.parts[name], strings, context.snapshot, index, name, False, keep,
                             moved=True, limit=model.limits.get(name))
            return _sheet_cells(scan, _placed(_iter_sheet(stream, _empty_structure()), alignment))


def _placed(cells, alignment):
    for cell in cells:
        row, col = _coordinate((cell.get("r") or "").replace("$", "").upper())
        target = alignment.place(row)
        if target is None:
            continue  # a row the template has no place for
        cell.set("r", _cell_name(target, col))
        yield cell


def _uploaded_change(item, history):
    if exactly_equal(item["old"], item["new"]):
        return None
    record = history.get((item["sheet_index"], item["r"], item["c"]))
    if record is not None and exactly_equal(item["new"], record["base"]):
        return None
    if record is not None:
        item["conflict"] = {"by": record["by"], "at": record["at"]}
    return item


def _inspect_sheet(cells, context):
    inputs, seen = {}, set()
    for reference, uploaded in cells.items():
        coordinate = context.sheet, uploaded["row"], uploaded["col"]
        expected = context.snapshot.cell(*coordinate)
        if context.snapshot.is_input(*coordinate):
            seen.add(coordinate)
        if uploaded["formula"] is not None:
            if _normalize_formula(uploaded["formula"]) != _normalize_formula(expected.get("formula") or ""):
                context.ignored.append({"sheet": context.name, "cell": reference,
                                        "reason": "formula or locked cell is ignored"})
            continue
        if not context.snapshot.is_input(*coordinate):
            if not exactly_equal(uploaded["value"], expected["v"]):
                context.ignored.append({"sheet": context.name, "cell": reference,
                                        "reason": "formula or locked cell is ignored"})
            continue
        try:
            value = validate_value(uploaded["value"])
        except ValueError as exc:
            raise AppError(422, "invalid_value", f"{context.name}!{reference}: {exc}") from exc
        current = context.overrides.get(coordinate)
        old = current["value"] if current else expected["v"]
        version = current["version"] if current else 0
        inputs[coordinate] = _input_item(
            CellIdentity(context.name, reference, coordinate),
            InputValues(old, value, expected["v"], version),
        )
    return inputs, seen


def _add_missing_inputs(inputs, seen, context, skipped_rows=frozenset()):
    """skipped_rows: template rows the upload no longer has (deleted), whose inputs stay as they are."""
    explicit = {
        tuple(map(int, key.split(":"))) for key in context.snapshot.input_mask
        if key.startswith(f"{context.sheet}:")
    }
    implicit_overrides = {
        coordinate for coordinate in context.overrides
        if coordinate[0] == context.sheet and (context.limit is None or coordinate[2] < context.limit)
        and context.snapshot.is_input(*coordinate)
    }
    for coordinate in explicit | implicit_overrides:
        if coordinate in seen or coordinate[1] in skipped_rows:
            continue
        reference = _cell_name(coordinate[1], coordinate[2])
        current = context.overrides.get(coordinate)
        template = context.snapshot.cell(*coordinate)["v"]
        old = current["value"] if current else template
        version = current["version"] if current else 0
        inputs[coordinate] = _input_item(
            CellIdentity(context.name, reference, coordinate),
            InputValues(old, None, template, version),
        )


def _input_item(identity, values):
    sheet, row, col = identity.coordinate
    return {"key": f"{sheet}:{row}:{col}", "sheet": identity.sheet_name, "sheet_index": sheet,
            "cell": identity.reference, "r": row, "c": col, "old": values.old, "new": values.new,
            "template": values.template, "expected_version": values.version}


def _overrides(connection, copy_id):
    rows = connection.execute(
        "SELECT sheet,row,col,value_json,version FROM pp_overrides WHERE copy_id=?", (copy_id,)
    ).fetchall()
    import json
    return {(row["sheet"], row["row"], row["col"]): {
        "value": json.loads(row["value_json"]), "version": row["version"]
    } for row in rows}


def _history_since(connection, request):
    if request.revision is None:
        return {}
    rows = connection.execute(
        "SELECT h.sheet,h.row,h.col,h.before_json,h.at,u.id,u.name FROM pp_history h "
        "JOIN pp_users u ON u.id=h.actor_id WHERE h.copy_id=? AND h.revision>? "
        "ORDER BY h.id",
        (request.copy_id, request.revision),
    ).fetchall()
    history = {}
    for row in rows:
        key = row["sheet"], row["row"], row["col"]
        record = history.setdefault(key, {"base": json.loads(row["before_json"])})
        record.update({"at": row["at"], "by": {"id": row["id"], "name": row["name"]}})
    return history


def _public_change(item):
    output = {key: item[key] for key in (
        "key", "sheet", "cell", "r", "c", "old", "new", "expected_version"
    )}
    if "conflict" in item:
        output["conflict"] = item["conflict"]
        output.update(item["conflict"])
    return output


def _excluded_conflicts(payload, conflicts):
    value = payload.get("excluded_conflicts", [])
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        raise AppError(422, "invalid", "Conflict selection is invalid")
    known = {item["key"] for item in conflicts}
    if set(value) - known:
        raise AppError(422, "invalid", "Conflict selection is invalid")
    return set(value)


def _operation_payload(filename, items):
    return {"op_id": f"upload-{uuid4().hex}", "action": f"upload: {filename}", "cells": [
        {"sheet": item["sheet_index"], "r": item["r"], "c": item["c"],
         "v": item["new"], "expected_version": item["expected_version"]}
        for item in items
    ]}


def _operation_receipt(ack, prepared):
    return {"op_id": ack["op_id"], "action": "upload", "cells": [
        {"sheet": item["sheet"], "r": item["row"], "c": item["col"],
         "v": item["value"], "before": item["before"],
         "before_version": item["version"] - 1, "after_version": item["version"]}
        for item in prepared
    ]}


def _cell_name(row, col):
    output, value = "", col + 1
    while value:
        value, remainder = divmod(value - 1, 26)
        output = chr(65 + remainder) + output
    return f"{output}{row + 1}"
