"""Bounded workbook upload inspection, preview, and confirmation."""

from __future__ import annotations

from dataclasses import dataclass
import hashlib
import json
from pathlib import Path, PurePosixPath
import re
import secrets
import tempfile
import threading
import time
from uuid import uuid4
import zipfile
import xml.etree.ElementTree as ET
import xml.parsers.expat as expat

from openpyxl.formula.translate import Translator
from openpyxl.utils.escape import unescape as unescape_ooxml

from .authorization import effective_role
from .errors import AppError
from .export.transfer import (
    _iter_cells,
    _local,
    _relationship_graph,
    _sheet_records,
    validate_package,
)
from .ops import EDIT_ROLES, UploadOperationRequest
from .values import exactly_equal, validate_value


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XLSX_MAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
XLSM_MAIN = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
MAX_MEMBERS = 10_000
MAX_UNCOMPRESSED = 500 * 1024 * 1024
MAX_RATIO = 200
RATIO_MIN_BYTES = 1024 * 1024
MAX_TREE_XML = 8 * 1024 * 1024
MAX_RELATIONSHIP_XML = 4 * 1024 * 1024
MAX_CUSTOM_PROPERTIES_XML = 1024 * 1024
PREVIEW_SECONDS = 15 * 60
MAX_ACTIVE_PREVIEWS = 128
MAX_PREVIEW_CACHE_BYTES = 2 * MAX_UNCOMPRESSED
XML_RELATIONSHIP_ROLES = {
    "officeDocument", "worksheet", "sharedStrings", "custom-properties",
}


@dataclass(frozen=True)
class WorkbookModel:
    sheets: tuple[str, ...]
    cells: dict[str, dict[str, dict]]
    structures: dict[str, tuple]
    fingerprint: str
    provenance: dict | None
    ignored: tuple[dict, ...]


@dataclass(frozen=True)
class UploadDependencies:
    database: object
    templates: object
    copies: object
    operations: object
    preview_dir: Path


@dataclass(frozen=True)
class PreviewRequest:
    path: Path
    filename: str
    user_id: str
    session_hash: str
    target_copy_id: str | None = None
    category_id: str | None = None


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


@dataclass(frozen=True)
class SheetContext:
    snapshot: object
    sheet: int
    name: str
    overrides: dict
    ignored: list


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
    user_id: str
    revision: int | None


@dataclass(frozen=True)
class PreviewEntry:
    path: Path
    user_id: str
    session_hash: str
    expires: float
    size: int


@dataclass(frozen=True)
class SheetScan:
    archive: object
    part: str
    strings: list
    snapshot: object
    index: int | None
    name: str
    all_formulas: bool


@dataclass(frozen=True)
class SheetData:
    cells: dict
    formulas: dict
    ignored_count: int


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
                time.monotonic() + PREVIEW_SECONDS, size,
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
                raise AppError(404, "preview_expired", "Upload preview expired; preview the file again")
            return json.loads(value.path.read_text(encoding="utf-8"))
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


class UploadService:
    def __init__(self, dependencies):
        self.database = dependencies.database
        self.templates = dependencies.templates
        self.copies = dependencies.copies
        self.operations = dependencies.operations
        self.previews = PreviewStore(dependencies.preview_dir)
        self._template_models = {}

    def preview(self, request):
        filename = _filename(request.filename)
        _validate_archive(request.path)
        identity = _read_workbook(request.path)
        with self.database.read_transaction() as connection:
            copy, role = self._target_copy(connection, identity, request)
            snapshot = (self.templates.get(connection, copy["template_id"]) if copy is not None
                        else self._new_copy_snapshot(connection, identity, request))
            model = _read_workbook(request.path, snapshot=snapshot)
            identification, revision = self._identify(model, snapshot, copy)
            details = _preview_details(PreviewContext(
                connection, model, self._template_model(snapshot), snapshot,
                copy, request.user_id, revision,
            ))
        stored = {
            **details,
            "user_id": request.user_id,
            "session_hash": request.session_hash,
            "copy_id": copy["id"] if copy is not None else None,
            "template_id": snapshot.hash,
            "filename": filename,
        }
        token = self.previews.put(stored)
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
        }

    def close(self):
        self.previews.close()

    def confirm(self, request):
        preview = self.previews.take(request.token, request.user_id, request.session_hash)
        if (not isinstance(request.payload, dict)
                or request.payload.get("mode") not in {"update", "new_copy"}):
            raise AppError(422, "invalid", "Choose update or new copy")
        if request.payload["mode"] == "update":
            return self._update(preview, request)
        return self._new_copy(preview, request)

    def _target_copy(self, connection, model, request):
        provenance_id = model.provenance["copy_id"] if model.provenance else None
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

    def _new_copy_snapshot(self, connection, model, request):
        member = connection.execute(
            "SELECT 1 FROM pp_category_members WHERE category_id=? AND user_id=?",
            (request.category_id, request.user_id),
        ).fetchone()
        if member is None:
            raise AppError(404, "not_found", "Choose a copy you can access")
        if model.provenance is not None:
            return self.templates.get(connection, model.provenance["template_hash"])
        return self.templates.current(connection)

    def _identify(self, model, snapshot, copy):
        if model.provenance:
            if model.provenance["template_hash"] != snapshot.hash:
                raise AppError(422, "different_template", "This workbook uses a different template")
            revision = (model.provenance["input_revision"]
                        if copy is not None and model.provenance["copy_id"] == copy["id"] else None)
            return "provenance", revision
        expected = self._template_model(snapshot)
        if (model.sheets, model.fingerprint) != (expected.sheets, expected.fingerprint):
            raise AppError(422, "different_template", "Workbook structure does not match this template")
        return "structure", None

    def _template_model(self, snapshot):
        model = self._template_models.get(snapshot.hash)
        if model is None:
            model = _read_workbook(snapshot.source_path, provenance=False)
            self._template_models[snapshot.hash] = model
        return model

    def _update(self, preview, request):
        if preview["copy_id"] is None:
            raise AppError(403, "forbidden", "Choose a copy you can edit")
        excluded = _excluded_conflicts(request.payload, preview["conflicts"])
        changes = [item for item in preview["changes"] if item["key"] not in excluded]
        if not changes:
            copy = self.copies.get(request.user_id, preview["copy_id"])
            return {"copy_id": preview["copy_id"], "copy": copy,
                    "accepted": 0, "operation": None}
        with self.database.transaction() as connection:
            current = connection.execute(
                "SELECT input_revision FROM pp_copies WHERE id=?", (preview["copy_id"],)
            ).fetchone()
            if current is None or current["input_revision"] != preview["current_revision"]:
                raise AppError(409, "refresh_preview", "The copy changed; preview the upload again")
            ack, _, prepared = self.operations.apply_upload_in_transaction(
                UploadOperationRequest(
                    connection, preview["copy_id"], request.user_id,
                    _operation_payload(preview["filename"], changes), request.session_check,
                )
            )
        copy = self.copies.get(request.user_id, preview["copy_id"])
        return {**ack, "copy_id": preview["copy_id"], "copy": copy,
                "operation": _operation_receipt(ack, prepared)}

    def _new_copy(self, preview, request):
        body = {
            "category_id": request.payload.get("category_id"),
            "name": request.payload.get("name"),
            "visibility": "private",
            "default_role": "viewer",
        }
        inputs = [{**item, "expected_version": 0} for item in preview["all_inputs"]]
        with self.database.transaction() as connection:
            if self.copies.template_id != preview["template_id"]:
                raise AppError(409, "refresh_preview", "The active template changed; preview the upload again")
            copy = self.copies.create_in_transaction(
                connection, request.user_id, body, request.session_check,
            )
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
        copy = self.copies.get(request.user_id, copy["id"])
        return {**ack, "copy": copy, "operation": operation}


def _filename(value):
    if not isinstance(value, str):
        raise AppError(422, "invalid_filename", "Upload filename is required")
    name = Path(value).name.strip()
    if not name or len(name) > 120 or any(ord(char) < 32 for char in name):
        raise AppError(422, "invalid_filename", "Upload filename is invalid")
    return name


def _safe_member(info):
    name = info.filename
    path = PurePosixPath(name)
    if (not name or "\\" in name or name.startswith("/") or ".." in path.parts
            or info.flag_bits & 1):
        raise AppError(422, "invalid_workbook", "Workbook archive contains an unsafe member")
    if info.file_size > MAX_UNCOMPRESSED:
        raise AppError(413, "archive_bomb", "Workbook expands beyond the safe limit")
    if info.file_size > RATIO_MIN_BYTES and info.file_size / max(info.compress_size, 1) > MAX_RATIO:
        raise AppError(413, "archive_bomb", "Workbook compression ratio is unsafe")


def _validate_archive(path):
    try:
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            names = [item.filename for item in infos]
            if len(infos) > MAX_MEMBERS or len(names) != len(set(names)):
                raise AppError(422, "invalid_workbook", "Workbook archive structure is invalid")
            for info in infos:
                _safe_member(info)
            if sum(item.file_size for item in infos) > MAX_UNCOMPRESSED:
                raise AppError(413, "archive_bomb", "Workbook expands beyond the safe limit")
            _validate_xml_members(archive)
            _validate_workbook_content_type(archive)
        validate_package(path)
    except AppError:
        raise
    except (KeyError, OSError, zipfile.BadZipFile, ET.ParseError, expat.ExpatError, ValueError) as exc:
        raise AppError(422, "invalid_workbook", "File is not a safe Excel workbook") from exc


def _validate_xml_members(archive):
    xml_members = _xml_members(archive)
    total = 0
    for info in archive.infolist():
        parser = _xml_parser() if info.filename in xml_members else None
        tree_limit = _tree_xml_limit(info.filename)
        expanded = 0
        with archive.open(info) as stream:
            while chunk := stream.read(1 << 20):
                expanded += len(chunk)
                total += len(chunk)
                if expanded > MAX_UNCOMPRESSED or total > MAX_UNCOMPRESSED:
                    raise AppError(413, "archive_bomb", "Workbook expands beyond the safe limit")
                if tree_limit is not None and expanded > tree_limit:
                    raise AppError(413, "archive_bomb", "Workbook XML metadata is too large")
                if parser is not None:
                    parser.Parse(chunk, False)
        if expanded > RATIO_MIN_BYTES and expanded / max(info.compress_size, 1) > MAX_RATIO:
            raise AppError(413, "archive_bomb", "Workbook compression ratio is unsafe")
        if parser is not None:
            parser.Parse(b"", True)
    _validate_xml_roles(archive, xml_members)


def _xml_members(archive):
    content_types = archive.getinfo("[Content_Types].xml")
    _screen_xml(archive, content_types, MAX_TREE_XML)
    with archive.open(content_types) as stream:
        root = ET.parse(stream).getroot()
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
    parser.Parse(b"", True)


def _tree_xml_limit(name):
    if name.endswith(".rels"):
        return MAX_RELATIONSHIP_XML
    if name in {"[Content_Types].xml", "xl/workbook.xml"}:
        return MAX_TREE_XML
    if name == "docProps/custom.xml":
        return MAX_CUSTOM_PROPERTIES_XML
    return None


def _xml_parser():
    parser = expat.ParserCreate()
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
            if target is not None and role in XML_RELATIONSHIP_ROLES and target not in xml_members:
                raise AppError(
                    422, "invalid_workbook",
                    "Workbook XML relationship has a non-XML content type",
                )


def _unsafe_xml(*_args):
    raise AppError(422, "unsafe_xml", "Workbook XML declarations are unsafe")


def _validate_workbook_content_type(archive):
    root = ET.fromstring(archive.read("[Content_Types].xml"))
    content = next(
        (node.get("ContentType") for node in root if _local(node.tag) == "Override"
         and node.get("PartName", "").lstrip("/") == "xl/workbook.xml"),
        None,
    )
    if content not in {XLSX_MAIN, XLSM_MAIN}:
        raise AppError(422, "invalid_workbook", "Only .xlsx and .xlsm workbook content is supported")


def _read_workbook(path, provenance=True, snapshot=None):
    with zipfile.ZipFile(path) as archive:
        graph = _relationship_graph(archive)
        sheets = _sheet_records(archive.read("xl/workbook.xml"), graph)
        names = tuple(item["name"] for item in sheets)
        if len({name.casefold() for name in names}) != len(names):
            raise AppError(422, "invalid_workbook", "Workbook sheet names must be unique")
        strings = _shared_strings(archive)
        indexes = ({item["name"].casefold(): item["index"] for item in snapshot.sheets}
                   if snapshot is not None else {})
        all_formulas = snapshot is None and not provenance
        scanned = {
            item["name"]: _sheet_cells(SheetScan(
                archive, item["part"], strings, snapshot,
                indexes.get(item["name"].casefold()), item["name"], all_formulas,
            )) for item in sheets
        }
        cells = {name: data.cells for name, data in scanned.items()}
        structures = {
            item["name"]: _sheet_structure(archive, item["part"], item["hidden"])
            for item in sheets
        }
        fingerprint = _fingerprint(names, {
            name: data.formulas for name, data in scanned.items()
        })
        props = _provenance(archive) if provenance else None
        ignored = tuple(
            {"sheet": name, "cell": None, "count": data.ignored_count,
             "reason": "formula or locked cells are ignored"}
            for name, data in scanned.items() if data.ignored_count
        )
    return WorkbookModel(names, cells, structures, fingerprint, props, ignored)


def _sheet_structure(archive, part, hidden):
    merges, rows, columns, protection = [], [], [], None
    with archive.open(part) as stream:
        for _, element in ET.iterparse(stream, events=("end",)):
            local = _local(element.tag)
            if local == "mergeCell":
                merges.append(element.get("ref", "").upper())
            elif local == "row":
                values = tuple(sorted(
                    (name, value) for name, value in element.attrib.items() if name != "r"
                ))
                if values:
                    rows.append((element.get("r"), values))
            elif local == "col":
                columns.append(tuple(sorted(element.attrib.items())))
            elif local == "sheetProtection":
                protection = tuple(sorted(element.attrib.items()))
            element.clear()
    return hidden, tuple(sorted(merges)), tuple(rows), tuple(columns), protection


def _shared_strings(archive):
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    output = []
    with archive.open("xl/sharedStrings.xml") as stream:
        for _, element in ET.iterparse(stream, events=("end",)):
            if _local(element.tag) == "si":
                output.append("".join(node.text or "" for node in element.iter() if _local(node.tag) == "t"))
                element.clear()
    return output


def _sheet_cells(request):
    cells, formulas, shared, ignored = {}, {}, {}, 0
    with request.archive.open(request.part) as stream:
        for cell in _iter_cells(stream):
            reference = (cell.get("r") or "").replace("$", "").upper()
            row, col = _coordinate(reference)
            if request.snapshot is None and not request.all_formulas:
                continue
            if request.snapshot is not None and request.index is None:
                continue
            formula = next((child for child in cell if _local(child.tag) == "f"), None)
            coordinate = (None if request.snapshot is None
                          else (request.index, row, col))
            expected = None if coordinate is None else request.snapshot.cell(*coordinate)
            expected_formula = expected is None or expected.get("formula") is not None
            if (expected_formula and formula is not None
                    and formula.get("t") == "shared" and formula.text):
                shared[formula.get("si")] = (reference, formula.text)
            formula_value = (_formula_value(formula, reference, shared)
                             if expected_formula else "=" if formula is not None else None)
            record = {
                "row": row, "col": col, "value": _cell_value(cell, request.strings),
                "formula": formula_value,
            }
            if expected_formula and formula_value is not None:
                formulas[reference] = record
            if request.snapshot is None:
                if formula_value is not None:
                    cells[reference] = record
                continue
            if request.snapshot.is_input(*coordinate):
                cells[reference] = record
            elif _locked_cell_changed(record, expected):
                ignored += 1
    return SheetData(cells, formulas, ignored)


def _locked_cell_changed(uploaded, expected):
    uploaded_formula = uploaded["formula"]
    expected_formula = expected.get("formula")
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
    return Translator("=" + master[1].lstrip("="), origin=master[0]).translate_formula(reference)


def _fingerprint(sheet_names, cells):
    digest = hashlib.sha256()
    for name in sheet_names:
        digest.update(b"S\0" + name.casefold().encode("utf-8") + b"\0")
        for reference, record in sorted(cells[name].items()):
            if record["formula"] is not None:
                digest.update(b"F\0" + reference.encode() + b"\0")
                digest.update(_normalize_formula(record["formula"]).encode("utf-8") + b"\0")
    return digest.hexdigest()


def _normalize_formula(formula):
    output, quoted = [], False
    for char in formula:
        if char == '"':
            quoted = not quoted
            output.append(char)
        elif quoted:
            output.append(char)
        elif not char.isspace():
            output.append(char.casefold())
    return "".join(output)


def _provenance(archive):
    if "docProps/custom.xml" not in archive.namelist():
        return None
    root = ET.fromstring(archive.read("docProps/custom.xml"))
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
    history = (_history_since(context.connection, HistoryRequest(
        copy["id"], context.user_id, context.revision,
    )) if copy is not None else {})
    changes, conflicts, ignored, all_inputs = [], [], list(model.ignored), []
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
            ignored.append({"sheet": sheet_name, "cell": None, "reason": "new sheet is ignored"})
            continue
        expected_name = context.template_model.sheets[index]
        if model.structures[sheet_name] != context.template_model.structures[expected_name]:
            ignored.append({
                "sheet": sheet_name, "cell": None,
                "reason": "row, column, merge, protection, or visibility changes are ignored",
            })
        sheet_context = SheetContext(snapshot, index, sheet_name, overrides, ignored)
        found, seen = _inspect_sheet(model.cells[sheet_name], sheet_context)
        _add_missing_inputs(found, seen, sheet_context)
        for item in sorted(found.values(), key=lambda value: (value["r"], value["c"])):
            if not exactly_equal(item["template"], item["new"]):
                all_inputs.append(item)
            if not exactly_equal(item["old"], item["new"]):
                conflict = history.get((index, item["r"], item["c"]))
                if conflict:
                    item["conflict"] = {"by": conflict["by"], "at": conflict["at"]}
                    conflicts.append(_public_change(item))
                changes.append(item)
    return {"changes": changes, "conflicts": conflicts, "ignored": ignored, "all_inputs": all_inputs,
            "current_revision": copy["input_revision"] if copy is not None else None}


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


def _add_missing_inputs(inputs, seen, context):
    explicit = {
        tuple(map(int, key.split(":"))) for key in context.snapshot.input_mask
        if key.startswith(f"{context.sheet}:")
    }
    implicit_overrides = {
        coordinate for coordinate in context.overrides
        if coordinate[0] == context.sheet and context.snapshot.is_input(*coordinate)
    }
    for coordinate in explicit | implicit_overrides:
        if coordinate in seen:
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
        "SELECT h.sheet,h.row,h.col,h.at,u.id,u.name FROM pp_history h "
        "JOIN pp_users u ON u.id=h.actor_id WHERE h.copy_id=? AND h.revision>? "
        "AND h.actor_id<>? ORDER BY h.id",
        (request.copy_id, request.revision, request.user_id),
    ).fetchall()
    return {(row["sheet"], row["row"], row["col"]): {
        "at": row["at"], "by": {"id": row["id"], "name": row["name"]}
    } for row in rows}


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
