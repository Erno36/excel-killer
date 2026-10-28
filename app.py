"""Prototipo web del Cruzador; loopback por defecto, LAN opt-in con lista permitida."""
from __future__ import annotations

import hashlib
import hmac
import ipaddress
import io
import json
import os
import secrets
import sqlite3
import sys
import tempfile
import threading
import time
import uuid
import zipfile
from dataclasses import replace
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from flask import (
    Flask,
    abort,
    flash,
    g,
    has_request_context,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from werkzeug.exceptions import HTTPException
from werkzeug.utils import secure_filename

ROOT = Path(__file__).resolve().parents[1]
DESKTOP_APP = ROOT / "cruzador_excel"
if str(DESKTOP_APP) not in sys.path:
    sys.path.insert(0, str(DESKTOP_APP))

import pandas as pd

from auth.service import AuthService
from audit.logger import make_timestamp
from data.db import init_db
from data.models import (
    BusinessRule,
    CleansingRule,
    CrossConfig,
    CrossStepSpec,
    FileEntry,
    FileMetadata,
    Recipe,
)
from engine.cleansing import OPERACIONES_VALIDAS, apply_all_rules, apply_rule
from engine.column_matcher import suggest_pairs
from engine.cross import CrossConfigError, execute_cross
from engine.exporter import XlsxExporter
from engine.filter_ops import FilterCondition, OPERADORES_VALIDOS, apply_filters
from engine.reader import EmptyFileError, ExcelFileReader, FileReadError
from engine.rules_engine import (
    apply_rule as apply_business_rule,
    CONDITION_OPERATORS,
)
from engine.recipe import build_recipe_from_state, list_recipes, save_recipe
from engine.stats import compute_stats
from engine._date_utils import parse_date_flexible


APP_VERSION = "0.1.0-local"
RULE_VERSION = "header-detect-v1-cross-v1"
MAX_FILES = 10
MAX_FILE_BYTES = 20 * 1024 * 1024
MAX_UNCOMPRESSED_XLSX_BYTES = 100 * 1024 * 1024
MAX_XLSX_PARTS = 2_000
MAX_ROWS = 250_000
MAX_COLUMNS = 500
MAX_ACTIVE_WORKFLOWS = 3
SESSION_TTL_SECONDS = 15 * 60
AUTH_TIMEOUT_SECONDS = 30 * 60
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}
_workflows: dict[str, dict] = {}
_workflow_lock = threading.RLock()
_audit_lock = threading.Lock()
BIND_HOST = os.getenv("CRUZADOR_WEB_BIND_HOST", "127.0.0.1")
try:
    _bind_address = ipaddress.ip_address(BIND_HOST)
    _allowed_remote_addresses = frozenset(
        ipaddress.ip_address(value.strip())
        for value in os.getenv("CRUZADOR_WEB_ALLOWED_CLIENTS", "").split(",")
        if value.strip()
    )
except ValueError as exc:
    raise RuntimeError("La configuracion de IP del servidor o clientes no es valida.") from exc

if _bind_address.is_unspecified or _bind_address.is_multicast:
    raise RuntimeError("CRUZADOR_WEB_BIND_HOST debe ser una IP especifica de esta PC.")
if not _bind_address.is_loopback and not _allowed_remote_addresses:
    raise RuntimeError(
        "Para escuchar en una IP de red, configura CRUZADOR_WEB_ALLOWED_CLIENTS."
    )


def _default_data_dir() -> Path:
    base = Path(os.getenv("XDG_STATE_HOME", Path.home() / ".local" / "state"))
    return base / "banco-piano" / "cruzador-excel-web"


_auth_db_path = Path(
    os.getenv("CRUZADOR_WEB_DB_PATH", _default_data_dir() / "cruzador.db")
)
_auth_db_lock = threading.Lock()
_auth_db_ready = False


def _auth_connection():
    global _auth_db_ready
    with _auth_db_lock:
        if not _auth_db_ready:
            connection = init_db(
                _auth_db_path,
                is_development=(
                    os.getenv("APP_ENV", "production").strip().lower() == "development"
                ),
            )
            connection.close()
            _auth_db_ready = True
    from data.db import get_connection
    return get_connection(_auth_db_path)


def _recipes_dir() -> Path:
    return Path(os.getenv("CRUZADOR_WEB_RECIPES_DIR", _default_data_dir() / "recipes"))


def _get_recipe(recipe_id: str) -> Recipe | None:
    if not recipe_id:
        return None
    for recipe in list_recipes(_recipes_dir()):
        if recipe.recipe_id == recipe_id:
            return recipe
    return None


class WebAuditLogger:
    """Log diario append-only con encadenamiento de hashes, sin celdas."""

    def __init__(self, directory: Path) -> None:
        self.directory = directory

    def write(
        self,
        session_id: str,
        event: str,
        result: str,
        detail: dict,
        input_hash: str = "",
        user: str | None = None,
    ) -> None:
        with _audit_lock:
            path = self.directory / f"{datetime.now().astimezone():%Y-%m-%d}.jsonl"
            path.parent.mkdir(parents=True, exist_ok=True)
            previous_hash = "GENESIS"
            if path.exists():
                lines = path.read_bytes().splitlines()
                if lines:
                    previous_hash = hashlib.sha256(lines[-1]).hexdigest()
            record = {
                "ts": make_timestamp(),
                "session_id": session_id,
                "user": user or (
                    session.get("username", "local-operator")
                    if has_request_context()
                    else "system"
                ),
                "event": event,
                "component": "cruzador-excel-web",
                "version": APP_VERSION,
                "input_hash": input_hash,
                "rule_version": RULE_VERSION,
                "result": result,
                "detail": detail,
                "prev_hash": previous_hash,
            }
            line = json.dumps(record, ensure_ascii=False, sort_keys=True)
            with path.open("a", encoding="utf-8") as handle:
                handle.write(line + "\n")


app = Flask(__name__)
web_secret = os.getenv("CRUZADOR_WEB_SECRET")
if os.getenv("CRUZADOR_WEB_REQUIRE_SECRET") == "1" and not web_secret:
    raise RuntimeError("CRUZADOR_WEB_SECRET es obligatorio para el servicio.")
app.config.update(
    SECRET_KEY=web_secret or secrets.token_hex(32),
    MAX_CONTENT_LENGTH=MAX_FILES * MAX_FILE_BYTES + 1024 * 1024,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Strict",
    SESSION_COOKIE_SECURE=os.getenv("CRUZADOR_WEB_COOKIE_SECURE", "1") == "1",
)
_audit = WebAuditLogger(Path(os.getenv("CRUZADOR_WEB_LOG_DIR", _default_data_dir() / "logs")))
_reader = ExcelFileReader()


def _cleanup_expired_workflows() -> None:
    cutoff = time.monotonic() - SESSION_TTL_SECONDS
    with _workflow_lock:
        expired = [
            workflow_id
            for workflow_id, workflow in _workflows.items()
            if workflow["last_activity"] < cutoff
        ]
        for workflow_id in expired:
            del _workflows[workflow_id]


@app.before_request
def _protect_local_workflow():
    try:
        remote_address = ipaddress.ip_address(request.remote_addr or "")
    except ValueError:
        abort(403)
    if (
        not remote_address.is_loopback
        and remote_address not in _allowed_remote_addresses
    ):
        abort(403)

    _cleanup_expired_workflows()
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_urlsafe(32)
    if "workflow_id" not in session:
        session["workflow_id"] = secrets.token_urlsafe(32)
    g.workflow_id = session["workflow_id"]
    g.error_id = uuid.uuid4().hex[:12].upper()

    public_endpoints = {"login", "static"}
    if not app.config["TESTING"] and request.endpoint not in public_endpoints:
        username = session.get("username")
        last_activity = session.get("last_activity")
        if not username:
            return redirect(url_for("login"))
        if not isinstance(last_activity, (int, float)) or time.time() - last_activity > AUTH_TIMEOUT_SECONDS:
            workflow_id = session.get("workflow_id")
            if workflow_id:
                with _workflow_lock:
                    _workflows.pop(workflow_id, None)
            session.clear()
            flash("La sesión venció por inactividad. Ingresá nuevamente.")
            return redirect(url_for("login"))
        session["last_activity"] = time.time()

    if request.method == "POST":
        supplied = request.form.get("csrf_token", "")
        if not hmac.compare_digest(supplied, session["csrf_token"]):
            abort(400, description="La solicitud vencio o no pudo validarse. Volve a cargar la pagina.")


@app.after_request
def _secure_response(response):
    response.headers["Cache-Control"] = "no-store"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "no-referrer"
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; style-src 'self'; img-src 'self' data:; "
        "form-action 'self'; frame-ancestors 'none'; base-uri 'self'"
    )
    return response


@app.context_processor
def _template_security_context():
    endpoint = request.endpoint or ""
    step_by_endpoint = {
        "index": 1,
        "upload_files": 2,
        "extend_cross": 2,
        "set_headers": 3,
        "review_headers": 2,
        "clean_data": 3,
        "add_cleansing_rule": 3,
        "remove_cleansing_rule": 3,
        "apply_cleansing_rules": 4,
        "cross_config_page": 4,
        "run_cross": 5,
        "business_rules": 5,
        "add_business_rule": 5,
        "remove_business_rule": 5,
        "apply_business_rules_route": 6,
        "skip_business_rules": 6,
        "show_result": 6,
        "result_stats": 7,
        "configure_statistics": 7,
        "export_configured_statistics": 7,
    }
    return {
        "csrf_token": session.get("csrf_token", ""),
        "current_user": session.get("username"),
        "workflow_step": getattr(g, "workflow_step", step_by_endpoint.get(endpoint)),
        "workflow_steps": [
            {"number": 1, "label": "Archivos"},
            {"number": 2, "label": "Encabezados"},
            {"number": 3, "label": "Limpieza"},
            {"number": 4, "label": "Cruce"},
            {"number": 5, "label": "Columnas calculadas"},
            {"number": 6, "label": "Revisión"},
            {"number": 7, "label": "Estadísticas"},
        ],
    }


@app.route("/login", methods=["GET", "POST"])
def login():
    if app.config["TESTING"]:
        return redirect(url_for("index"))
    if request.method == "GET":
        return render_template("login.html")
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")
    if not username or not password:
        flash("Ingresa usuario y contraseña.")
        return render_template("login.html"), 400
    connection = None
    try:
        connection = _auth_connection()
        auth_result = AuthService(connection).authenticate(username, password)
    except (OSError, RuntimeError, sqlite3.Error):
        _audit.write(
            "unauthenticated",
            "login_fail",
            "error",
            {"reason": "authentication_service_unavailable"},
            user=username or "unknown",
        )
        return _render_error("No se pudo validar el acceso. Contacta al administrador.", 503)
    finally:
        if connection is not None:
            connection.close()
    if (
        not auth_result.success
        or auth_result.user is None
        or (
            auth_result.user.is_test
            and os.getenv("APP_ENV", "production").strip().lower() != "development"
        )
    ):
        _audit.write(
            "unauthenticated",
            "login_fail",
            "rechazado",
            {"reason": "invalid_credentials"},
            user=username or "unknown",
        )
        flash(auth_result.message or "Usuario o contrasena incorrectos.")
        return render_template("login.html"), 401
    session.clear()
    session["username"] = auth_result.user.username
    session["last_activity"] = time.time()
    session["csrf_token"] = secrets.token_urlsafe(32)
    session["workflow_id"] = secrets.token_urlsafe(32)
    g.workflow_id = session["workflow_id"]
    _audit.write(g.workflow_id, "login_ok", "ok", {})
    return redirect(url_for("index"))


@app.post("/logout")
def logout():
    username = session.get("username", "")
    workflow_id = session.get("workflow_id", "unauthenticated")
    try:
        _audit.write(workflow_id, "logout", "ok", {}, user=username or "unknown")
    finally:
        with _workflow_lock:
            _workflows.pop(workflow_id, None)
        session.clear()
    flash("Sesión finalizada.")
    return redirect(url_for("login"))


def _current_workflow() -> dict | None:
    workflow_id = session.get("workflow_id")
    if not workflow_id:
        return None
    with _workflow_lock:
        workflow = _workflows.get(workflow_id)
        if workflow:
            workflow["last_activity"] = time.monotonic()
        return workflow


def _read_upload(
    content: bytes,
    filename: str,
    header_row: int | None = None,
    sheet_index: int = 0,
):
    extension = Path(filename).suffix.lower()
    if extension not in ALLOWED_EXTENSIONS:
        raise FileReadError("Solo se aceptan archivos .xlsx y .xls.")
    if not content or len(content) > MAX_FILE_BYTES:
        raise FileReadError("El archivo esta vacio o supera el limite de 20 MB.")
    if extension == ".xlsx":
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                parts = archive.infolist()
                if (
                    len(parts) > MAX_XLSX_PARTS
                    or sum(part.file_size for part in parts) > MAX_UNCOMPRESSED_XLSX_BYTES
                ):
                    raise FileReadError(
                        "El libro supera el limite de contenido descomprimido permitido."
                    )
        except zipfile.BadZipFile as exc:
            raise FileReadError("El archivo .xlsx no es un libro valido.") from exc

    with tempfile.TemporaryDirectory(prefix="cruzador-web-") as temporary_dir:
        safe_name = secure_filename(filename) or "archivo.xlsx"
        temporary_path = Path(temporary_dir) / safe_name
        temporary_path.write_bytes(content)
        dataframe, metadata = _reader.read(
            temporary_path,
            sheet_index=sheet_index,
            header_row=header_row,
        )
        if len(dataframe) > MAX_ROWS or len(dataframe.columns) > MAX_COLUMNS:
            raise FileReadError(
                f"El archivo supera el limite de {MAX_ROWS:,} filas o {MAX_COLUMNS} columnas."
            )
        return dataframe, replace(metadata, path=None)


def _input_hash(files: list[dict]) -> str:
    hashes = [
        hashlib.sha256(item["content"]).hexdigest()
        for item in files
        if item.get("content") is not None
    ]
    return hashlib.sha256("|".join(hashes).encode("ascii")).hexdigest()


def _read_upload_for_recipe(
    content: bytes,
    filename: str,
    expected_columns: list[str],
) -> tuple[pd.DataFrame, FileMetadata, int]:
    candidates: list[tuple[int, pd.DataFrame, FileMetadata]] = []
    _, initial_metadata = _read_upload(content, filename)
    for sheet_index in range(len(initial_metadata.sheet_names)):
        try:
            dataframe, metadata = _read_upload(
                content, filename, sheet_index=sheet_index
            )
        except (EmptyFileError, FileReadError):
            continue
        if set(expected_columns).issubset(dataframe.columns):
            candidates.append((sheet_index, dataframe, metadata))

    if not candidates:
        expected = ", ".join(expected_columns) or "las columnas detectadas"
        raise FileReadError(
            f"No se encontró una hoja con el encabezado esperado ({expected}) "
            f"en '{filename}'. Revisá el archivo o ejecutá la receta paso a paso."
        )
    if len(candidates) > 1:
        exact = [candidate for candidate in candidates if set(candidate[1].columns) == set(expected_columns)]
        if len(exact) == 1:
            return exact[0][1], replace(exact[0][2], path=None), exact[0][0]
        raise FileReadError(
            f"Hay más de una hoja compatible en '{filename}'. "
            "Ejecutá la receta paso a paso para elegir la hoja."
        )
    sheet_index, dataframe, metadata = candidates[0]
    return dataframe, replace(metadata, path=None), sheet_index


def _render_error(message: str, status: int = 400, technical_type: str = "validation_error"):
    error_id = getattr(g, "error_id", uuid.uuid4().hex[:12].upper())
    audit_available = True
    audit_failure = ""
    try:
        _audit.write(
            getattr(g, "workflow_id", "unknown"),
            "error",
            "error",
            {
                "error_id": error_id,
                "status": status,
                "endpoint": request.endpoint or "unknown",
                "error_type": technical_type,
            },
        )
    except Exception as exc:
        audit_available = False
        audit_failure = f"{type(exc).__name__}: la auditoría no está disponible."
    return render_template(
        "error.html",
        message=message,
        error_id=error_id,
        audit_available=audit_available,
        audit_failure=audit_failure,
    ), status


@app.get("/")
def index():
    workflow = _current_workflow()
    return render_template(
        "upload.html",
        max_files=MAX_FILES,
        max_size_mb=MAX_FILE_BYTES // (1024 * 1024),
        recipes=list_recipes(_recipes_dir()),
        staged_files=workflow.get("files", []) if workflow else [],
    )


@app.post("/upload")
def upload_files():
    with _workflow_lock:
        if (
            g.workflow_id not in _workflows
            and len(_workflows) >= MAX_ACTIVE_WORKFLOWS
        ):
            return _render_error(
                "El prototipo alcanzo el limite de sesiones activas. Intenta mas tarde.",
                503,
            )

    uploads = [item for item in request.files.getlist("files") if item.filename]
    if len(uploads) < 2:
        return _render_error("Selecciona al menos 2 archivos Excel.")
    if len(uploads) > MAX_FILES:
        return _render_error(f"El limite actual es de {MAX_FILES} archivos por cruce.")

    try:
        recipe = _get_recipe(request.form.get("recipe_id", ""))
    except (OSError, ValueError, TypeError) as exc:
        return _render_error("No se pudo leer la configuracion guardada.")
    if request.form.get("recipe_id") and recipe is None:
        return _render_error("La configuracion seleccionada no existe o no es valida.")
    run_automatically = request.form.get("run_mode") == "automatic"
    if run_automatically and recipe is None:
        return _render_error("Selecciona una configuración guardada para ejecutarla automáticamente.")
    if recipe and len(recipe.file_specs) != len(uploads):
        return _render_error(
            f"La configuracion requiere {len(recipe.file_specs)} archivos; seleccionaste {len(uploads)}."
        )

    staged: list[dict] = []
    try:
        for index, uploaded in enumerate(uploads):
            filename = secure_filename(Path(uploaded.filename).name)
            if not filename:
                raise FileReadError("Uno de los archivos no tiene un nombre valido.")
            content = uploaded.stream.read(MAX_FILE_BYTES + 1)
            dataframe, metadata = _read_upload(content, filename)
            base_label = (
                recipe.file_specs[index].label
                if recipe
                else (Path(filename).stem or f"Archivo {index + 1}")
            )
            used_labels = {item["label"] for item in staged}
            label = base_label
            suffix = 2
            while label in used_labels:
                label = f"{base_label} ({suffix})"
                suffix += 1
            staged.append({
                "label": label,
                "filename": filename,
                "content": content,
                "df": dataframe,
                "metadata": metadata,
            })
    except (EmptyFileError, FileReadError) as exc:
        _audit.write(g.workflow_id, "file_loaded", "rechazado", {"reason": str(exc)})
        return _render_error(str(exc))
    except OSError:
        _audit.write(g.workflow_id, "file_loaded", "error", {"reason": "temporary_file_error"})
        return _render_error("No se pudieron preparar los archivos temporalmente.", 500)

    digest = _input_hash(staged)
    _audit.write(
        g.workflow_id,
        "file_loaded",
        "ok",
        {
            "files": [
                {
                    "name": item["filename"],
                    "sheet": item["metadata"].sheet_used,
                    "rows": item["metadata"].row_count,
                    "columns": item["metadata"].column_names,
                    "detected_header_row": item["metadata"].header_row_index + 1,
                }
                for item in staged
            ]
        },
        input_hash=digest,
    )
    with _workflow_lock:
        _workflows[g.workflow_id] = {
            "files": staged,
            "last_activity": time.monotonic(),
            "result": None,
            "input_hash": digest,
            "recipe": recipe,
            "auto_recipe": run_automatically,
            "business_rules": list(recipe.business_rules) if recipe else [],
        }
    if run_automatically:
        return _run_saved_recipe(_workflows[g.workflow_id])
    return render_template("headers.html", files=staged)


@app.get("/headers")
def review_headers():
    workflow = _current_workflow()
    if not workflow or not workflow.get("files"):
        return redirect(url_for("index"))
    return render_template("headers.html", files=workflow["files"])


@app.post("/extend")
def extend_cross():
    workflow = _current_workflow()
    result = workflow.get("result") if workflow else None
    if not result or result.df is None:
        return _render_error("Ejecuta un cruce antes de agregar más archivos.")
    uploads = [item for item in request.files.getlist("files") if item.filename]
    if not uploads:
        return _render_error("Selecciona al menos un archivo Excel para continuar el cruce.")
    if len(uploads) + 1 > MAX_FILES:
        return _render_error(f"El cruce encadenado admite hasta {MAX_FILES} fuentes en total.")

    previous_metadata = FileMetadata(
        path=None,
        sheet_names=["Resultado previo"],
        sheet_used="Resultado previo",
        row_count=len(result.df),
        column_names=list(result.df.columns),
        header_scan_rows=0,
    )
    previous_entry = FileEntry(
        path=None,
        label="Resultado previo",
        sheet_index=-1,
        df=result.df.copy(),
        metadata=previous_metadata,
        is_virtual=True,
    )
    staged = [{
        "label": previous_entry.label,
        "filename": "Resultado previo",
        "content": None,
        "df": previous_entry.df,
        "metadata": previous_metadata,
        "entry": previous_entry,
        "is_virtual": True,
    }]
    try:
        for index, uploaded in enumerate(uploads):
            filename = secure_filename(Path(uploaded.filename).name)
            if not filename:
                raise FileReadError("Uno de los archivos no tiene un nombre valido.")
            content = uploaded.stream.read(MAX_FILE_BYTES + 1)
            dataframe, metadata = _read_upload(content, filename)
            base_label = Path(filename).stem or f"Archivo {index + 1}"
            used_labels = {item["label"] for item in staged}
            label = base_label
            suffix = 2
            while label in used_labels:
                label = f"{base_label} ({suffix})"
                suffix += 1
            staged.append({
                "label": label,
                "filename": filename,
                "content": content,
                "df": dataframe,
                "metadata": metadata,
            })
    except (EmptyFileError, FileReadError) as exc:
        return _render_error(str(exc))
    except OSError:
        return _render_error("No se pudieron preparar los archivos temporalmente.", 500)

    digest = _input_hash(staged)
    combined_hash = hashlib.sha256(
        f"{workflow['input_hash']}|{digest}".encode("ascii")
    ).hexdigest()
    with _workflow_lock:
        _workflows[g.workflow_id] = {
            "files": staged,
            "last_activity": time.monotonic(),
            "result": None,
            "input_hash": combined_hash,
            "previous_input_hash": workflow["input_hash"],
            "recipe": None,
            "is_extension": True,
            "business_rules": [],
        }
    _audit.write(
        g.workflow_id,
        "file_loaded",
        "ok",
        {
            "files": [
                {"name": item["filename"], "sheet": item["metadata"].sheet_used,
                 "rows": item["metadata"].row_count,
                 "columns": item["metadata"].column_names}
                for item in staged
            ],
            "event_type": "sequential_cross",
        },
        input_hash=combined_hash,
    )
    return render_template("headers.html", files=staged)


@app.post("/headers")
def set_headers():
    workflow = _current_workflow()
    if not workflow:
        return _render_error("La sesion de trabajo vencio. Volve a cargar los archivos.", 410)

    staged = workflow["files"]
    entries: list[FileEntry] = []
    try:
        for index, item in enumerate(staged):
            if item.get("is_virtual"):
                entries.append(item["entry"])
                continue
            selected_sheet = int(request.form.get(f"sheet_{index}", "0"))
            header_row = int(request.form.get(f"header_{index}", ""))
            if selected_sheet != -1 and not 0 <= selected_sheet < len(item["metadata"].sheet_names):
                raise ValueError("Selecciona una hoja valida.")
            if not 0 <= header_row < item["metadata"].header_scan_rows:
                raise ValueError("Selecciona una fila de encabezado valida.")
            dataframe, metadata = _read_upload(
                item["content"], item["filename"], header_row, selected_sheet
            )
            entry = FileEntry(
                path=None,
                label=item["label"],
                sheet_index=selected_sheet,
                df=dataframe,
                metadata=metadata,
            )
            item["df"] = dataframe
            item["metadata"] = metadata
            entries.append(entry)
    except (ValueError, EmptyFileError, FileReadError) as exc:
        return _render_error(str(exc))

    workflow["entries"] = entries
    _audit.write(
        g.workflow_id,
        "columns_selected",
        "ok",
        {
            "header_rows": [entry.metadata.header_row_index + 1 for entry in entries],
            "column_counts": [len(entry.metadata.column_names) for entry in entries],
        },
        input_hash=workflow["input_hash"],
    )
    workflow["source_entries"] = entries
    recipe = workflow.get("recipe")
    workflow["cleansing_rules"] = list(recipe.cleansing_rules) if recipe else []
    recipe_warnings = []
    if recipe:
        from engine.recipe import validate_recipe_against_entries
        recipe_warnings = validate_recipe_against_entries(recipe, entries)
    return render_template(
        "cleanse.html",
        entries=entries,
        rules=workflow["cleansing_rules"],
        operations=sorted(OPERACIONES_VALIDAS),
        recipe_warnings=recipe_warnings,
    )


def _cleansing_rule_from_form(form) -> CleansingRule:
    workflow = _current_workflow()
    entries: list[FileEntry] = workflow["source_entries"]
    entry_index = int(form.get("file_index", "-1"))
    if entry_index < 0 or entry_index >= len(entries):
        raise ValueError("Selecciona un archivo valido.")
    entry = entries[entry_index]
    column_value = form.get("column", "")
    operation = form.get("operation", "")
    try:
        column_index_text, column = column_value.split("::", 1)
        column_index = int(column_index_text)
    except (ValueError, AttributeError) as exc:
        raise ValueError("Selecciona una columna valida.") from exc
    if column_index != entry_index:
        raise ValueError("La columna no corresponde al archivo seleccionado.")
    if column not in entry.df.columns or operation not in OPERACIONES_VALIDAS:
        raise ValueError("Selecciona una columna y una operacion validas.")

    params: dict = {}
    if operation == "fill_empty":
        params["value"] = form.get("value", "")
    elif operation in {"replace_literal", "replace_regex"}:
        key = "find" if operation == "replace_literal" else "pattern"
        params[key] = form.get(key, "")
        params["replace"] = form.get("replace", "")
    elif operation == "remove_chars":
        params["chars"] = form.get("chars", "")
    elif operation == "pad_left":
        try:
            params["width"] = int(form.get("width", "0"))
        except ValueError as exc:
            raise ValueError("El ancho de relleno debe ser un numero entero.") from exc
        params["char"] = form.get("char", "0")
    elif operation == "normalize_date":
        params["format_from"] = form.get("format_from", "%d/%m/%Y")
        params["format_to"] = form.get("format_to", "%Y-%m-%d")
    elif operation in {"filter_date_before", "filter_date_after"}:
        params["date"] = form.get("date", "")
    elif operation == "filter_date_range":
        params["date_from"] = form.get("date_from", "")
        params["date_to"] = form.get("date_to", "")
    if operation in {"filter_date_before", "filter_date_after"} and not parse_date_flexible(
        str(params.get("date", ""))
    ):
        raise ValueError("La fecha de corte no tiene un formato reconocido.")
    if operation == "filter_date_range":
        date_from = parse_date_flexible(str(params["date_from"]))
        date_to = parse_date_flexible(str(params["date_to"]))
        if date_from is None or date_to is None or date_from > date_to:
            raise ValueError("El rango de fechas es inválido.")

    if operation == "pad_left" and (params["width"] < 0 or params["width"] > 64):
        raise ValueError("El ancho de relleno debe estar entre 0 y 64.")
    rule = CleansingRule.new(entry.label, column, operation, params)
    if operation not in {"drop_column", "filter_date_before", "filter_date_after", "filter_date_range"}:
        apply_rule(entry.df[column].head(5), rule)
    return rule


@app.get("/clean")
def clean_data():
    workflow = _current_workflow()
    if not workflow or "source_entries" not in workflow:
        return _render_error("La sesion de trabajo vencio. Volve a cargar los archivos.", 410)
    return render_template(
        "cleanse.html",
        entries=workflow["source_entries"],
        rules=workflow["cleansing_rules"],
        operations=sorted(OPERACIONES_VALIDAS),
    )


@app.get("/cross-config")
def cross_config_page():
    workflow = _current_workflow()
    entries = workflow.get("entries") if workflow else None
    if not entries:
        return redirect(url_for("clean_data"))
    suggestions = [
        suggest_pairs(left.metadata.column_names, right.metadata.column_names)
        for left, right in zip(entries, entries[1:])
    ]
    config = workflow.get("cross_config", {})
    selected_pairs = config.get("steps_key_pairs") or suggestions
    recipe = workflow.get("recipe")
    return render_template(
        "keys.html",
        files=workflow["files"],
        entries=entries,
        suggested_pairs=suggestions,
        selected_pairs=selected_pairs,
        warnings=[],
        confirm_duplicates=False,
        selected_join=config.get("join_type") or (recipe.join_type if recipe else None),
        normalize=config.get("normalize_keys", recipe.normalize_keys if recipe else True),
    )


@app.post("/clean/add")
def add_cleansing_rule():
    workflow = _current_workflow()
    if not workflow or "source_entries" not in workflow:
        return _render_error("La sesion de trabajo vencio. Volve a cargar los archivos.", 410)
    try:
        rule = _cleansing_rule_from_form(request.form)
    except (ValueError, TypeError, KeyError) as exc:
        return _render_error(str(exc))
    workflow["cleansing_rules"].append(rule)
    return redirect(url_for("clean_data"))


@app.post("/clean/remove/<int:rule_index>")
def remove_cleansing_rule(rule_index: int):
    workflow = _current_workflow()
    if not workflow or "cleansing_rules" not in workflow:
        return _render_error("La sesion de trabajo vencio. Volve a cargar los archivos.", 410)
    rules = workflow["cleansing_rules"]
    if not 0 <= rule_index < len(rules):
        return _render_error("La regla seleccionada ya no existe.")
    rules.pop(rule_index)
    return redirect(url_for("clean_data"))


@app.post("/clean/apply")
def apply_cleansing_rules():
    workflow = _current_workflow()
    if not workflow or "source_entries" not in workflow:
        return _render_error("La sesion de trabajo vencio. Volve a cargar los archivos.", 410)
    rules: list[CleansingRule] = workflow["cleansing_rules"]
    try:
        cleaned_entries = apply_all_rules(
            workflow["source_entries"], rules, strict=True
        )
    except (ValueError, TypeError, KeyError) as exc:
        _audit.write(
            g.workflow_id,
            "error",
            "error",
            {"type": type(exc).__name__, "message": "cleansing failed"},
            input_hash=workflow["input_hash"],
        )
        return _render_error("No se pudieron aplicar las reglas de limpieza.")
    workflow["entries"] = cleaned_entries
    suggestions = [
        suggest_pairs(left.metadata.column_names, right.metadata.column_names)
        for left, right in zip(cleaned_entries, cleaned_entries[1:])
    ]
    recipe = workflow.get("recipe")
    recipe_warnings = []
    selected_pairs = []
    if recipe:
        from engine.recipe import validate_recipe_against_entries
        recipe_warnings = validate_recipe_against_entries(recipe, cleaned_entries)
        for left, right, suggested in zip(
            cleaned_entries, cleaned_entries[1:], suggestions
        ):
            step = next(
                (
                    item
                    for item in recipe.cross_steps
                    if item.left_label == left.label and item.right_label == right.label
                ),
                None,
            )
            step_pairs = step.key_pairs if step else []
            if any(
                left_col not in left.df.columns or right_col not in right.df.columns
                for left_col, right_col in step_pairs
            ):
                step_pairs = []
            selected_pairs.append(step_pairs[:5] or suggested[:5])
    _audit.write(
        g.workflow_id,
        "process_finished",
        "ok",
        {
            "event_type": "cleansing",
            "rule_count": len(rules),
            "rows_by_file": [len(entry.df) for entry in cleaned_entries],
            "operations": [rule.operation for rule in rules],
        },
        input_hash=workflow["input_hash"],
    )
    return render_template(
        "keys.html",
        files=workflow["files"],
        entries=cleaned_entries,
        suggested_pairs=selected_pairs,
        selected_pairs=selected_pairs,
        warnings=recipe_warnings,
        confirm_duplicates=False,
        selected_join=recipe.join_type if recipe else None,
        normalize=recipe.normalize_keys if recipe else True,
    )


def _execute_cross_configuration(
    workflow: dict,
    steps_key_pairs: list[list[tuple[str, str]]],
    join_type: str,
    normalize: bool,
):
    _audit.write(
        g.workflow_id,
        "process_started",
        "ok",
        {
            "files": [item["filename"] for item in workflow["files"]],
            "columns": [column for step in steps_key_pairs for pair in step for column in pair],
            "join_type": join_type,
            "event_type": "cross",
        },
        input_hash=workflow["input_hash"],
    )
    started = time.monotonic()
    try:
        result = execute_cross(CrossConfig(
            file_entries=workflow["entries"],
            key_pairs=[step[0] for step in steps_key_pairs],
            join_type=join_type,
            normalize_keys=normalize,
            steps_key_pairs=steps_key_pairs,
        ))
    except (CrossConfigError, ValueError, TypeError, KeyError) as exc:
        _audit.write(
            g.workflow_id,
            "error",
            "error",
            {"type": type(exc).__name__, "message": "cross execution failed"},
            input_hash=workflow["input_hash"],
        )
        return None

    workflow["result"] = result
    workflow["cross_config"] = {
        "steps_key_pairs": steps_key_pairs,
        "join_type": join_type,
        "normalize_keys": normalize,
    }
    _audit.write(
        g.workflow_id,
        "process_finished",
        "ok",
        {
            "input_rows": [entry.metadata.row_count for entry in workflow["entries"]],
            "output_rows": result.row_count,
            "matched_rows": result.matched_rows,
            "duration_ms": round((time.monotonic() - started) * 1000, 2),
            "warnings_count": len(result.warnings),
        },
        input_hash=workflow["input_hash"],
    )
    return result


def _run_saved_recipe(workflow: dict):
    recipe = workflow.get("recipe")
    if recipe is None:
        return _render_error("No hay una configuración guardada para ejecutar.")
    staged = workflow["files"]
    if len(staged) != len(recipe.file_specs):
        return _render_error("La cantidad de archivos no coincide con la configuración.")

    entries: list[FileEntry] = []
    try:
        for item, file_spec in zip(staged, recipe.file_specs):
            dataframe, metadata, sheet_index = _read_upload_for_recipe(
                item["content"], item["filename"], file_spec.expected_columns
            )
            item["df"] = dataframe
            item["metadata"] = metadata
            item["selected_sheet"] = sheet_index
            entries.append(FileEntry(
                path=None,
                label=file_spec.label,
                sheet_index=sheet_index,
                df=dataframe,
                metadata=metadata,
            ))
        from engine.recipe import validate_recipe_against_entries
        warnings = validate_recipe_against_entries(recipe, entries)
        if warnings:
            raise ValueError("; ".join(warnings))
        cleaned_entries = apply_all_rules(
            entries, recipe.cleansing_rules, strict=True
        )
    except (EmptyFileError, FileReadError, OSError, ValueError, TypeError, KeyError) as exc:
        workflow["auto_recipe"] = False
        _audit.write(
            g.workflow_id,
            "process_finished",
            "error",
            {"event_type": "recipe_validation", "type": type(exc).__name__},
            input_hash=workflow["input_hash"],
        )
        flash(
            "No se pudo ejecutar automáticamente la configuración. "
            "Podés continuar manualmente con los archivos ya cargados."
        )
        return render_template(
            "headers.html",
            files=staged,
            recipe_error=str(exc),
        )

    if len(recipe.cross_steps) != max(0, len(cleaned_entries) - 1):
        workflow["auto_recipe"] = False
        flash(
            "La configuración no define todos los pasos de cruce. "
            "Continuá manualmente para revisar y completar la configuración."
        )
        return render_template(
            "headers.html",
            files=staged,
            recipe_error="La cantidad de pasos guardados no coincide con los archivos.",
        )

    steps_key_pairs: list[list[tuple[str, str]]] = []
    duplicate_warnings: list[str] = []
    try:
        for index, (left_entry, right_entry, recipe_step) in enumerate(
            zip(cleaned_entries, cleaned_entries[1:], recipe.cross_steps)
        ):
            if (
                recipe_step.left_label != left_entry.label
                or recipe_step.right_label != right_entry.label
                or not recipe_step.key_pairs
                or len(recipe_step.key_pairs) > 5
            ):
                raise ValueError(f"El paso {index + 1} no coincide con el orden de archivos.")
            if any(
                left not in left_entry.df.columns or right not in right_entry.df.columns
                for left, right in recipe_step.key_pairs
            ):
                raise ValueError(f"Falta una columna clave del paso {index + 1}.")
            pairs = list(recipe_step.key_pairs)
            steps_key_pairs.append(pairs)
            for entry, columns in (
                (left_entry, [pair[0] for pair in pairs]),
                (right_entry, [pair[1] for pair in pairs]),
            ):
                duplicate_rows = int(entry.df.duplicated(subset=columns, keep=False).sum())
                if duplicate_rows:
                    duplicate_warnings.append(
                        f"{entry.label}: {duplicate_rows} filas repiten la clave configurada."
                    )
    except (ValueError, KeyError, TypeError) as exc:
        workflow["auto_recipe"] = False
        flash(
            "La receta no se puede aplicar exactamente a estos archivos. "
            "Continuá manualmente para revisar las columnas."
        )
        return render_template(
            "headers.html",
            files=staged,
            recipe_error=str(exc),
        )

    workflow["source_entries"] = entries
    workflow["entries"] = cleaned_entries
    workflow["cleansing_rules"] = list(recipe.cleansing_rules)
    workflow["business_rules"] = list(recipe.business_rules)
    workflow["cross_config"] = {
        "steps_key_pairs": steps_key_pairs,
        "join_type": recipe.join_type,
        "normalize_keys": recipe.normalize_keys,
    }
    _audit.write(
        g.workflow_id,
        "columns_selected",
        "ok",
        {
            "header_rows": [entry.metadata.header_row_index + 1 for entry in entries],
            "column_counts": [len(entry.metadata.column_names) for entry in entries],
            "recipe_id": recipe.recipe_id,
        },
        input_hash=workflow["input_hash"],
    )
    _audit.write(
        g.workflow_id,
        "process_finished",
        "ok",
        {
            "event_type": "recipe_cleansing",
            "rule_count": len(recipe.cleansing_rules),
            "rows_by_file": [len(entry.df) for entry in cleaned_entries],
        },
        input_hash=workflow["input_hash"],
    )
    if duplicate_warnings:
        signature = json.dumps(
            steps_key_pairs, ensure_ascii=False, separators=(",", ":")
        )
        g.workflow_step = 4
        return render_template(
            "keys.html",
            files=staged,
            entries=cleaned_entries,
            suggested_pairs=steps_key_pairs,
            selected_pairs=steps_key_pairs,
            key_signature=signature,
            warnings=duplicate_warnings,
            confirm_duplicates=True,
            selected_join=recipe.join_type,
            normalize=recipe.normalize_keys,
        )

    result = _execute_cross_configuration(
        workflow, steps_key_pairs, recipe.join_type, recipe.normalize_keys
    )
    if result is None:
        return _render_error(
            "No se pudo ejecutar el cruce guardado. La sesión conserva los archivos para continuar manualmente."
        )
    return _apply_recipe_business_rules(workflow)


def _apply_recipe_business_rules(workflow: dict):
    result = workflow["result"]
    rules = workflow.get("business_rules", [])
    try:
        result_df = result.df.copy()
        for rule in rules:
            result_df = apply_business_rule(result_df, rule)
    except (ValueError, TypeError, KeyError) as exc:
        _audit.write(
            g.workflow_id,
            "error",
            "error",
            {"type": type(exc).__name__, "message": "recipe business rules failed"},
            input_hash=workflow["input_hash"],
        )
        return _render_error(
            "El cruce terminó, pero no se pudieron aplicar todas las columnas calculadas. "
            "Volvé a la configuración para continuar."
        )
    workflow["result"] = replace(result, df=result_df, row_count=len(result_df))
    _audit.write(
        g.workflow_id,
        "process_finished",
        "ok",
        {
            "event_type": "recipe_business_rules",
            "rule_count": len(rules),
            "rows": len(result_df),
        },
        input_hash=workflow["input_hash"],
    )
    flash(
        f"Se ejecutó automáticamente la configuración '{workflow['recipe'].name}'. "
        "Revisá el resultado antes de exportarlo."
    )
    return redirect(url_for("show_result"))


@app.post("/run")
def run_cross():
    workflow = _current_workflow()
    if not workflow or "entries" not in workflow:
        return _render_error("La sesion de trabajo vencio. Volve a cargar los archivos.", 410)

    entries: list[FileEntry] = workflow["entries"]
    steps_key_pairs: list[list[tuple[str, str]]] = []
    try:
        for step_index, (left_entry, right_entry) in enumerate(zip(entries, entries[1:])):
            left_slots = request.form.getlist(f"left_key_{step_index}")
            right_slots = request.form.getlist(f"right_key_{step_index}")
            if len(left_slots) != len(right_slots):
                raise ValueError(f"El paso {step_index + 1} tiene una configuracion incompleta.")
            if any(bool(left) != bool(right) for left, right in zip(left_slots, right_slots)):
                raise ValueError(
                    f"Completa ambos lados de cada par de clave en el paso {step_index + 1}."
                )
            step_pairs = [
                (left, right)
                for left, right in zip(left_slots, right_slots)
                if left and right
            ]
            if not step_pairs:
                raise ValueError(f"Selecciona al menos un par de claves en el paso {step_index + 1}.")
            if any(column not in left_entry.df.columns for column, _ in step_pairs):
                raise ValueError(f"Hay una columna izquierda invalida en el paso {step_index + 1}.")
            if any(column not in right_entry.df.columns for _, column in step_pairs):
                raise ValueError(f"Hay una columna derecha invalida en el paso {step_index + 1}.")
            steps_key_pairs.append(step_pairs)
    except ValueError as exc:
        selected_pairs = []
        for step_index in range(max(0, len(entries) - 1)):
            left_slots = request.form.getlist(f"left_key_{step_index}")
            right_slots = request.form.getlist(f"right_key_{step_index}")
            selected_pairs.append(list(zip(left_slots, right_slots)))
        _audit.write(
            g.workflow_id,
            "error",
            "rechazado",
            {"error_type": "cross_keys_incomplete", "step": "cross_configuration", "status": 400},
        )
        return render_template(
            "keys.html",
            files=workflow["files"],
            entries=entries,
            suggested_pairs=selected_pairs,
            selected_pairs=selected_pairs,
            warnings=[],
            confirm_duplicates=False,
            validation_error=str(exc),
            selected_join=request.form.get("join_type"),
            normalize=request.form.get("normalize") == "yes",
        ), 200

    join_type = request.form.get("join_type", "")
    normalize = request.form.get("normalize") == "yes"
    if join_type not in {"inner", "left", "outer"}:
        return _render_error("El tipo de cruce seleccionado no es valido.")

    duplicate_warnings = []
    for step_index, step_pairs in enumerate(steps_key_pairs):
        for entry_index, columns in (
            (step_index, [pair[0] for pair in step_pairs]),
            (step_index + 1, [pair[1] for pair in step_pairs]),
        ):
            entry = entries[entry_index]
            duplicate_rows = int(entry.df.duplicated(subset=columns, keep=False).sum())
            if duplicate_rows:
                duplicate_warnings.append(
                    f"{entry.label}: {duplicate_rows} filas repiten la clave compuesta seleccionada."
                )
    confirmation_matches = (
        request.form.get("confirm_duplicates") == "yes"
        and request.form.get("confirmed_key_signature")
        == json.dumps(steps_key_pairs, ensure_ascii=False, separators=(",", ":"))
    )
    if duplicate_warnings and not confirmation_matches:
        g.workflow_step = 4
        return render_template(
            "keys.html",
            files=workflow["files"],
            entries=entries,
            suggested_pairs=steps_key_pairs,
            selected_pairs=steps_key_pairs,
            key_signature=json.dumps(
                steps_key_pairs, ensure_ascii=False, separators=(",", ":")
            ),
            warnings=duplicate_warnings,
            confirm_duplicates=True,
            selected_join=join_type,
            normalize=normalize,
        )

    result = _execute_cross_configuration(
        workflow, steps_key_pairs, join_type, normalize
    )
    if result is None:
        return _render_error("No se pudo completar el cruce. Revisa las columnas y los archivos.")

    recipe = workflow.get("recipe")
    if recipe:
        workflow["business_rules"] = list(recipe.business_rules)
    else:
        workflow.setdefault("business_rules", [])
    if workflow.get("auto_recipe"):
        return _apply_recipe_business_rules(workflow)
    return render_template(
        "rules.html",
        result=result,
        columns=list(result.df.columns) if result.df is not None else [],
        rules=workflow["business_rules"],
        operators=sorted(CONDITION_OPERATORS),
    )


@app.get("/rules")
def business_rules():
    workflow = _current_workflow()
    if workflow is not None:
        workflow.setdefault("business_rules", [])
    result = workflow.get("result") if workflow else None
    if not result or result.df is None:
        return _render_error("Primero ejecuta el cruce.", 410)
    return render_template(
        "rules.html",
        result=result,
        columns=list(result.df.columns),
        rules=workflow["business_rules"],
        operators=sorted(CONDITION_OPERATORS),
    )


def _business_rule_from_form(form) -> BusinessRule:
    output_column = form.get("output_column", "").strip()
    rule_type = form.get("rule_type", "")
    if not output_column:
        raise ValueError("Indica el nombre de la columna nueva.")
    if rule_type == "formula":
        params = {
            "left": form.get("left", "").strip(),
            "operator": form.get("formula_operator", ""),
            "right": form.get("right", "").strip(),
            "format": form.get("format", "number"),
        }
    elif rule_type == "date_diff":
        params = {
            "col_a": form.get("col_a", ""),
            "col_b": form.get("col_b", ""),
            "unit": form.get("unit", "days"),
            "absolute": form.get("absolute") == "yes",
        }
    elif rule_type == "if_then_else":
        columns = form.getlist("condition_column")
        operators = form.getlist("condition_operator")
        values = form.getlist("condition_value")
        if not (len(columns) == len(operators) == len(values)):
            raise ValueError("Las condiciones están incompletas.")
        conditions = [
            {"column": column, "operator": operator, "value": value}
            for column, operator, value in zip(columns, operators, values)
            if column
        ]
        if not conditions or any(
            condition["column"] not in workflow_result_columns()
            or condition["operator"] not in CONDITION_OPERATORS
            for condition in conditions
        ):
            raise ValueError("Configura al menos una condición válida.")
        params = {
            "conditions": conditions,
            "combinator": form.get("combinator", "AND"),
            "value_true": form.get("value_true", ""),
            "value_false": form.get("value_false", ""),
        }
        if params["combinator"] not in {"AND", "OR"}:
            raise ValueError("El combinador debe ser AND u OR.")
    else:
        raise ValueError("El tipo de regla no es válido.")
    return BusinessRule.new(output_column, rule_type, params)


def workflow_result_columns() -> list[str]:
    workflow = _current_workflow()
    result = workflow.get("result") if workflow else None
    columns = list(result.df.columns) if result and result.df is not None else []
    if workflow:
        columns.extend(
            rule.output_column
            for rule in workflow.get("business_rules", [])
            if rule.output_column not in columns
        )
    return columns


@app.post("/rules/add")
def add_business_rule():
    workflow = _current_workflow()
    if workflow is not None:
        workflow.setdefault("business_rules", [])
    result = workflow.get("result") if workflow else None
    if not result or result.df is None:
        return _render_error("Primero ejecuta el cruce.", 410)
    try:
        rule = _business_rule_from_form(request.form)
        if rule.output_column in workflow_result_columns():
            raise ValueError("El nombre de la columna calculada ya existe.")
        test_df = result.df.copy()
        for existing_rule in workflow["business_rules"]:
            test_df = apply_business_rule(test_df, existing_rule)
        apply_business_rule(test_df, rule)
    except (ValueError, TypeError, KeyError) as exc:
        return _render_error(str(exc))
    workflow["business_rules"].append(rule)
    return redirect(url_for("business_rules"))


@app.post("/rules/remove/<int:rule_index>")
def remove_business_rule(rule_index: int):
    workflow = _current_workflow()
    if not workflow or "business_rules" not in workflow:
        return _render_error("La sesion de trabajo vencio. Volve a cargar los archivos.", 410)
    rules = workflow["business_rules"]
    if not 0 <= rule_index < len(rules):
        return _render_error("La regla seleccionada ya no existe.")
    rules.pop(rule_index)
    return redirect(url_for("business_rules"))


@app.post("/rules/apply")
def apply_business_rules_route():
    workflow = _current_workflow()
    if workflow is not None:
        workflow.setdefault("business_rules", [])
    result = workflow.get("result") if workflow else None
    if not result or result.df is None:
        return _render_error("Primero ejecuta el cruce.", 410)
    rules: list[BusinessRule] = workflow["business_rules"]
    try:
        result_df = result.df.copy()
        for rule in rules:
            result_df = apply_business_rule(result_df, rule)
    except (ValueError, TypeError, KeyError) as exc:
        _audit.write(
            g.workflow_id,
            "error",
            "error",
            {"type": type(exc).__name__, "message": "business rule execution failed"},
            input_hash=workflow["input_hash"],
        )
        return _render_error("No se pudieron aplicar todas las reglas. Revisa su configuración.")
    workflow["result"] = replace(result, df=result_df, row_count=len(result_df))
    _audit.write(
        g.workflow_id,
        "process_finished",
        "ok",
        {
            "event_type": "business_rules",
            "rule_count": len(rules),
            "output_columns": [rule.output_column for rule in rules],
            "rows": len(result_df),
        },
        input_hash=workflow["input_hash"],
    )
    return redirect(url_for("show_result"))


@app.post("/rules/skip")
def skip_business_rules():
    return redirect(url_for("show_result"))


def _display_value(value):
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.isoformat(sep=" ", timespec="seconds")
    return str(value)


def _filtered_result_frame(workflow, result, quick_search, conditions):
    displayed = result.df
    if conditions or quick_search:
        displayed = apply_filters(result.df, conditions, quick_search=quick_search)
    return displayed


@app.get("/result")
def show_result():
    workflow = _current_workflow()
    result = workflow.get("result") if workflow else None
    if not result or result.df is None:
        return _render_error("No hay un resultado de cruce disponible.", 410)
    current_filters = workflow.get("result_filters", [])
    quick_search = workflow.get("result_quick_search", "")
    displayed = result.df
    if current_filters or quick_search:
        displayed = apply_filters(
            result.df,
            [
                FilterCondition(
                    column=item["column"],
                    operator=item["operator"],
                    value=item["value"],
                    value2=item["value2"],
                )
                for item in current_filters
            ],
            quick_search=quick_search,
        )
    workflow["displayed_indices"] = list(displayed.index)
    preview_frame = displayed.head(50)
    preview = [
        {
            "source_index": int(row_index),
            "values": {
                column: _display_value(row[column])
                for column in result.df.columns
            },
        }
        for row_index, row in preview_frame.iterrows()
    ]
    return render_template(
        "result.html",
        result=result,
        columns=list(result.df.columns) if result.df is not None else [],
        rows=preview,
        preview_truncated=len(displayed) > len(preview),
        filtered_count=len(displayed),
        current_filters=current_filters,
        quick_search=quick_search,
        filter_operators=sorted(OPERADORES_VALIDOS),
    )


@app.get("/result/preview")
def preview_result():
    workflow = _current_workflow()
    result = workflow.get("result") if workflow else None
    if not result or result.df is None:
        return jsonify({"error": "No hay un resultado disponible."}), 410

    conditions = []
    for column, operator, value, value2 in zip(
        request.args.getlist("column"),
        request.args.getlist("operator"),
        request.args.getlist("value"),
        request.args.getlist("value2"),
    ):
        if column and column in result.df.columns and operator in OPERADORES_VALIDOS:
            conditions.append(FilterCondition(column, operator, value, value2))

    quick_search = request.args.get("quick_search", "")
    displayed = _filtered_result_frame(workflow, result, quick_search, conditions)
    preview_frame = displayed.head(50)
    rows = []
    for row_index, row in preview_frame.iterrows():
        rows.append({
            "source_index": int(row_index),
            "values": {
                column: _display_value(row[column])
                for column in result.df.columns
            },
        })
    return jsonify({
        "columns": list(result.df.columns),
        "rows": rows,
        "filtered_count": len(displayed),
        "preview_truncated": len(displayed) > len(rows),
    })


@app.post("/result/filter")
def filter_result():
    workflow = _current_workflow()
    result = workflow.get("result") if workflow else None
    if not result or result.df is None:
        return _render_error("No hay un resultado de cruce disponible.", 410)
    columns = request.form.getlist("column")
    operators = request.form.getlist("operator")
    values = request.form.getlist("value")
    values2 = request.form.getlist("value2")
    if not (len(columns) == len(operators) == len(values) == len(values2)):
        return _render_error("La configuración de filtros está incompleta.")
    filters = []
    for column, operator, value, value2 in zip(columns, operators, values, values2):
        if not column:
            continue
        if column not in result.df.columns or operator not in OPERADORES_VALIDOS:
            return _render_error("Hay un filtro seleccionado que no es válido.")
        filters.append({
            "column": column,
            "operator": operator,
            "value": value,
            "value2": value2,
        })
    if filters:
        workflow["result_filters"] = filters
    else:
        workflow.pop("result_filters", None)
    quick_search = request.form.get("quick_search", "")
    if quick_search:
        workflow["result_quick_search"] = quick_search
    else:
        workflow.pop("result_quick_search", None)
    return redirect(url_for("show_result"))


@app.post("/result/edit")
def edit_result_cell():
    workflow = _current_workflow()
    result = workflow.get("result") if workflow else None
    if not result or result.df is None:
        return _render_error("No hay un resultado de cruce disponible.", 410)
    try:
        row_index = int(request.form.get("row_index", ""))
    except ValueError:
        return _render_error("Indica una fila válida.")
    column = request.form.get("column", "")
    if row_index not in result.df.index or column not in result.df.columns:
        return _render_error("La fila o columna seleccionada no existe.")
    updated = result.df.copy()
    updated.at[row_index, column] = request.form.get("value", "")
    workflow["result"] = replace(result, df=updated)
    _audit.write(
        g.workflow_id,
        "process_finished",
        "ok",
        {"event_type": "result_cell_edited", "row": row_index, "column": column},
        input_hash=workflow["input_hash"],
    )
    return redirect(url_for("show_result"))


@app.post("/result/delete/<int:row_index>")
def delete_result_row(row_index: int):
    workflow = _current_workflow()
    result = workflow.get("result") if workflow else None
    if not result or result.df is None or row_index not in result.df.index:
        return _render_error("La fila seleccionada no existe.", 404)
    updated = result.df.drop(index=row_index).copy()
    workflow["result"] = replace(
        result,
        df=updated,
        row_count=len(updated),
    )
    workflow.pop("result_filters", None)
    workflow.pop("result_quick_search", None)
    _audit.write(
        g.workflow_id,
        "process_finished",
        "ok",
        {"event_type": "result_row_deleted", "rows_remaining": len(updated)},
        input_hash=workflow["input_hash"],
    )
    return redirect(url_for("show_result"))


@app.post("/result/clear-filter")
def clear_result_filter():
    workflow = _current_workflow()
    if workflow:
        workflow.pop("result_filters", None)
        workflow.pop("result_quick_search", None)
    return redirect(url_for("show_result"))


@app.get("/stats")
def result_stats():
    workflow = _current_workflow()
    result = workflow.get("result") if workflow else None
    if not result or result.df is None:
        return _render_error("No hay un resultado de cruce disponible.", 410)
    stats = compute_stats(
        result.df,
        matched_rows=result.matched_rows,
        left_only_rows=result.left_only_rows,
        right_only_rows=result.right_only_rows,
    )
    _audit.write(
        g.workflow_id,
        "process_finished",
        "ok",
        {
            "event_type": "statistics",
            "rows": stats.total_rows,
            "columns": stats.total_cols,
            "duration_ms": round(stats.duration_ms, 2),
        },
        input_hash=workflow["input_hash"],
    )
    return render_template(
        "stats.html",
        result=result,
        stats=stats,
        columns=list(result.df.columns),
        configured=workflow.get("configured_statistics"),
        configured_columns=workflow.get("configured_statistics_columns", []),
        configured_rows=workflow.get("configured_statistics_rows", []),
        measure_operations=[
            ("count_rows", "Cantidad de filas"),
            ("count_nonempty", "Valores no vacíos"),
            ("count_distinct", "Valores distintos"),
            ("empty", "Valores vacíos"),
            ("sum", "Suma"),
            ("mean", "Promedio"),
            ("min", "Mínimo"),
            ("max", "Máximo"),
        ],
    )


def _parse_decimal(value: object) -> Decimal | None:
    if pd.isna(value):
        return None
    text = str(value).strip().replace(" ", "")
    if not text:
        return None
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def _calculate_configured_statistics(
    df: pd.DataFrame,
    dimensions: list[str],
    measures: list[tuple[str, str]],
) -> tuple[list[str], list[dict]]:
    descriptive = {column.name: column for column in compute_stats(df).columns}
    for column, operation in measures:
        if operation not in {"count_rows", "count_nonempty", "count_distinct", "empty", "sum", "mean", "min", "max"}:
            raise ValueError("Hay una operación estadística no válida.")
        if operation != "count_rows" and column not in df.columns:
            raise ValueError("Selecciona una columna válida para cada medida.")
        if operation in {"sum", "mean", "min", "max"} and descriptive[column].is_identifier:
            raise ValueError(
                f"La columna '{column}' está identificada como identificador y no puede resumirse como medida numérica."
            )

    output_columns = [*dimensions]
    measure_names = [
        "Filas" if operation == "count_rows" else f"{operation}: {column}"
        for column, operation in measures
    ]
    output_columns.extend(measure_names)
    if not measures:
        raise ValueError("Agrega al menos una medida estadística.")

    if dimensions:
        grouped = df.groupby(dimensions, dropna=False, sort=False)
        groups = list(grouped)
    else:
        groups = [("Total", df)]

    rows: list[dict] = []
    for group_key, group_df in groups:
        key_values = (
            tuple(group_key) if dimensions and isinstance(group_key, tuple)
            else (group_key,) if dimensions else ()
        )
        row = {
            column: ("(Vacío)" if pd.isna(value) else value)
            for column, value in zip(dimensions, key_values)
        }
        for (column, operation), output_name in zip(measures, measure_names):
            if operation == "count_rows":
                row[output_name] = len(group_df)
                continue
            series = group_df[column]
            empty_mask = series.isna() | (series.astype(str).str.strip() == "")
            valid = series[~empty_mask]
            if operation == "count_nonempty":
                row[output_name] = int(len(valid))
            elif operation == "count_distinct":
                row[output_name] = int(valid.nunique())
            elif operation == "empty":
                row[output_name] = int(empty_mask.sum())
            else:
                is_date = descriptive[column].dtype_label == "fecha"
                if is_date and operation not in {"min", "max"}:
                    raise ValueError(f"La columna '{column}' es de fecha; solo admite mínimo o máximo.")
                if is_date:
                    values = [
                        parse_date_flexible(str(value))
                        for value in valid.tolist()
                    ]
                else:
                    values = [_parse_decimal(value) for value in valid.tolist()]
                parsed = [value for value in values if value is not None]
                invalid_count = len(values) - len(parsed)
                if operation == "sum":
                    aggregate = sum(parsed, Decimal("0"))
                elif operation == "mean":
                    aggregate = sum(parsed, Decimal("0")) / len(parsed) if parsed else None
                elif operation == "min":
                    aggregate = min(parsed) if parsed else None
                else:
                    aggregate = max(parsed) if parsed else None
                if is_date and aggregate is not None:
                    aggregate = aggregate.strftime("%d/%m/%Y")
                row[output_name] = (
                    f"{aggregate} (válidos: {len(parsed)}, inválidos: {invalid_count})"
                    if aggregate is not None
                    else f"Sin valores válidos (inválidos: {invalid_count})"
                )
        rows.append(row)
    if dimensions:
        rows.sort(key=lambda row: tuple(str(row[column]) for column in dimensions))
    return output_columns, rows


@app.post("/stats/configure")
def configure_statistics():
    workflow = _current_workflow()
    result = workflow.get("result") if workflow else None
    if not result or result.df is None:
        return _render_error("No hay un resultado de cruce disponible.", 410)
    dimensions = request.form.getlist("dimensions")
    if any(column not in result.df.columns for column in dimensions):
        return _render_error("Hay una dimensión que no pertenece al resultado.")
    if dimensions and request.form.get("confirm_dimensions") != "yes":
        return _render_error(
            "Confirma que los valores de las dimensiones pueden mostrarse en pantalla."
        )
    raw_columns = request.form.getlist("measure_column")
    raw_operations = request.form.getlist("measure_operation")
    if len(raw_columns) != len(raw_operations):
        return _render_error("La configuración de medidas está incompleta.")
    measures = [
        (column, operation)
        for column, operation in zip(raw_columns, raw_operations)
        if operation and (operation == "count_rows" or column)
    ]
    try:
        columns, rows = _calculate_configured_statistics(
            result.df, dimensions, measures
        )
    except (ValueError, TypeError, KeyError) as exc:
        return _render_error(str(exc))
    workflow["configured_statistics_columns"] = columns
    workflow["configured_statistics_rows"] = rows
    workflow["configured_statistics_config"] = {
        "dimensions": dimensions,
        "measures": [(column, operation) for column, operation in measures],
    }
    _audit.write(
        g.workflow_id,
        "process_finished",
        "ok",
        {
            "event_type": "configured_statistics",
            "dimension_columns": dimensions,
            "measure_operations": [operation for _, operation in measures],
            "input_rows": len(result.df),
            "group_count": len(rows),
        },
        input_hash=workflow["input_hash"],
    )
    return redirect(url_for("result_stats"))


@app.post("/stats/export")
def export_configured_statistics():
    workflow = _current_workflow()
    rows = workflow.get("configured_statistics_rows") if workflow else None
    columns = workflow.get("configured_statistics_columns") if workflow else None
    if rows is None or columns is None:
        return _render_error("Primero configura las estadísticas.", 410)
    dataframe = pd.DataFrame(rows, columns=columns)
    output = io.BytesIO()
    try:
        XlsxExporter()._sanitize_dataframe(dataframe).to_excel(
            output, index=False, sheet_name="Estadisticas", engine="openpyxl"
        )
    except (OSError, ValueError, TypeError) as exc:
        _audit.write(
            g.workflow_id,
            "export",
            "error",
            {"type": type(exc).__name__, "kind": "statistics"},
            input_hash=workflow["input_hash"],
        )
        return _render_error("No se pudo preparar el archivo de estadísticas.", 500)
    output.seek(0)
    _audit.write(
        g.workflow_id,
        "export",
        "ok",
        {"rows": len(rows), "columns": len(columns), "kind": "statistics"},
        input_hash=workflow["input_hash"],
    )
    return send_file(
        output,
        as_attachment=True,
        download_name="cruzador_estadisticas.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/recipes/save")
def save_current_recipe():
    workflow = _current_workflow()
    result = workflow.get("result") if workflow else None
    cross_config = workflow.get("cross_config") if workflow else None
    if (
        not result
        or result.df is None
        or not cross_config
        or workflow.get("is_extension")
    ):
        return _render_error(
            "Ejecuta un cruce nuevo para guardar su configuración; "
            "los cruces encadenados no se pueden guardar como receta."
        )
    name = request.form.get("name", "").strip()
    description = request.form.get("description", "").strip()
    if not name or len(name) > 100 or len(description) > 500:
        return _render_error("El nombre es obligatorio (máximo 100 caracteres) y la descripción admite hasta 500.")

    entries: list[FileEntry] = workflow["source_entries"]
    steps = [
        CrossStepSpec(
            left_label=left.label,
            right_label=right.label,
            key_pairs=step_pairs,
        )
        for left, right, step_pairs in zip(
            entries,
            entries[1:],
            cross_config["steps_key_pairs"],
        )
    ]
    recipe = build_recipe_from_state(
        name=name,
        description=description,
        file_entries=entries,
        cleansing_rules=workflow["cleansing_rules"],
        join_type=cross_config["join_type"],
        normalize_keys=cross_config["normalize_keys"],
        cross_steps=steps,
        business_rules=workflow.get("business_rules", []),
    )
    try:
        save_recipe(recipe, _recipes_dir() / f"{recipe.recipe_id}.json")
    except OSError as exc:
        _audit.write(
            g.workflow_id,
            "error",
            "error",
            {"type": type(exc).__name__, "message": "recipe persistence failed"},
            input_hash=workflow["input_hash"],
        )
        return _render_error("No se pudo guardar la configuración en el servidor.", 500)
    _audit.write(
        g.workflow_id,
        "process_finished",
        "ok",
        {
            "event_type": "recipe_saved",
            "file_count": len(entries),
            "cleansing_rule_count": len(workflow["cleansing_rules"]),
            "business_rule_count": len(workflow.get("business_rules", [])),
        },
        input_hash=workflow["input_hash"],
    )
    flash("Configuración guardada en el servidor.")
    return redirect(url_for("show_result"))


@app.post("/export")
def export_result():
    workflow = _current_workflow()
    result = workflow.get("result") if workflow else None
    if not result or result.df is None:
        return _render_error("No hay un resultado disponible para exportar.", 410)

    sanitized = XlsxExporter()._sanitize_dataframe(result.df)
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            sanitized.to_excel(writer, sheet_name="Resultado", index=False)
    except (OSError, ValueError) as exc:
        _audit.write(
            g.workflow_id,
            "export",
            "error",
            {"type": type(exc).__name__},
            input_hash=workflow["input_hash"],
        )
        return _render_error("No se pudo preparar el archivo de salida.", 500)

    output.seek(0)
    _audit.write(
        g.workflow_id,
        "export",
        "ok",
        {"rows": result.row_count, "columns": len(result.df.columns)},
        input_hash=workflow["input_hash"],
    )
    return send_file(
        output,
        as_attachment=True,
        download_name="cruzador_resultado.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/reset")
def reset_workflow():
    with _workflow_lock:
        _workflows.pop(g.workflow_id, None)
    session.pop("workflow_id", None)
    flash("Se limpio la sesion de trabajo.")
    return redirect(url_for("index"))


@app.errorhandler(Exception)
def handle_request_error(error):
    if isinstance(error, HTTPException):
        messages = {
            400: "La solicitud no pudo validarse. Actualizá la página e intentá nuevamente.",
            403: "Esta conexión no está autorizada para usar el servicio.",
            404: "La página o recurso solicitado no existe.",
            405: "La operación solicitada no está disponible desde este enlace.",
            413: "La carga supera el límite permitido. Reducí el tamaño de los archivos.",
            503: "El servicio no puede aceptar otra solicitud en este momento.",
        }
        status = error.code or 500
        return _render_error(
            messages.get(status, "Solicitud no válida."),
            status,
            technical_type=f"http_{status}",
        )
    app.logger.exception("Unhandled application error")
    return _render_error(
        "Ocurrió un error interno. No se pudo completar la solicitud. Intentá nuevamente o contactá a soporte con la referencia mostrada.",
        500,
        technical_type=type(error).__name__,
    )


if __name__ == "__main__":
    app.run(host=BIND_HOST, port=8000, debug=False)
