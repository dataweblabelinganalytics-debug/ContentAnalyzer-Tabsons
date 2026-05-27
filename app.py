"""
Content Analyzer Flask backend with MongoDB Atlas storage.

Serves the dashboard frontend and provides APIs for file processing,
sheet retrieval, brand management, and report downloads.
"""

import io
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
import zipfile
from datetime import datetime
from functools import wraps
from pathlib import Path
from barc_nct_comparison import run_comparison as engine_run_comparison
import pandas as pd
from bson.decimal128 import Decimal128
from bson.binary import Binary
from bson.objectid import ObjectId
from dotenv import load_dotenv
from flask import Flask, jsonify, request, send_file, send_from_directory
from flask_cors import CORS
from werkzeug.exceptions import HTTPException
from pymongo import ASCENDING, MongoClient
from pymongo.errors import InvalidURI, PyMongoError

load_dotenv()

MONGO_URI = os.getenv("MONGO_URI")
DB_NAME = os.getenv("DB_NAME") or os.getenv("MONGO_DB_NAME", "content_analyzer")
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_FILE = BASE_DIR / "brand_comparison_template.xlsx"
COMPARISON_SCRIPT = BASE_DIR / "barc_nct_comparison.py"

app = Flask(__name__, static_folder=".", static_url_path="")
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "dev-secret-key-change-me")
CORS(app)

logging.basicConfig(
    level=getattr(logging, LOG_LEVEL, logging.INFO),
    format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
)
app.logger.setLevel(getattr(logging, LOG_LEVEL, logging.INFO))

_mongo_client = None
_mongo_db = None
_indexes_ready = False


def is_api_request() -> bool:
    return request.path.startswith("/api/") or request.path in {"/analyze", "/healthz"}


def to_json_safe(value):
    """Recursively convert MongoDB/BSON and pandas values into JSON-safe data."""
    if isinstance(value, ObjectId):
        return str(value)
    if isinstance(value, Decimal128):
        return str(value)
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, Binary):
        return f"<binary {len(value)} bytes>"
    if isinstance(value, (bytes, bytearray, memoryview)):
        return f"<binary {len(value)} bytes>"
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, dict):
        return {str(to_json_safe(key)): to_json_safe(val) for key, val in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [to_json_safe(item) for item in value]

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if hasattr(value, "item"):
        try:
            return value.item()
        except (TypeError, ValueError):
            pass

    return value


def validate_json_payload(payload, route_name: str = "") -> None:
    """Fail fast if an API payload still contains non-serializable objects."""
    try:
        json.dumps(payload)
    except TypeError as exc:
        app.logger.exception("JSON validation failed for %s: %s", route_name or request.path, exc)
        raise


def json_response(payload, status: int = 200):
    safe_payload = to_json_safe(payload)
    validate_json_payload(safe_payload, request.endpoint or request.path)
    response = jsonify(safe_payload)
    response.status_code = status
    return response


def api_error(message: str, status: int = 500, **extra):
    payload = {"success": False, "error": str(message)}
    payload.update(extra)
    return json_response(payload, status)


def utc_now_iso() -> str:
    return datetime.utcnow().isoformat()


def get_mongo_client() -> MongoClient:
    """Return the shared MongoDB client with connection pooling enabled."""
    global _mongo_client
    if _mongo_client is None:
        if not MONGO_URI:
            raise RuntimeError("MONGO_URI environment variable is required")

        try:
            _mongo_client = MongoClient(
                MONGO_URI,
                appname="content-analyzer",
                maxPoolSize=int(os.getenv("MONGO_MAX_POOL_SIZE", "50")),
                minPoolSize=int(os.getenv("MONGO_MIN_POOL_SIZE", "0")),
                serverSelectionTimeoutMS=int(os.getenv("MONGO_SERVER_SELECTION_TIMEOUT_MS", "5000")),
                connectTimeoutMS=int(os.getenv("MONGO_CONNECT_TIMEOUT_MS", "10000")),
                socketTimeoutMS=int(os.getenv("MONGO_SOCKET_TIMEOUT_MS", "20000")),
                retryWrites=True,
            )
        except InvalidURI as exc:
            raise RuntimeError(
                "Invalid MONGO_URI. URL-encode the MongoDB username/password "
                "in the Atlas URI, for example @ as %40, # as %23, / as %2F."
            ) from exc
    return _mongo_client


def get_mongo_database():
    """Return the configured MongoDB database."""
    global _mongo_db
    if _mongo_db is None:
        _mongo_db = get_mongo_client()[DB_NAME]
    return _mongo_db


def validate_mongo_connection() -> None:
    """Validate MongoDB Atlas connectivity with a ping command."""
    app.logger.debug("Validating MongoDB connection for database '%s'", DB_NAME)
    get_mongo_client().admin.command("ping")
    app.logger.info("MongoDB connection validated for database '%s'", DB_NAME)


def ensure_mongo_indexes() -> None:
    """Create indexes required by the dashboard lookup patterns."""
    global _indexes_ready
    if _indexes_ready:
        return

    validate_mongo_connection()
    database = get_mongo_database()
    database.processed_files.create_index(
        [("channel_name", ASCENDING), ("date", ASCENDING)],
        name="idx_pf_lookup",
        unique=True,
    )
    database.processed_files.create_index(
        [("file_id", ASCENDING)],
        name="idx_pf_file_id",
        unique=True,
    )
    database.sheets.create_index(
        [("channel_name", ASCENDING), ("date", ASCENDING), ("sheet_name", ASCENDING)],
        name="idx_sheets_lookup",
        unique=True,
    )
    database.sheets.create_index(
        [("file_id", ASCENDING)],
        name="idx_sheets_file_id",
    )
    database.brand_modifications.create_index(
        [("channel_name", ASCENDING), ("date", ASCENDING), ("timestamp", ASCENDING)],
        name="idx_bm_lookup",
    )
    _indexes_ready = True
    app.logger.info("MongoDB indexes are ready")


def get_collections():
    ensure_mongo_indexes()
    database = get_mongo_database()
    return (
        database.processed_files,
        database.sheets,
        database.brand_modifications,
    )


def database_error_response(exc):
    app.logger.exception("Database operation failed: %s", exc)
    return api_error("Database operation failed", 500, details=str(exc))


@app.before_request
def log_request_start():
    request._started_at = time.time()
    if is_api_request():
        file_names = [storage.filename for storage in request.files.values()]
        json_keys = []
        if request.is_json:
            body = request.get_json(silent=True) or {}
            json_keys = list(body.keys()) if isinstance(body, dict) else []
        app.logger.info(
            "API request start method=%s path=%s args=%s json_keys=%s files=%s",
            request.method,
            request.path,
            dict(request.args),
            json_keys,
            file_names,
        )


@app.after_request
def log_request_finish(response):
    if is_api_request():
        elapsed_ms = int((time.time() - getattr(request, "_started_at", time.time())) * 1000)
        app.logger.info(
            "API request finish method=%s path=%s status=%s mimetype=%s elapsed_ms=%s",
            request.method,
            request.path,
            response.status_code,
            response.mimetype,
            elapsed_ms,
        )
    return response


@app.errorhandler(Exception)
def handle_exception(e):
    print("GLOBAL ERROR:", str(e))
    status = e.code if isinstance(e, HTTPException) else 500
    if status >= 500:
        app.logger.exception("Unhandled exception on %s %s: %s", request.method, request.path, e)
    else:
        app.logger.warning("HTTP exception on %s %s: %s", request.method, request.path, e)
    return json_response({"success": False, "error": str(e)}, status)


def route_guard(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            app.logger.debug("Entering route endpoint=%s path=%s", request.endpoint, request.path)
            return func(*args, **kwargs)
        except Exception as exc:
            return handle_exception(exc)

    wrapper._route_guarded = True
    return wrapper


def fmt_size(b: int) -> str:
    if b < 1024:
        return f"{b} B"
    if b < 1024**2:
        return f"{b / 1024:.1f} KB"
    return f"{b / 1024**2:.1f} MB"


def build_filename(channel: str, date_str: str) -> str:
    channel_clean = re.sub(r"[^A-Z0-9]", "", str(channel).upper().strip())
    if not channel_clean:
        channel_clean = "UNKNOWN"

    match = re.match(r"^(\d{2})/(\d{2})/(\d{4})$", str(date_str).strip())
    if match:
        date_clean = f"{match.group(1)}{match.group(2)}{match.group(3)}"
    else:
        iso_match = re.match(r"^(\d{4})-(\d{2})-(\d{2})", str(date_str).strip())
        date_clean = (
            f"{iso_match.group(3)}{iso_match.group(2)}{iso_match.group(1)}"
            if iso_match
            else "00000000"
        )

    return f"{channel_clean}({date_clean}) barc_nct_comparison"


def run_comparison(file_bytes: bytes, original_name: str) -> tuple:
    """Run barc_nct_comparison.py and return (xlsx_bytes, output_filename, stats)."""
    output_filename = "output.xlsx"
    stats = {}

    try:
        df = pd.read_excel(io.BytesIO(file_bytes), dtype=str)
        source_series = df.get("source", pd.Series([""] * len(df)))
        normalized_source = source_series.fillna("").astype(str).str.upper().str.strip()
        df_barc = df[normalized_source == "BARC XML"]

        if len(df_barc):
            channel = str(df_barc["channel name"].iloc[0])
            date_val = str(df_barc["TelecastDate"].iloc[0])
            output_filename = build_filename(channel, date_val) + ".xlsx"
            stats["channel"] = channel
            stats["date"] = date_val
            stats["barc_rows"] = int(len(df_barc))
            stats["nct_rows"] = int((normalized_source == "NCT").sum())
    except Exception as exc:
        stats["metadata_error"] = str(exc)

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = Path(tmpdir) / "brand_comparison_template.xlsx"
        input_path.write_bytes(file_bytes)

        temp_script = Path(tmpdir) / "barc_nct_comparison.py"
        shutil.copy(COMPARISON_SCRIPT, temp_script)
        child_env = os.environ.copy()
        child_env["PYTHONIOENCODING"] = "utf-8"
        subprocess.run([sys.executable, str(temp_script)], cwd=tmpdir, env=child_env, check=True)

        output_path = Path(tmpdir) / "barc_nct_comparison.xlsx"
        if not output_path.exists():
            raise FileNotFoundError("barc_nct_comparison.xlsx was not generated.")

        xlsx_bytes = output_path.read_bytes()

    return xlsx_bytes, output_filename, stats


def parse_workbook_sheets(xlsx_bytes: bytes, file_id: str, channel_name: str, date_str: str):
    uploaded_at = utc_now_iso()
    sheet_documents = []
    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str, header=None)
        rows_data = []
        for _, row in df.iterrows():
            rows_data.append([str(value) if pd.notna(value) else "" for value in row])

        headers = rows_data[0] if rows_data else []
        sheet_documents.append(
            {
                "file_id": file_id,
                "channel_name": channel_name,
                "date": date_str,
                "sheet_name": sheet_name,
                "headers": headers,
                "rows": rows_data,
                "row_count": len(rows_data),
                "col_count": len(headers),
                "uploaded_at": uploaded_at,
            }
        )

    return sheet_documents


def upload_to_db(xlsx_bytes: bytes, channel_name: str, date_str: str, original_filename: str):
    """Parse the processed Excel workbook and store it in MongoDB."""
    processed_files, sheets, _ = get_collections()
    file_id = str(uuid.uuid4())
    uploaded_at = utc_now_iso()
    sheet_documents = parse_workbook_sheets(xlsx_bytes, file_id, channel_name, date_str)

    file_document = {
        "file_id": file_id,
        "channel_name": channel_name,
        "date": date_str,
        "original_filename": original_filename,
        "xlsx_data": Binary(xlsx_bytes),
        "uploaded_at": uploaded_at,
    }

    lookup = {"channel_name": channel_name, "date": date_str}
    sheets.delete_many(lookup)
    processed_files.replace_one(lookup, file_document, upsert=True)
    if sheet_documents:
        sheets.insert_many(sheet_documents)

    print(f"[MongoDB] Stored {len(sheet_documents)} sheets for {channel_name} / {date_str}")
    return file_id


def clean_num(val):
    if not val or val == "None" or val == "nan":
        return 0
    text = str(val).replace(",", "").strip()
    try:
        return int(float(text))
    except (ValueError, TypeError):
        return 0


def parse_count_cell(value) -> int:
    text = str(value or "").strip()
    if text in {"", "-", "\u2013", "\u2014", "\u00e2\u20ac\u201d"}:
        return 0
    match = re.search(r"-?\d+(?:\.\d+)?", text.replace(",", ""))
    if not match:
        return 0
    try:
        return int(float(match.group(0)))
    except (TypeError, ValueError):
        return 0


def workbook_bytes_from_document(document) -> bytes:
    return bytes(document.get("xlsx_data") or b"")


@app.route("/")
def index():
    return send_from_directory(str(BASE_DIR), "content_analyzer.html")


@app.route("/healthz")
def healthz():
    return json_response({"status": "ok"})


@app.route("/api/channels-dates", methods=["GET"])
def get_channels_dates():
    """Return distinct channel/date combinations from MongoDB."""
    try:
        processed_files, _, _ = get_collections()
        docs = processed_files.find(
            {},
            {"_id": 0, "channel_name": 1, "date": 1},
        ).sort([("channel_name", ASCENDING), ("date", ASCENDING)])
        return json_response(
            [
                {"channel_name": doc.get("channel_name", ""), "date": doc.get("date", "")}
                for doc in docs
            ]
        )
    except (PyMongoError, RuntimeError) as exc:
        return database_error_response(exc)


@app.route("/api/dashboard", methods=["GET"])
def get_dashboard():
    """Return KPI data for the dashboard."""
    channel = request.args.get("channel", "")
    date = request.args.get("date", "")
    source = request.args.get("source", "TABSONS-BARC")
    data_type = request.args.get("data_type", "COUNT")

    if not channel or not date:
        return api_error("channel and date required", 400)

    try:
        _, sheets, _ = get_collections()
        row = sheets.find_one(
            {"channel_name": channel, "date": date, "sheet_name": "TABSONS SUMMARY"},
            {"_id": 0, "headers": 1, "rows": 1},
        )
    except (PyMongoError, RuntimeError) as exc:
        return database_error_response(exc)

    if not row:
        return api_error("No data found for this channel/date", 404)

    all_rows = row.get("rows") or []
    headers = all_rows[1] if len(all_rows) > 1 else row.get("headers", [])
    data_row = all_rows[2] if len(all_rows) > 2 else []

    def get_val(header_keyword, r=data_row, hdrs=headers):
        for i, header in enumerate(hdrs):
            if header_keyword.upper() in str(header).upper() and i < len(r):
                return r[i]
        return "0"

    result = {"source": source, "data_type": data_type}

    if source == "TABSONS":
        if data_type == "COUNT":
            result["total_line_item"] = clean_num(get_val("TABSONS LINE ITEM"))
            result["commercial"] = clean_num(get_val("TABSONS COMMERCIAL COUNT"))
            result["promo"] = clean_num(get_val("TABSONS PROMO COUNT"))
            result["promo_sponsor"] = clean_num(get_val("TABSONS PROMO SPONSOR COUNT"))
            result["program"] = clean_num(get_val("TABSONS PROGRAM COUNT"))
        else:
            result["total_line_item"] = get_val("TABSONS DURATION")
            result["commercial"] = get_val("TABSONS COMMERCIAL DURATION")
            result["promo"] = get_val("TABSONS PROMO DURATION")
            result["promo_sponsor"] = get_val("TABSONS PROMO SPONSOR COUNT DURATIO")
            result["program"] = get_val("TABSONS PROGRAM DURATION")

    elif source == "BARC XML":
        if data_type == "COUNT":
            result["total_line_item"] = clean_num(get_val("BARC LINE ITEM"))
            result["commercial"] = clean_num(get_val("BARC COMMERCIAL COUNT"))
            result["promo"] = clean_num(get_val("BARC PROMO COUNT"))
            result["promo_sponsor"] = clean_num(get_val("BARC PROMO SPONSOR COUNT"))
            result["program"] = clean_num(get_val("BARC PROGRAM COUNT"))
        else:
            result["total_line_item"] = get_val("BARC DURATION")
            result["commercial"] = get_val("BARC COMMERCIAL DURATION")
            result["promo"] = get_val("BARC PROMO DURATION")
            result["promo_sponsor"] = get_val("BARC PROMO SPONSOR COUNT DURATION")
            result["program"] = get_val("BARC PROGRAM DURATION")

    else:
        if data_type == "COUNT":
            result["tabsons_total"] = clean_num(get_val("TABSONS LINE ITEM"))
            result["tabsons_commercial"] = clean_num(get_val("TABSONS COMMERCIAL COUNT"))
            result["tabsons_promo"] = clean_num(get_val("TABSONS PROMO COUNT"))
            result["tabsons_promo_sponsor"] = clean_num(get_val("TABSONS PROMO SPONSOR COUNT"))
            result["tabsons_program"] = clean_num(get_val("TABSONS PROGRAM COUNT"))
            result["barc_total"] = clean_num(get_val("BARC LINE ITEM"))
            result["barc_commercial"] = clean_num(get_val("BARC COMMERCIAL COUNT"))
            result["barc_promo"] = clean_num(get_val("BARC PROMO COUNT"))
            result["barc_promo_sponsor"] = clean_num(get_val("BARC PROMO SPONSOR COUNT"))
            result["barc_program"] = clean_num(get_val("BARC PROGRAM COUNT"))
        else:
            result["tabsons_total"] = get_val("TABSONS DURATION")
            result["tabsons_commercial"] = get_val("TABSONS COMMERCIAL DURATION")
            result["tabsons_promo"] = get_val("TABSONS PROMO DURATION")
            result["tabsons_promo_sponsor"] = get_val("TABSONS PROMO SPONSOR COUNT DURATIO")
            result["tabsons_program"] = get_val("TABSONS PROGRAM DURATION")
            result["barc_total"] = get_val("BARC DURATION")
            result["barc_commercial"] = get_val("BARC COMMERCIAL DURATION")
            result["barc_promo"] = get_val("BARC PROMO DURATION")
            result["barc_promo_sponsor"] = get_val("BARC PROMO SPONSOR COUNT DURATION")
            result["barc_program"] = get_val("BARC PROGRAM DURATION")

    return json_response(result)


@app.route("/api/sheets", methods=["GET"])
def get_sheets():
    """Return available sheets for a channel/date."""
    channel = request.args.get("channel", "")
    date = request.args.get("date", "")

    if not channel or not date:
        return api_error("channel and date required", 400)

    try:
        _, sheets, _ = get_collections()
        docs = sheets.find(
            {"channel_name": channel, "date": date},
            {"_id": 0, "sheet_name": 1, "row_count": 1, "col_count": 1},
        ).sort([("_id", ASCENDING)])
        return json_response(
            [
                {
                    "sheet_name": doc.get("sheet_name", ""),
                    "row_count": doc.get("row_count", 0),
                    "col_count": doc.get("col_count", 0),
                }
                for doc in docs
            ]
        )
    except (PyMongoError, RuntimeError) as exc:
        return database_error_response(exc)


@app.route("/api/sheet-data", methods=["GET"])
def get_sheet_data():
    """Return full sheet data for a channel/date/sheet."""
    channel = request.args.get("channel", "")
    date = request.args.get("date", "")
    sheet = request.args.get("sheet", "")

    if not channel or not date or not sheet:
        return api_error("channel, date, and sheet required", 400)

    try:
        _, sheets, _ = get_collections()
        row = sheets.find_one(
            {"channel_name": channel, "date": date, "sheet_name": sheet},
            {
                "_id": 0,
                "file_id": 1,
                "channel_name": 1,
                "date": 1,
                "sheet_name": 1,
                "headers": 1,
                "rows": 1,
                "row_count": 1,
                "col_count": 1,
                "uploaded_at": 1,
            },
        )
    except (PyMongoError, RuntimeError) as exc:
        return database_error_response(exc)

    if not row:
        return api_error("Sheet not found", 404)

    return json_response(
        {
            "file_id": row.get("file_id", ""),
            "channel_name": row.get("channel_name", ""),
            "date": row.get("date", ""),
            "sheet_name": row.get("sheet_name", ""),
            "headers": row.get("headers", []),
            "rows": row.get("rows", []),
            "row_count": row.get("row_count", 0),
            "col_count": row.get("col_count", 0),
            "uploaded_at": row.get("uploaded_at", ""),
        }
    )


@app.route("/api/commercial-comparison", methods=["GET"])
def get_commercial_comparison():
    """Return matched and unmatched brand data from COMMERCIAL COMPARISION sheet."""
    channel = request.args.get("channel", "")
    date = request.args.get("date", "")

    if not channel or not date:
        return api_error("channel and date required", 400)

    try:
        _, sheets, brand_modifications = get_collections()
        row = sheets.find_one(
            {"channel_name": channel, "date": date, "sheet_name": "COMMERCIAL COMPARISION"},
            {"_id": 0, "rows": 1},
        )
        mods = list(
            brand_modifications.find(
                {"channel_name": channel, "date": date},
                {"_id": 0},
            ).sort([("timestamp", ASCENDING)])
        )
    except (PyMongoError, RuntimeError) as exc:
        return database_error_response(exc)

    if not row:
        return api_error("Commercial comparison data not found", 404)

    rows = row.get("rows") or []
    headers_row = None
    matched_rows = []
    unmatched_rows = []
    section = "none"

    for r in rows:
        first_cell = str(r[0]).strip() if r else ""

        if first_cell == "SOURCE" and headers_row is None:
            headers_row = r
            continue

        if "BARC COMMERCIAL vs NCT COMMERCIAL" in first_cell and "MATCHED" in first_cell:
            section = "matched"
            continue

        if "NCT COMMERCIAL BRANDS" in first_cell and "NOT MATCHED" in first_cell:
            section = "unmatched"
            continue

        if first_cell == "MATCHED" or (
            "MATCHED" in first_cell
            and "UNMATCHED" not in first_cell
            and "NCT COMMERCIAL BRANDS" not in first_cell
            and "BARC COMMERCIAL" not in first_cell
            and first_cell not in ("MATCHING COMMERCIAL TOTAL",)
        ):
            section = "matched"
            continue

        if "NOT MATCHED" in first_cell or "UNMATCHED" in first_cell:
            section = "unmatched"
            continue

        if first_cell in (
            "",
            "MATCHING COMMERCIAL TOTAL",
            "NCT UNMATCHED TOTAL",
            "GRAND TOTAL",
        ) or "COMMERCIAL COMPARISION" in first_cell:
            continue

        if headers_row and len(r) >= 5:
            row_dict = {}
            header_labels = [
                "source",
                "channel_name",
                "date",
                "barc_brand",
                "nct_brand",
                "barc_count",
                "nct_count",
                "barc_duration",
                "nct_duration",
                "nct_ps_count",
                "nct_ps_duration",
                "remarks",
            ]
            for j, label in enumerate(header_labels):
                if j < len(r):
                    row_dict[label] = str(r[j]) if r[j] is not None else ""

            if section == "matched" and first_cell in ("BARC XML", "NCT"):
                matched_rows.append(row_dict)
            elif section == "unmatched" and first_cell in ("NCT", "BARC XML"):
                unmatched_rows.append(row_dict)

    for mod in mods:
        action = mod.get("action")
        brand_name = mod.get("brand_name", "")

        if action == "remove_from_matched":
            removed = None
            for j, row_data in enumerate(matched_rows):
                if row_data.get("barc_brand", "").strip() == brand_name.strip():
                    removed = matched_rows.pop(j)
                    break

            if not removed:
                for j, row_data in enumerate(matched_rows):
                    if row_data.get("nct_brand", "").strip() == brand_name.strip():
                        removed = matched_rows.pop(j)
                        break

            if removed:
                removed["remarks"] = "REMOVED FROM MATCHED"
                unmatched_rows.append(removed)

        elif action == "merge_to_matched":
            target_barc = mod.get("target_barc_brand") or ""
            merged = None
            for j, row_data in enumerate(unmatched_rows):
                if row_data.get("nct_brand", "").strip() == brand_name.strip():
                    merged = unmatched_rows.pop(j)
                    break

            if merged and target_barc:
                for row_data in matched_rows:
                    if row_data.get("barc_brand", "").strip() == target_barc.strip():
                        existing_count = parse_count_cell(row_data.get("nct_count", "0"))
                        merge_count = parse_count_cell(merged.get("nct_count", "0"))
                        row_data["nct_count"] = str(existing_count + merge_count)
                        row_data["remarks"] = "MATCHED (MERGED)"
                        break

    return json_response(
        {
            "matched": matched_rows,
            "unmatched": unmatched_rows,
            "headers": headers_row or [],
        }
    )


@app.route("/api/commercial/move-brand", methods=["POST"])
def move_brand():
    """Move a brand between matched and unmatched tables."""
    data = request.get_json(silent=True) or {}
    channel = data.get("channel", "")
    date = data.get("date", "")
    action = data.get("action", "")
    brand_name = data.get("brand_name", "")
    target_barc_brand = data.get("target_barc_brand", "")

    if not channel or not date or not action or not brand_name:
        return api_error("Missing required fields", 400)

    try:
        _, _, brand_modifications = get_collections()
        brand_modifications.insert_one(
            {
                "channel_name": channel,
                "date": date,
                "action": action,
                "brand_name": brand_name,
                "target_barc_brand": target_barc_brand,
                "timestamp": utc_now_iso(),
            }
        )
    except (PyMongoError, RuntimeError) as exc:
        return database_error_response(exc)

    return json_response({"success": True, "message": f"Brand '{brand_name}' {action} successfully"})


@app.route("/api/commercial/undo-modifications", methods=["POST"])
def undo_modifications():
    """Clear all brand modifications for a channel/date."""
    data = request.get_json(silent=True) or {}
    channel = data.get("channel", "")
    date = data.get("date", "")

    try:
        _, _, brand_modifications = get_collections()
        brand_modifications.delete_many({"channel_name": channel, "date": date})
    except (PyMongoError, RuntimeError) as exc:
        return database_error_response(exc)

    return json_response({"success": True})


@app.route("/api/download/preprocessed", methods=["GET"])
def download_preprocessed():
    """Download the original processed Excel file."""
    channel = request.args.get("channel", "")
    date = request.args.get("date", "")

    try:
        processed_files, _, _ = get_collections()
        row = processed_files.find_one(
            {"channel_name": channel, "date": date},
            {"_id": 0, "xlsx_data": 1, "original_filename": 1},
        )
    except (PyMongoError, RuntimeError) as exc:
        return database_error_response(exc)

    if not row:
        return api_error("File not found", 404)

    return send_file(
        io.BytesIO(workbook_bytes_from_document(row)),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=row.get("original_filename") or "report.xlsx",
    )


@app.route("/api/download/updated", methods=["GET"])
def download_updated():
    """Download the updated Excel report with brand modifications applied."""
    channel = request.args.get("channel", "")
    date = request.args.get("date", "")

    try:
        processed_files, _, brand_modifications = get_collections()
        row = processed_files.find_one(
            {"channel_name": channel, "date": date},
            {"_id": 0, "xlsx_data": 1, "original_filename": 1},
        )
        mods = list(
            brand_modifications.find(
                {"channel_name": channel, "date": date},
                {"_id": 0},
            ).sort([("timestamp", ASCENDING)])
        )
    except (PyMongoError, RuntimeError) as exc:
        return database_error_response(exc)

    if not row:
        return api_error("File not found", 404)

    xlsx_data = workbook_bytes_from_document(row)

    if mods:
        import openpyxl

        workbook = openpyxl.load_workbook(io.BytesIO(xlsx_data))

        if "COMMERCIAL COMPARISION" in workbook.sheetnames:
            worksheet = workbook["COMMERCIAL COMPARISION"]
            last_row = worksheet.max_row + 2
            worksheet.cell(row=last_row, column=1, value="MODIFICATIONS APPLIED:")
            for i, mod in enumerate(mods):
                worksheet.cell(
                    row=last_row + i + 1,
                    column=1,
                    value=f"{mod.get('action')}: {mod.get('brand_name')}",
                )
                target_barc_brand = mod.get("target_barc_brand")
                if target_barc_brand:
                    worksheet.cell(
                        row=last_row + i + 1,
                        column=2,
                        value=f"\u2192 {target_barc_brand}",
                    )

        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        xlsx_data = output.read()

    filename = row.get("original_filename") or "report.xlsx"
    if mods:
        base, ext = os.path.splitext(filename)
        filename = f"{base}_UPDATED{ext}"

    return send_file(
        io.BytesIO(xlsx_data),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


@app.route("/analyze", methods=["POST"])
def analyze():
    """Process uploaded file(s), store generated workbook data, and return downloads."""
    files = request.files.getlist("files")
    if not files:
        return api_error("No files provided", 400)

    results = []
    errors = []

    for uploaded_file in files:
        try:
            file_bytes = uploaded_file.read()
            xlsx_bytes, output_filename, stats = run_comparison(file_bytes, uploaded_file.filename)

            channel = stats.get("channel", "UNKNOWN")
            date = stats.get("date", "00/00/0000")
            try:
                file_id = upload_to_db(xlsx_bytes, channel, date, output_filename)
                stats["uploaded_to_db"] = True
                stats["file_id"] = file_id
            except Exception as db_err:
                stats["uploaded_to_db"] = False
                stats["db_error"] = str(db_err)
                app.logger.exception("MongoDB upload failed: %s", db_err)

            results.append({"fname": output_filename, "data": xlsx_bytes, "stats": stats})
        except Exception as exc:
            errors.append({"file": uploaded_file.filename, "error": str(exc)})

    if len(files) == 1 and len(results) == 1:
        result = results[0]
        return send_file(
            io.BytesIO(result["data"]),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=result["fname"],
        )

    if results:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for result in results:
                zip_file.writestr(result["fname"], result["data"])
        zip_buf.seek(0)
        zip_name = f"barc_nct_batch_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        return send_file(
            zip_buf,
            mimetype="application/zip",
            as_attachment=True,
            download_name=zip_name,
        )

    return api_error("All files failed", 500, details=errors)


@app.route("/api/compare", methods=["POST"])
def compare_report():
    """Upload a file, run comparison, store it, and return the result."""
    uploaded_file = request.files.get("file")
    if not uploaded_file:
        return api_error("No file provided", 400)

    try:
        file_bytes = uploaded_file.read()
        xlsx_bytes, output_filename, stats = run_comparison(file_bytes, uploaded_file.filename)

        channel = stats.get("channel", "UNKNOWN")
        date = stats.get("date", "00/00/0000")
        try:
            upload_to_db(xlsx_bytes, channel, date, output_filename)
        except Exception as db_err:
            app.logger.exception("MongoDB upload failed: %s", db_err)

        return send_file(
            io.BytesIO(xlsx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=output_filename,
        )
    except Exception as exc:
        return api_error(str(exc), 500)


@app.route("/api/template", methods=["GET"])
def download_template():
    """Download the brand comparison template file bundled with the project."""
    if not TEMPLATE_FILE.exists():
        return api_error("Template file not found", 404)

    return send_file(
        TEMPLATE_FILE,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="brand_comparison_template.xlsx",
    )


def install_route_guards() -> None:
    """Wrap every Flask view in a route-level try/except without changing routes."""
    for endpoint, view_func in list(app.view_functions.items()):
        if endpoint == "static" or getattr(view_func, "_route_guarded", False):
            continue
        app.view_functions[endpoint] = route_guard(view_func)
    app.logger.debug("Route guards installed for Flask views")


def run_startup_diagnostics() -> None:
    """Log production startup details and validate MongoDB when configured."""
    app.logger.info(
        "Startup diagnostics flask_env=%s db_name=%s mongo_uri_set=%s template_exists=%s processor_exists=%s",
        os.getenv("FLASK_ENV", ""),
        DB_NAME,
        bool(MONGO_URI),
        TEMPLATE_FILE.exists(),
        COMPARISON_SCRIPT.exists(),
    )

    if os.getenv("MONGO_DB_NAME") and not os.getenv("DB_NAME"):
        app.logger.warning("MONGO_DB_NAME is deprecated. Set DB_NAME on Render and locally.")

    if not MONGO_URI:
        app.logger.warning("MONGO_URI is not set. Database routes require MongoDB Atlas.")
        return

    try:
        ensure_mongo_indexes()
    except Exception as exc:
        app.logger.warning("MongoDB startup validation/index initialization deferred: %s", exc)


install_route_guards()
run_startup_diagnostics()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_ENV", "").lower() == "development"
    app.run(host="0.0.0.0", port=port, debug=debug)
