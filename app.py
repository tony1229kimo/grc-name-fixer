"""
Flask 網頁主程式
- / : 處理檔案主頁
- /dictionary : 字典管理頁
- /api/process : 上傳兩個檔案並處理
- /api/download/<job_id> : 下載處理結果
- /api/dictionary : 字典 CRUD
"""
from __future__ import annotations

import json
import os
import re
import time
import urllib.error
import urllib.request
import uuid
from datetime import datetime
from pathlib import Path

from flask import (
    Flask, render_template, request, jsonify, send_file, abort
)
from werkzeug.utils import secure_filename

import database
import matcher
from parser import parse_calendar_report, parse_flat_reference
from _version import VERSION, RELEASES_API_URL, RELEASES_PAGE_URL

# ---------- 基本設定 ----------

MODULE_DIR = Path(__file__).parent

# 資源目錄（templates, static）→ 打包時由 PyInstaller 解壓到 _MEIPASS
RESOURCE_DIR = Path(os.environ.get("GRC_RESOURCE_DIR") or str(MODULE_DIR))
# 資料目錄（uploads, outputs, dictionary.db）→ 可寫入位置
DATA_DIR = Path(os.environ.get("GRC_DATA_DIR") or str(MODULE_DIR))

UPLOAD_DIR = DATA_DIR / "uploads"
OUTPUT_DIR = DATA_DIR / "outputs"
LOG_DIR = DATA_DIR / "logs"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)

JOB_ID_RE = re.compile(r"^[a-f0-9]{32}$")


def _log_error(prefix: str, err_id: str, detail: str) -> None:
    """把錯誤詳細寫到 logs/error.log，給同事回報用"""
    log_file = LOG_DIR / "error.log"
    try:
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"\n=== [{datetime.now():%Y-%m-%d %H:%M:%S}] {prefix} 代號 {err_id} ===\n")
            f.write(detail)
            f.write("\n")
    except OSError:
        pass

MAX_FILE_SIZE = 20 * 1024 * 1024  # 20 MB
ALLOWED_EXT = {".xlsx", ".xlsm"}

app = Flask(__name__,
            template_folder=str(RESOURCE_DIR / "templates"),
            static_folder=str(RESOURCE_DIR / "static"))
app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_SIZE


# ---------- 工具 ----------

def _is_allowed(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXT


def _save_upload(file_storage, prefix: str) -> Path:
    fname = secure_filename(file_storage.filename or f"{prefix}.xlsx")
    unique = f"{prefix}_{uuid.uuid4().hex[:8]}_{fname}"
    dst = UPLOAD_DIR / unique
    file_storage.save(str(dst))
    return dst


def _cleanup_old_files(days: int = 3):
    """清理 3 天前的上傳 / 輸出檔（含 .name sidecar）"""
    from time import time
    now = time()
    cutoff = now - days * 86400
    for folder in (UPLOAD_DIR, OUTPUT_DIR):
        for f in folder.iterdir():
            if f.is_file() and f.stat().st_mtime < cutoff:
                try:
                    f.unlink()
                except OSError:
                    pass


# ---------- 頁面 ----------

@app.route("/")
def index():
    dict_count = database.count()
    return render_template("index.html", dict_count=dict_count)


@app.route("/dictionary")
def dictionary_page():
    return render_template("dictionary.html")


# ---------- API：處理檔案 ----------

@app.route("/api/process", methods=["POST"])
def api_process():
    _cleanup_old_files()

    if "calendar" not in request.files:
        return jsonify({"ok": False, "error": "請上傳 GRC 縮寫檔（必填）"}), 400

    calendar_file = request.files["calendar"]
    if not calendar_file.filename:
        return jsonify({"ok": False, "error": "請選擇 GRC 縮寫檔"}), 400
    if not _is_allowed(calendar_file.filename):
        return jsonify({"ok": False, "error": "檔案格式不支援，請使用 .xlsx"}), 400

    reference_file = request.files.get("reference")
    if reference_file and reference_file.filename:
        if not _is_allowed(reference_file.filename):
            return jsonify(
                {"ok": False, "error": "參考檔格式不支援，請使用 .xlsx"}
            ), 400

    # 儲存上傳檔
    calendar_path = _save_upload(calendar_file, "calendar")
    reference_path = None
    if reference_file and reference_file.filename:
        reference_path = _save_upload(reference_file, "reference")

    # 產出檔在硬碟上一律用純 ASCII UUID 命名（避免中文 / 怪字元造成下載失敗）
    # 「給使用者看的友善檔名」另存 sidecar，由 download 路由設給瀏覽器
    job_id = uuid.uuid4().hex
    output_path = OUTPUT_DIR / f"{job_id}.xlsx"

    original_name = Path(calendar_file.filename).stem
    download_name = f"{original_name}_已修正_{datetime.now():%Y%m%d_%H%M%S}.xlsx"

    try:
        summary = matcher.process(
            calendar_file=calendar_path,
            reference_file=reference_path,
            output_file=output_path,
        )
    except ValueError as e:
        # parser 拋出的已知錯誤（欄位找不到、格式不對等）
        print(f"[處理錯誤-格式] {e}")
        return jsonify({
            "ok": False,
            "error": f"檔案格式問題：{e}。請確認上傳的檔案是否正確（縮寫檔是 GRC 日曆格式，參考檔是扁平格式）。"
        }), 400
    except Exception as e:
        import traceback
        err_id = uuid.uuid4().hex[:8]
        err_detail = traceback.format_exc()
        print(f"[處理錯誤-未知 代號 {err_id}]\n{err_detail}")
        _log_error("處理錯誤-未知", err_id, err_detail)
        return jsonify({
            "ok": False,
            "error": (
                f"處理時發生未預期錯誤（代號 {err_id}）。"
                f"請把 logs/error.log 傳給 IT 協助排查。"
            ),
        }), 500

    # 寫 sidecar 記錄友善檔名（給 download 路由用）
    try:
        (OUTPUT_DIR / f"{job_id}.name").write_text(download_name, encoding="utf-8")
    except OSError:
        pass  # 寫不進去就算了，前端還是會用 download_name 設 a[download] attr

    return jsonify({
        "ok": True,
        "job_id": job_id,
        "summary": {
            "total": summary.total,
            "from_dictionary": summary.from_dictionary,
            "from_reference": summary.from_reference,
            "newly_learned": summary.newly_learned,
            "unchanged": summary.unchanged,
            "unmatched": summary.unmatched,
            "unmatched_rows": summary.unmatched_rows,
        },
        "download_name": download_name,
    })


@app.route("/api/download/<job_id>")
def api_download(job_id: str):
    # 安全：job_id 一定是 32 字元 hex（UUID4），否則 reject
    if not JOB_ID_RE.match(job_id):
        abort(400)
    path = OUTPUT_DIR / f"{job_id}.xlsx"
    if not path.exists() or not path.is_file():
        abort(404)
    # 讀回友善檔名（給瀏覽器顯示用）；Flask/Werkzeug 會自動做 RFC 5987 編碼，中文 OK
    name_file = OUTPUT_DIR / f"{job_id}.name"
    if name_file.exists():
        try:
            download_name = name_file.read_text(encoding="utf-8").strip() or "fixed.xlsx"
        except OSError:
            download_name = "fixed.xlsx"
    else:
        download_name = "fixed.xlsx"
    return send_file(
        str(path),
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------- API：版本檢查（自動更新通知） ----------

_VERSION_CACHE: dict = {"latest": None, "checked_at": 0.0}
_VERSION_CACHE_TTL = 4 * 60 * 60  # 4 小時:GitHub API rate limit 60/hr,夠用有餘


def _parse_version(s: str | None) -> tuple:
    if not s:
        return (0,)
    parts = s.lstrip("vV").strip().split(".")
    try:
        return tuple(int(p) for p in parts)
    except ValueError:
        return (0,)


def _is_newer(latest: str | None, current: str) -> bool:
    if not latest:
        return False
    return _parse_version(latest) > _parse_version(current)


def _fetch_latest_version() -> str | None:
    """打 GitHub Releases API。失敗(網路斷、rate limit、無 release)就回 None。"""
    try:
        req = urllib.request.Request(
            RELEASES_API_URL,
            headers={"User-Agent": f"grc-name-fixer/{VERSION}"},
        )
        with urllib.request.urlopen(req, timeout=3) as resp:
            data = json.load(resp)
            return data.get("tag_name")
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, OSError):
        return None


def _get_latest_version_cached() -> str | None:
    now = time.time()
    if now - _VERSION_CACHE["checked_at"] < _VERSION_CACHE_TTL:
        return _VERSION_CACHE["latest"]
    latest = _fetch_latest_version()
    _VERSION_CACHE["latest"] = latest
    _VERSION_CACHE["checked_at"] = now
    return latest


@app.route("/api/version")
def api_version():
    latest = _get_latest_version_cached()
    return jsonify({
        "current": VERSION,
        "latest": latest,
        "has_update": _is_newer(latest, VERSION),
        "release_url": RELEASES_PAGE_URL,
    })


# ---------- API：字典管理 ----------

@app.route("/api/dictionary", methods=["GET"])
def api_dict_list():
    search = request.args.get("q", "").strip()
    items = database.list_all(search=search)
    return jsonify({
        "ok": True,
        "count": len(items),
        "items": items,
    })


@app.route("/api/dictionary", methods=["POST"])
def api_dict_add():
    data = request.get_json(silent=True) or {}
    abbr = (data.get("abbreviation") or "").strip()
    full = (data.get("full_name") or "").strip()
    note = (data.get("note") or "").strip() or None
    if not abbr or not full:
        return jsonify({"ok": False, "error": "縮寫與完整名稱皆為必填"}), 400
    status = database.add_or_update(abbr, full, source="manual", note=note)
    return jsonify({"ok": True, "status": status})


@app.route("/api/dictionary/<path:abbreviation>", methods=["DELETE"])
def api_dict_delete(abbreviation: str):
    ok = database.delete(abbreviation)
    return jsonify({"ok": ok})


@app.route("/api/dictionary/count")
def api_dict_count():
    return jsonify({"count": database.count()})


# ---------- Error handler ----------

@app.errorhandler(413)
def too_large(e):
    return jsonify({"ok": False, "error": "檔案過大（最大 20 MB）"}), 413


# ---------- 啟動 ----------

if __name__ == "__main__":
    database.init_db()
    print("=" * 60)
    print("GRC 團名修正工具已啟動")
    print(f"請在瀏覽器打開: http://127.0.0.1:5000")
    print("關閉請按 Ctrl+C 或直接關閉此視窗")
    print("=" * 60)
    app.run(host="127.0.0.1", port=5000, debug=False)
