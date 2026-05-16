"""
應用程式入口 - 給 PyInstaller 打包用
- 自動定位資源路徑（支援打包後的 .exe）
- 自動找可用 port（5000 占用就試 5001、5002...）
- 偵測雙開（已有 instance 在跑就提示並退出）
- 自動開啟瀏覽器
- 在系統匣 / 終端視窗顯示啟動資訊
"""
from __future__ import annotations

import atexit
import os
import socket
import sys
import tempfile
import threading
import time
import webbrowser
from pathlib import Path

from _version import VERSION


PORT_START = 5000
PORT_MAX_TRIES = 20
LOCK_FILE = Path(tempfile.gettempdir()) / "grc-name-fixer.lock"


def resource_path(relative: str) -> str:
    """取得資源真實路徑（支援 PyInstaller --onefile / --onedir）"""
    if getattr(sys, "frozen", False):
        # PyInstaller: sys._MEIPASS 指向解壓後的資源目錄
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    else:
        base = Path(__file__).parent
    return str(base / relative)


def data_path(relative: str) -> Path:
    """
    取得「可寫入」的資料路徑（dictionary.db、uploads、outputs）。
    打包後的程式應該把資料存在 .exe 所在目錄，而不是 _MEIPASS（會被清除）。
    """
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).parent
    else:
        base = Path(__file__).parent
    p = base / relative
    return p


def ensure_data_dirs():
    """建立必要的資料資料夾"""
    for name in ("uploads", "outputs", "logs"):
        d = data_path(name)
        d.mkdir(parents=True, exist_ok=True)


def copy_seed_db_if_missing():
    """
    第一次啟動時：若工作目錄沒有 dictionary.db，從 dictionary.seed.db 複製一份。
    之後就都用工作目錄的 dictionary.db（持續累積學習）。

    這樣設計的好處：升級時新 zip 解壓覆蓋會更新 dictionary.seed.db，
    但同事的 dictionary.db（含累積學習）不會被沖掉。
    """
    target = data_path("dictionary.db")
    if target.exists():
        return
    seeded = Path(resource_path("dictionary.seed.db"))
    if seeded.exists():
        import shutil
        shutil.copy2(str(seeded), str(target))
        print(f"[初始化] 已載入預灌字典 → {target.name}")


def find_free_port(start: int = PORT_START, max_tries: int = PORT_MAX_TRIES) -> int:
    """從 start 開始往上找可用 port。找不到就 raise。"""
    for p in range(start, start + max_tries):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                s.bind(("127.0.0.1", p))
                return p
        except OSError:
            continue
    raise RuntimeError(
        f"找不到可用 port（試了 {start}~{start + max_tries - 1}）"
    )


def detect_running_instance() -> int | None:
    """
    偵測是否已有 instance 在跑。
    回傳 port 表示有；回傳 None 表示沒有（lock 不存在或是 stale）。
    """
    if not LOCK_FILE.exists():
        return None
    try:
        existing_port = int(LOCK_FILE.read_text(encoding="utf-8").strip())
    except (ValueError, OSError):
        return None
    # 確認那個 port 真的有人在 listening（不是 stale lock）
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.5)
            if s.connect_ex(("127.0.0.1", existing_port)) == 0:
                return existing_port
    except OSError:
        pass
    return None  # stale lock，忽略


def write_lock(port: int):
    try:
        LOCK_FILE.write_text(str(port), encoding="utf-8")
    except OSError:
        pass


def remove_lock():
    try:
        LOCK_FILE.unlink()
    except OSError:
        pass


def open_browser_delayed(url: str, delay: float = 1.5):
    def _open():
        time.sleep(delay)
        try:
            webbrowser.open(url)
        except Exception:
            pass
    threading.Thread(target=_open, daemon=True).start()


def banner(port: int):
    url = f"http://127.0.0.1:{port}"
    print("=" * 62)
    print(f"  GRC 團名修正工具  v{VERSION}")
    print("  台中勤美洲際酒店")
    print("=" * 62)
    print("")
    print(f"  伺服器位置: {url}")
    print("  瀏覽器會自動開啟，如未開啟請手動貼上網址")
    print("")
    print("  ★ 要結束程式，請直接關閉這個黑色視窗 ★")
    print("")
    print("=" * 62)
    print("")


def banner_already_running(port: int):
    url = f"http://127.0.0.1:{port}"
    print("=" * 62)
    print("  ⚠ GRC 團名修正工具已經在執行中")
    print("=" * 62)
    print("")
    print(f"  請直接打開瀏覽器：{url}")
    print("")
    print("  ※ 若瀏覽器無法連上，可能是上次沒關乾淨：")
    print(f"     1. 找到所有「GRC-團名修正工具.exe」的黑色視窗並關閉")
    print(f"     2. 刪除暫存檔：{LOCK_FILE}")
    print(f"     3. 重新雙擊本程式")
    print("")
    print("=" * 62)
    # 嘗試幫他開瀏覽器導向已存在的那個 instance
    try:
        webbrowser.open(url)
    except Exception:
        pass


def main():
    # 切換工作目錄到 exe 所在目錄，避免相對路徑混亂
    if getattr(sys, "frozen", False):
        os.chdir(str(Path(sys.executable).parent))

    # 偵測雙開
    existing_port = detect_running_instance()
    if existing_port is not None:
        banner_already_running(existing_port)
        input("\n按 Enter 關閉這個視窗（不會影響已在執行的程式）...")
        return

    # 找可用 port
    try:
        port = find_free_port()
    except RuntimeError as e:
        print(f"\n[錯誤] {e}")
        print("可能是 port 5000~5019 全部被別的程式占用。")
        print("請關閉其他開發工具（Skype、其他 web server 等）後再試。")
        input("\n按 Enter 關閉...")
        return

    # 註冊 lock（程式結束時自動清除）
    write_lock(port)
    atexit.register(remove_lock)

    banner(port)
    ensure_data_dirs()
    copy_seed_db_if_missing()

    # 把路徑注入 Flask app 用的 BASE_DIR（透過環境變數）
    os.environ["GRC_DATA_DIR"] = str(data_path(""))
    os.environ["GRC_RESOURCE_DIR"] = str(Path(resource_path("")))

    # 延後匯入 app（確保 env 已設定）
    from app import app
    import database
    database.init_db()

    url = f"http://127.0.0.1:{port}"
    open_browser_delayed(url, delay=1.5)

    # 關閉 Flask 啟動訊息（已由 banner 取代）
    import logging
    logging.getLogger("werkzeug").setLevel(logging.WARNING)

    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[結束] 已關閉服務")
    except Exception as e:
        import traceback
        print("\n[錯誤] 程式發生問題：")
        traceback.print_exc()
        # 也寫到 logs/error.log，方便同事回報
        try:
            log = data_path("logs")
            log.mkdir(parents=True, exist_ok=True)
            from datetime import datetime
            with open(log / "error.log", "a", encoding="utf-8") as f:
                f.write(f"\n=== [{datetime.now():%Y-%m-%d %H:%M:%S}] launcher 啟動失敗 ===\n")
                f.write(traceback.format_exc())
                f.write("\n")
        except Exception:
            pass
        input("\n按 Enter 關閉...")
