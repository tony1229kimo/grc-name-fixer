"""
應用程式入口 - 給 PyInstaller 打包用
- 自動定位資源路徑（支援打包後的 .exe）
- 自動開啟瀏覽器
- 在系統匣 / 終端視窗顯示啟動資訊
"""
from __future__ import annotations

import os
import sys
import threading
import time
import webbrowser
from pathlib import Path


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
    for name in ("uploads", "outputs"):
        d = data_path(name)
        d.mkdir(parents=True, exist_ok=True)


def copy_seed_db_if_missing():
    """
    第一次啟動時：若工作目錄沒有 dictionary.db，但打包資源中有預灌版 → 複製過去。
    之後就都用工作目錄的那份（可持續累積）。
    """
    target = data_path("dictionary.db")
    if target.exists():
        return
    seeded = Path(resource_path("dictionary.db"))
    if seeded.exists() and seeded.resolve() != target.resolve():
        import shutil
        shutil.copy2(str(seeded), str(target))
        print(f"[初始化] 已載入預灌字典 → {target.name}")


def open_browser_delayed(url: str, delay: float = 1.5):
    def _open():
        time.sleep(delay)
        try:
            webbrowser.open(url)
        except Exception:
            pass
    threading.Thread(target=_open, daemon=True).start()


def banner():
    print("=" * 62)
    print("  GRC 團名修正工具")
    print("  台中勤美洲際酒店")
    print("=" * 62)
    print("")
    print("  伺服器位置: http://127.0.0.1:5000")
    print("  瀏覽器會自動開啟，如未開啟請手動貼上網址")
    print("")
    print("  ★ 要結束程式，請直接關閉這個黑色視窗 ★")
    print("")
    print("=" * 62)
    print("")


def main():
    # 切換工作目錄到 exe 所在目錄，避免相對路徑混亂
    if getattr(sys, "frozen", False):
        os.chdir(str(Path(sys.executable).parent))

    banner()
    ensure_data_dirs()
    copy_seed_db_if_missing()

    # 把路徑注入 Flask app 用的 BASE_DIR（透過環境變數）
    os.environ["GRC_DATA_DIR"] = str(data_path(""))
    os.environ["GRC_RESOURCE_DIR"] = str(Path(resource_path("")))

    # 延後匯入 app（確保 env 已設定）
    from app import app
    import database
    database.init_db()

    port = 5000
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
        input("\n按 Enter 關閉...")
