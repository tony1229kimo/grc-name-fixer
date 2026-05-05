"""
SQLite 字典資料庫
儲存「縮寫 → 完整團名」的學習對應
"""
import os
import sqlite3
from pathlib import Path
from datetime import datetime
from contextlib import contextmanager


def _resolve_db_path() -> Path:
    """
    資料庫位置：
    - 若有設定 GRC_DATA_DIR（打包後由 launcher 注入） → 使用該資料夾
    - 否則用 module 所在目錄（開發模式）
    """
    data_dir = os.environ.get("GRC_DATA_DIR", "").strip()
    if data_dir:
        return Path(data_dir) / "dictionary.db"
    return Path(__file__).parent / "dictionary.db"


DB_PATH = _resolve_db_path()


def init_db():
    """初始化資料庫（第一次執行時建立）"""
    with get_conn() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS name_dictionary (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                abbreviation TEXT NOT NULL UNIQUE,
                full_name TEXT NOT NULL,
                source TEXT NOT NULL DEFAULT 'auto',
                usage_count INTEGER NOT NULL DEFAULT 1,
                first_seen TEXT NOT NULL DEFAULT (datetime('now','localtime')),
                last_used TEXT NOT NULL DEFAULT (datetime('now','localtime')),
                note TEXT
            );

            CREATE INDEX IF NOT EXISTS idx_abbreviation
                ON name_dictionary(abbreviation);
        """)


@contextmanager
def get_conn():
    """取得資料庫連線（context manager）"""
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def lookup(abbreviation: str) -> str | None:
    """查詢縮寫對應的完整名稱；同時更新使用次數。"""
    if not abbreviation:
        return None
    with get_conn() as conn:
        row = conn.execute(
            "SELECT full_name FROM name_dictionary WHERE abbreviation = ?",
            (abbreviation,)
        ).fetchone()
        if row:
            conn.execute(
                "UPDATE name_dictionary SET usage_count = usage_count + 1, "
                "last_used = datetime('now','localtime') WHERE abbreviation = ?",
                (abbreviation,)
            )
            return row["full_name"]
    return None


def add_or_update(abbreviation: str, full_name: str, source: str = "auto",
                  note: str | None = None) -> str:
    """
    新增或更新字典項目。
    回傳: 'added' / 'updated' / 'unchanged'
    """
    abbreviation = abbreviation.strip()
    full_name = full_name.strip()
    if not abbreviation or not full_name:
        return "unchanged"

    with get_conn() as conn:
        row = conn.execute(
            "SELECT full_name FROM name_dictionary WHERE abbreviation = ?",
            (abbreviation,)
        ).fetchone()

        if row is None:
            conn.execute(
                "INSERT INTO name_dictionary (abbreviation, full_name, source, note) "
                "VALUES (?, ?, ?, ?)",
                (abbreviation, full_name, source, note)
            )
            return "added"

        if row["full_name"] != full_name:
            conn.execute(
                "UPDATE name_dictionary SET full_name = ?, source = ?, "
                "last_used = datetime('now','localtime'), note = ? "
                "WHERE abbreviation = ?",
                (full_name, source, note, abbreviation)
            )
            return "updated"

        return "unchanged"


def list_all(search: str = "") -> list[dict]:
    """列出所有字典項目（可選擇搜尋）"""
    with get_conn() as conn:
        if search:
            like = f"%{search}%"
            rows = conn.execute(
                "SELECT * FROM name_dictionary "
                "WHERE abbreviation LIKE ? OR full_name LIKE ? "
                "ORDER BY last_used DESC",
                (like, like)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM name_dictionary ORDER BY last_used DESC"
            ).fetchall()
        return [dict(r) for r in rows]


def delete(abbreviation: str) -> bool:
    """刪除字典項目"""
    with get_conn() as conn:
        cur = conn.execute(
            "DELETE FROM name_dictionary WHERE abbreviation = ?",
            (abbreviation,)
        )
        return cur.rowcount > 0


def count() -> int:
    """總項目數"""
    with get_conn() as conn:
        row = conn.execute("SELECT COUNT(*) AS c FROM name_dictionary").fetchone()
        return row["c"]


if __name__ == "__main__":
    init_db()
    print(f"資料庫已建立於: {DB_PATH}")
    print(f"目前字典項目數: {count()}")
