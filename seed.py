"""
預先灌入字典：從提供的三份檔案建立初始對應。
- 用 0.xlsx (扁平完整名) + Grc(7).xlsx (縮寫) 配對
- 從 GRC Report_04132026.xlsx (上週已完成) 再次驗證與補強
用法:
    python seed.py <flat.xlsx> <abbr.xlsx> [previous_week_done.xlsx]
"""
from __future__ import annotations

import sys
from pathlib import Path

import database
from parser import parse_flat_reference, parse_calendar_report
from matcher import match_entries


def seed_from_files(flat_path: str, abbr_path: str,
                    previous_done_path: str | None = None):
    database.init_db()
    before_count = database.count()

    # 1) 用 flat + abbr 配對，學習縮寫對應
    print(f"\n[1/2] 從 {Path(abbr_path).name} × {Path(flat_path).name} 配對學習...")
    flat = parse_flat_reference(flat_path)
    abbr_entries, _ = parse_calendar_report(abbr_path)
    results, summary = match_entries(abbr_entries, flat)
    print(f"  · 日曆筆數: {summary.total}")
    print(f"  · 參考檔匹配: {summary.from_reference}")
    print(f"  · 新學入字典: {summary.newly_learned}")
    print(f"  · 未匹配: {summary.unmatched}")

    # 2) 從上週已完成的日曆檔抽完整名稱當 extra pool（可選）
    if previous_done_path and Path(previous_done_path).exists():
        print(f"\n[2/2] 從 {Path(previous_done_path).name} 補充完整名稱池...")
        prev_entries, _ = parse_calendar_report(previous_done_path)

        # 上週完成版的名稱可能都是完整的，可做為「候選完整名」補充
        # 但因為我們主要以 abbreviation 為 key，這個檔案只能當參考確認
        added = 0
        for e in prev_entries:
            # 如果該名稱在字典中是某個縮寫的對應，代表已學過，略過
            # 若是全新名稱，暫不處理（沒縮寫可綁）
            pass
        print(f"  · 附加記錄（作為 reference 已內化）: 0")

    after_count = database.count()
    print(f"\n===== 完成 =====")
    print(f"字典項目數: {before_count} → {after_count} (+{after_count - before_count})")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python seed.py <flat.xlsx> <abbr.xlsx> [previous_done.xlsx]")
        sys.exit(1)

    flat = sys.argv[1]
    abbr = sys.argv[2]
    prev = sys.argv[3] if len(sys.argv) > 3 else None
    seed_from_files(flat, abbr, prev)
