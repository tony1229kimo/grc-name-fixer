"""
Excel 解析器
- parse_flat_reference(): 解析「扁平參考檔」(0.xlsx 風格)
- parse_calendar_report(): 解析「日曆格式 GRC 報表」(Grc (x).xlsx 風格)
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime, date
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter


# ---------- 資料結構 ----------

@dataclass
class FlatEntry:
    """扁平參考檔的一筆資料"""
    full_name: str
    arrival: date
    guestroom_revenue: float
    event_revenue: float
    total_revenue: float
    status: str


@dataclass
class CalendarEntry:
    """日曆報表的一筆團體資料"""
    sheet_name: str
    row: int                 # 該列在 sheet 中的 row number（用於後續寫回）
    raw_name: str            # 原始（可能是縮寫）的名稱
    arrival: date | None     # 第一個有房數的日期
    total_rooms: float | None
    total_revenue: float | None
    section: str             # 'Definite' / 'Tentative' / 'Prospect'


@dataclass
class CalendarSheetMeta:
    """日曆分頁的欄位 metadata"""
    sheet_name: str
    year: int
    month: int
    header_row: int          # 含日期數字的那一列 (通常是 R12)
    name_col: int            # 名稱所在欄（通常 A = 1）
    day_col_map: dict[int, int] = field(default_factory=dict)  # day -> col number
    total_rev_col: int | None = None
    total_rms_col: int | None = None
    data_rows: list[int] = field(default_factory=list)         # 實際資料列
    section_row_map: dict[int, str] = field(default_factory=dict)  # row -> section name


# ---------- 扁平參考檔 ----------

EXPECTED_FLAT_HEADERS = {
    "full_name": ["Booking: Booking Post As", "Booking Post As", "Post As"],
    "guestroom_rev": ["Blended Guestroom Revenue Total"],
    "event_rev": ["Blended Event Revenue Total"],
    "total_rev": ["Blended Revenue Total"],
    "arrival": ["Arrival"],
    "status": ["Status"],
}


def _match_header(cell_value: str, candidates: list[str]) -> bool:
    if not cell_value:
        return False
    val = str(cell_value).strip().lower()
    return any(val == c.lower() for c in candidates)


def parse_flat_reference(file_path: str | Path) -> list[FlatEntry]:
    """解析扁平參考檔（0.xlsx 格式）"""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active

    # 找出欄位 index (只看前 10 欄即可)
    col_idx = {}
    header_row = 1
    for col in range(1, min(ws.max_column + 1, 20)):
        val = ws.cell(row=header_row, column=col).value
        if not val:
            continue
        for key, candidates in EXPECTED_FLAT_HEADERS.items():
            if _match_header(str(val), candidates) and key not in col_idx:
                col_idx[key] = col
                break

    required = ["full_name", "arrival", "guestroom_rev"]
    missing = [k for k in required if k not in col_idx]
    if missing:
        raise ValueError(f"參考檔缺少必要欄位: {missing}. 找到的欄位: {col_idx}")

    entries: list[FlatEntry] = []
    for row in range(header_row + 1, ws.max_row + 1):
        name = ws.cell(row=row, column=col_idx["full_name"]).value
        if not name:
            continue

        arrival_val = ws.cell(row=row, column=col_idx["arrival"]).value
        arrival = _to_date(arrival_val)
        if arrival is None:
            continue

        guestroom = _to_float(ws.cell(row=row, column=col_idx["guestroom_rev"]).value)
        event_rev = _to_float(
            ws.cell(row=row, column=col_idx["event_rev"]).value
        ) if "event_rev" in col_idx else 0.0
        total_rev = _to_float(
            ws.cell(row=row, column=col_idx["total_rev"]).value
        ) if "total_rev" in col_idx else (guestroom + event_rev)
        status = str(ws.cell(row=row, column=col_idx["status"]).value or "") \
            if "status" in col_idx else ""

        entries.append(FlatEntry(
            full_name=str(name).strip(),
            arrival=arrival,
            guestroom_revenue=guestroom,
            event_revenue=event_rev,
            total_revenue=total_rev,
            status=status.strip(),
        ))
    return entries


# ---------- 日曆格式報表 ----------

MONTH_PATTERN = re.compile(
    r"(?i)(January|February|March|April|May|June|July|August|"
    r"September|October|November|December|"
    r"Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s*-\s*(\d{4})"
)

MONTH_MAP = {m.lower(): i for i, m in enumerate([
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
], start=1)}
MONTH_MAP.update({m.lower(): i for i, m in enumerate([
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
], start=1)})

SECTION_KEYWORDS = ("Definite", "Tentative", "Prospect")
TOTAL_KEYWORDS = ("Definite Totals", "Tentative Totals", "Prospect Totals",
                  "Monthly Totals", "Totals")


def _detect_section_header(cell_value: str) -> str | None:
    """判斷 cell 內容是否為分段標題。回傳 section 名稱或 None"""
    if not cell_value:
        return None
    val = str(cell_value).strip()
    for kw in SECTION_KEYWORDS:
        if val == kw:
            return kw
    return None


def _is_totals_row(cell_value: str) -> bool:
    if not cell_value:
        return False
    val = str(cell_value).strip()
    return any(val == kw or val.endswith(" Totals") for kw in TOTAL_KEYWORDS)


def _parse_month_year(cell_value: str) -> tuple[int, int] | None:
    if not cell_value:
        return None
    m = MONTH_PATTERN.search(str(cell_value))
    if not m:
        return None
    mon = MONTH_MAP.get(m.group(1).lower())
    year = int(m.group(2))
    if mon:
        return year, mon
    return None


def _analyze_calendar_sheet(ws) -> CalendarSheetMeta | None:
    """分析一個月份分頁，回傳 metadata。無法辨識則回傳 None。"""
    # 先掃前 20 列找 Month-Year 與 day header
    year_month: tuple[int, int] | None = None
    month_row: int | None = None
    for row in range(1, min(ws.max_row + 1, 20)):
        for col in range(1, min(ws.max_column + 1, 5)):
            v = ws.cell(row=row, column=col).value
            ym = _parse_month_year(v if isinstance(v, str) else "")
            if ym:
                year_month = ym
                month_row = row
                break
        if year_month:
            break

    if not year_month:
        return None

    year, month = year_month

    # 找 day header row：從 month_row+1 開始找一列含多個小整數 (1..31) 的列
    # （注意：Delphi 輸出可能把數字存成字串，所以兩種都要接受）
    header_row = None
    day_col_map: dict[int, int] = {}
    for row in range(month_row + 1, min(month_row + 6, ws.max_row + 1)):
        tmp_map = {}
        for col in range(1, ws.max_column + 1):
            v = ws.cell(row=row, column=col).value
            day_num = _to_day_int(v)
            if day_num is not None:
                tmp_map[day_num] = col
        if len(tmp_map) >= 20:   # 至少 20 天 → 認定為日期標頭列
            header_row = row
            day_col_map = tmp_map
            break

    if not header_row:
        return None

    # 找 Total Rev / Rms 欄（從 header_row 或 header_row+1 找「Rev」「Rms」）
    total_rev_col = None
    total_rms_col = None
    for probe_row in (header_row, header_row + 1):
        for col in range(1, ws.max_column + 1):
            v = ws.cell(row=probe_row, column=col).value
            if isinstance(v, str):
                val = v.strip().lower()
                if val == "rev" and total_rev_col is None:
                    total_rev_col = col
                if val == "rms" and total_rms_col is None:
                    total_rms_col = col

    # 找 name column：通常 A 欄，也從 header_row 往下看哪一欄有文字名稱
    name_col = 1

    return CalendarSheetMeta(
        sheet_name=ws.title,
        year=year,
        month=month,
        header_row=header_row,
        name_col=name_col,
        day_col_map=day_col_map,
        total_rev_col=total_rev_col,
        total_rms_col=total_rms_col,
    )


def _scan_data_rows(ws, meta: CalendarSheetMeta) -> list[CalendarEntry]:
    """從分析好的 meta 掃描所有資料列"""
    entries: list[CalendarEntry] = []
    current_section = "Definite"  # 預設第一段是 Definite

    # 從 header_row 往下掃，遇到 totals 就切換 section
    row = meta.header_row + 1
    max_row = ws.max_row + 1

    while row < max_row:
        name_val = ws.cell(row=row, column=meta.name_col).value
        name_str = str(name_val).strip() if name_val else ""

        # 檢查是否為 section 標題（Definite/Tentative/Prospect）
        sec = _detect_section_header(name_str)
        if sec:
            current_section = sec
            # 跳過標題列 + 週間日列（通常緊接下一列）
            row += 2
            continue

        # 檢查是否為 totals 列
        if _is_totals_row(name_str):
            # 如果是 Monthly Totals，後面可能還有 forecast/budget 等，但我們不需要
            if "Monthly" in name_str:
                break
            row += 1
            continue

        if not name_str:
            row += 1
            continue

        # 這是一個資料列 → 解析
        # 找第一個有房數的 day col
        arrival_day = None
        for day in sorted(meta.day_col_map.keys()):
            col = meta.day_col_map[day]
            v = ws.cell(row=row, column=col).value
            if v is not None and _to_float(v) > 0:
                arrival_day = day
                break

        arrival_date = None
        if arrival_day:
            try:
                arrival_date = date(meta.year, meta.month, arrival_day)
            except ValueError:
                arrival_date = None

        total_rev = None
        if meta.total_rev_col:
            total_rev = _to_float(ws.cell(row=row, column=meta.total_rev_col).value)

        total_rms = None
        if meta.total_rms_col:
            total_rms = _to_float(ws.cell(row=row, column=meta.total_rms_col).value)

        entries.append(CalendarEntry(
            sheet_name=meta.sheet_name,
            row=row,
            raw_name=name_str,
            arrival=arrival_date,
            total_rooms=total_rms,
            total_revenue=total_rev,
            section=current_section,
        ))

        row += 1

    return entries


def parse_calendar_report(file_path: str | Path) -> tuple[list[CalendarEntry], dict[str, CalendarSheetMeta]]:
    """
    解析日曆格式 GRC 報表。
    回傳 (所有資料列 entries, 每個月份分頁的 meta)
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    all_entries: list[CalendarEntry] = []
    all_meta: dict[str, CalendarSheetMeta] = {}

    for sheet_name in wb.sheetnames:
        if sheet_name.lower() in ("document map",):
            continue
        ws = wb[sheet_name]
        meta = _analyze_calendar_sheet(ws)
        if not meta:
            continue
        entries = _scan_data_rows(ws, meta)
        all_entries.extend(entries)
        all_meta[sheet_name] = meta

    return all_entries, all_meta


# ---------- Helpers ----------

def _to_date(v) -> date | None:
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        v = v.strip()
        for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(v, fmt).date()
            except ValueError:
                continue
    return None


def _to_day_int(v) -> int | None:
    """判斷是否為 1-31 的日期數字（可能是 int 或字串）"""
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return v if 1 <= v <= 31 else None
    if isinstance(v, float):
        if v.is_integer():
            iv = int(v)
            return iv if 1 <= iv <= 31 else None
        return None
    if isinstance(v, str):
        s = v.strip()
        if s.isdigit():
            iv = int(s)
            return iv if 1 <= iv <= 31 else None
    return None


def _to_float(v) -> float:
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


# ---------- CLI 測試 ----------

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python parser.py <xlsx-file> [--flat|--calendar]")
        sys.exit(1)

    path = sys.argv[1]
    mode = sys.argv[2] if len(sys.argv) > 2 else "--auto"

    if mode in ("--flat", "--auto"):
        try:
            flat = parse_flat_reference(path)
            print(f"\n=== Flat entries: {len(flat)} ===")
            for e in flat[:5]:
                print(f"  {e.full_name[:40]:40s} | {e.arrival} | "
                      f"guestroom={e.guestroom_revenue} total={e.total_revenue}")
            if mode == "--flat":
                sys.exit(0)
        except Exception as ex:
            if mode == "--flat":
                raise
            print(f"Not a flat reference file: {ex}")

    if mode in ("--calendar", "--auto"):
        cal_entries, metas = parse_calendar_report(path)
        print(f"\n=== Calendar entries: {len(cal_entries)} across {len(metas)} sheets ===")
        for sn, m in metas.items():
            print(f"  sheet={sn} | year-month={m.year}-{m.month:02d} | "
                  f"header_row={m.header_row} | rev_col={m.total_rev_col}")
        print("\nFirst 8 entries:")
        for e in cal_entries[:8]:
            print(f"  [{e.section:9s}] R{e.row:3d} {e.raw_name[:30]:30s} | "
                  f"arrival={e.arrival} | rms={e.total_rooms} | rev={e.total_revenue}")
