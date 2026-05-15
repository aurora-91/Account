"""
modules/so_ban.py — Gộp số bán từ các file biên bản → So_Ban.xlsx.

Hỗ trợ quét TẤT CẢ các sheet có cột STB:
  • Tự động lấy tên sheet làm Mã lô (nếu có).
  • Đọc từ trên xuống trong cột STB: gặp chữ -> nhớ Mã lô, gặp số -> gán Mã lô.

Public API:
    run(config, date_overrides=None)  -> dict
    scan_files(config)                -> list[dict]   ← dùng cho web UI
"""
import logging
import os
import re
import tempfile
from pathlib import Path
from typing import Optional

import pandas as pd

from config import Config
from utils.date_extractor import extract_date
from utils.phone import clean_phone

logger = logging.getLogger(__name__)


# ── PUBLIC ─────────────────────────────────────────────────────────────────────

def process_files(uploaded_files, dates: dict | None = None):
    dates = dates or {}

    with tempfile.TemporaryDirectory() as tmp:
        for f in uploaded_files:
            data = f.getvalue() if hasattr(f, 'getvalue') else f.read()
            fp = os.path.join(tmp, f.name)
            Path(fp).write_bytes(data)

        config = Config(thu_muc_bienban=tmp, output_folder=tmp)
        result = run(config, date_overrides=dates)

        if not result['success']:
            return {
                "success": False,
                "df": None,
                "count": 0,
                "errors": result['errors'],
                "message": result['message']
            }

        return {
            "success": True,
            "df": pd.DataFrame(result["rows"]),
            "count": result["count"],
            "errors": result["errors"],
            "message": result["message"]
        }

def run(config: Config, date_overrides: Optional[dict[str, str]] = None) -> dict:
    logger.info("═" * 50)
    logger.info(f"SO_BAN | Quét: {config.thu_muc_bienban}")

    date_overrides = date_overrides or {}
    errors: list[str] = []

    files = _list_files(config)
    if not files:
        return _fail("Không tìm thấy file .xlsx hợp lệ", errors)

    n = len(files)
    if config.so_ban_so_file_test > 0:
        files = files[:config.so_ban_so_file_test]

    all_records: list[dict] = []

    for fp in files:
        fname = os.path.basename(fp)
        manual = date_overrides.get(fname, "")
        recs = _read_file(fp, config, manual)

        # 🔥 FIX: tránh None
        if recs:
            all_records.extend(recs)

        logger.info(f"✓ {fname}: {len(recs)} số")

    if not all_records:
        return _fail("Không lấy được dữ liệu STB", errors)

    df_all = pd.DataFrame(all_records)

    # 🔥 FIX SORT SAFE
    df_all["ngay_ban"] = df_all["ngay_ban"].fillna("")
    df_all = df_all.sort_values("ngay_ban", key=lambda s: s.map(_sort_key))

    rows: list[dict] = []

    for stt, (so, grp) in enumerate(df_all.groupby("msisdn", sort=False), start=1):
        r0 = grp.iloc[0]
        sl = len(grp)
        
        note = ""
        if sl > 1:
            # Lấy danh sách các tháng/file duy nhất mà số này xuất hiện
            unique_dates = grp["ngay_ban"].replace("", pd.NA).dropna().unique().tolist()
            
            if len(unique_dates) == 1:
                # Nếu chỉ có 1 tháng duy nhất -> Bị lặp trong cùng 1 file
                note = f"Lặp {sl} lần trong cùng: {unique_dates[0]}"
            else:
                # Nếu có nhiều tháng -> Bị lặp chéo giữa các file
                note = "Lặp chéo tại: " + " & ".join(unique_dates)
                
        rows.append({
            "STT": stt,
            "STB": so,
            "Ngày bán": r0.get("ngay_ban", ""),
            "Mã lô": r0.get("ma_lo", ""),
            "Số lần xuất hiện": len(grp),
            "Ghi chú": note
        })

    df_out = pd.DataFrame(rows)

    config.ensure_dirs()
    df_out.to_excel(config.file_so_ban, index=False)

    so_lap = int((df_out["Số lần xuất hiện"] > 1).sum())

    msg = f"Hoàn thành: {len(rows)} số | {so_lap} số lặp"

    return {
        "success": True,
        "count": len(rows),
        "output_file": config.file_so_ban,
        "rows": rows,   # 🔥 CÁI QUAN TRỌNG NHẤT CHO UI
        "errors": errors,
        "message": msg
    }


def scan_files(config: Config) -> list[dict]:
    return [
        {
            "filename":      os.path.basename(fp),
            "folder":        os.path.basename(os.path.dirname(fp)),
            "detected_date": extract_date(
                filename=os.path.basename(fp),
                folder_name=os.path.basename(os.path.dirname(fp)),
            ),
        }
        for fp in _list_files(config)
    ]


# ── FILE READING ───────────────────────────────────────────────────────────────

def _list_files(config: Config) -> list[str]:
    files: list[str] = []
    for root, _, names in os.walk(config.thu_muc_bienban):
        for name in names:
            if name.endswith('.xlsx') and not name.startswith('~$'):
                files.append(os.path.join(root, name))
    files.sort(key=lambda fp: (
        _sort_key(os.path.basename(os.path.dirname(fp))),
        os.path.basename(fp).lower(),
    ))
    return files

def _read_file(fp: str, config: Config, manual_date: str = "") -> list[dict]:
    fname = os.path.basename(fp)
    folder = os.path.basename(os.path.dirname(fp))
    ngay_ban = extract_date(filename=fname, folder_name=folder, manual_override=manual_date)

    try:
        xls = pd.ExcelFile(fp)
        recs = []

        sheets_map = {str(s).lower().strip(): s for s in xls.sheet_names}

        chitiet_sheet = None
        for s_lower, s_orig in sheets_map.items():
            if _matches_any(s_lower, config.chitiet_keywords):
                chitiet_sheet = s_orig
                break

        if not chitiet_sheet:
            return []

        df_chitiet = _load_sheet(xls, chitiet_sheet, config.so_ban_tu_khoa_cot)
        if df_chitiet is None:
            return []

        stb_cols = [i for i, c in enumerate(df_chitiet.columns) if config.so_ban_tu_khoa_cot in c]
        lo_cols = [i for i, c in enumerate(df_chitiet.columns) if c in ['mã lô','ma lo','malo','lô','lo']]

        for _, row in df_chitiet.iterrows():
            so = ''
            for c in stb_cols:
                so = clean_phone(str(row.iloc[c]).strip())
                if so:
                    break

            if not so:
                continue

            ma_lo = ''
            for c in lo_cols:
                val = str(row.iloc[c]).strip()
                if val.lower() not in ('nan','none',''):
                    ma_lo = val
                    break

            recs.append({
                'msisdn': so,
                'ngay_ban': ngay_ban,
                'ma_lo': ma_lo,
                'file': fname,
            })

        return recs

    except Exception as e:
        logger.error(f"[{fname}] Lỗi: {e}")
        return []

# ── HELPERS ────────────────────────────────────────────────────────────────────

def _load_sheet(xls, sheet: str, keyword: str) -> pd.DataFrame | None:
    header = _find_header_keyword(xls, sheet, keyword)
    df = pd.read_excel(xls, sheet_name=sheet, header=header, dtype=str)
    df.columns = df.columns.astype(str).str.strip().str.lower()
    has_kw = any(keyword in c for c in df.columns)
    return df if has_kw else None


def _find_header_keyword(xls, sheet: str, keyword: str) -> int:
    best_h, best_score = 0, -1
    for h in range(10):
        try:
            tmp = pd.read_excel(xls, sheet_name=sheet, header=h, nrows=0)
            score = sum(1 for c in tmp.columns if keyword in str(c).lower())
            if score > best_score:
                best_score, best_h = score, h
        except Exception:
            continue
    return best_h if best_score > 0 else 0


def _extract_malo_from_name(sheet_name: str) -> str:
    m = re.search(r'(L[Oo]\d+|[Ll][ôo]\s*\d+|\d{3,})', sheet_name)
    return m.group(0).strip() if m else ""


def _matches_any(text: str, keywords: list[str]) -> bool:
    return any(kw.lower() == text.lower() for kw in keywords)


def _sort_key(text: str) -> tuple:
    nums = re.findall(r'\d+', str(text))
    if not nums:
        return (9999, 999)
    four = [n for n in nums if len(n) == 4]
    if four:
        year  = int(four[0])
        month = next((int(n) for n in nums if len(n) <= 2 and 1 <= int(n) <= 12), 1)
        return (year, month)
    return (9999, int(nums[0]))


def _fail(msg: str, errors: list) -> dict:
    logger.error(f"SO_BAN | {msg}")
    return {"success": False, "count": 0, "output_file": None, "errors": errors, "message": msg}