"""
modules/kho.py — Gộp tất cả file kho thành Kho_Sach.xlsx.

Public API:
    run(config: Config) -> dict
"""
import glob
import logging
import os
import tempfile
from pathlib import Path

import pandas as pd

from config import Config
from utils.phone import clean_phone

logger = logging.getLogger(__name__)

_COT_OUTPUT = ['stt', 'msisdn', 'tên kho', 'hạng số', 'trạng thái', 'ngày bán', 'mã lô bán', 'ghi chú']
_RENAME = {
    'stt': 'STT',
    'msisdn': 'MSISDN',
    'tên kho': 'Tên Kho',
    'hạng số': 'Hạng Số',
    'trạng thái': 'Trạng Thái',
    'ngày bán': 'Ngày Bán',
    'mã lô bán': 'Mã Lô Bán',
    'ghi chú': 'Ghi Chú',
}


# ── PUBLIC ─────────────────────────────────────────────────────────────────────

def process_files(uploaded_files, names: dict | None = None, config: Config | None = None):
    names = names or {}
    errors = []
    frames = []

    with tempfile.TemporaryDirectory() as tmp:
        for f in uploaded_files:
            try:
                # luôn ưu tiên getvalue
                data = f.getvalue() if hasattr(f, "getvalue") else f.read()

                if not data:
                    errors.append(f"{f.name}: file rỗng")
                    continue

                fp = os.path.join(tmp, f.name)
                Path(fp).write_bytes(data)

                ten_kho = (names.get(f.name) or "").strip() or Path(f.name).stem

                config = config or Config()
                df = _read_file(fp, config, ten_kho_override=ten_kho)

                if df is not None and not df.empty:
                    frames.append(df)
                else:
                    errors.append(f"Bỏ qua {f.name}: không đọc được dữ liệu")

            except Exception as e:
                errors.append(f"Lỗi {f.name}: {str(e)}")

    if not frames:
        return {
            "success": False,
            "df": None,
            "count": 0,
            "errors": errors,
            "message": "Không đọc được dữ liệu",
        }

    df = pd.concat(frames, ignore_index=True)

    # FIX nhẹ: tránh duplicate column crash
    df = df.loc[:, ~df.columns.duplicated()]
    
    is_dup = df.duplicated(subset="msisdn", keep=False)
    if is_dup.any():
        # Tạo cột ghi chú nếu chưa có
        if "ghi chú" not in df.columns:
            df["ghi chú"] = ""
            
        # Nhóm các số bị trùng, ghép "tên kho" của chúng lại với nhau
        dup_info = df[is_dup].groupby("msisdn")["tên kho"].apply(
            lambda x: "Trùng tại: " + " + ".join(x.dropna().unique())
        )
        
        # Map ghi chú này ngược lại vào bảng df
        df.loc[is_dup, "ghi chú"] = df.loc[is_dup, "msisdn"].map(dup_info)

    df = df.drop_duplicates(subset="msisdn", keep="first")
    df = _format_output(df)

    return {
        "success": True,
        "df": df,
        "count": len(df),
        "errors": errors,
        "message": f"{len(df):,} MSISDN từ {len(frames)} file",
    }


def run(config: Config) -> dict:
    logger.info("═" * 50)
    logger.info(f"KHO | Quét: {config.thu_muc_kho}")

    errors = []
    files = _list_files(config)

    if not files:
        return _fail("Không tìm thấy file .xlsx hợp lệ trong thư mục kho", errors)

    frames = []

    for fp in files:
        df = _read_file(fp, config)

        if df is not None and not df.empty:
            frames.append(df)
            logger.info(f"✓ {Path(fp).stem}: {len(df):,} dòng")
        else:
            errors.append(f"Bỏ qua: {os.path.basename(fp)}")

    if not frames:
        return _fail("Không đọc được dữ liệu từ bất kỳ file nào", errors)

    master = pd.concat(frames, ignore_index=True)

    before = len(master)
    master = master.drop_duplicates(subset="msisdn", keep="first")
    dup = before - len(master)

    if dup:
        errors.append(f"Loại {dup:,} MSISDN trùng")

    master = _format_output(master)

    config.ensure_dirs()
    master.to_excel(config.file_kho_sach, index=False)

    msg = f"Hoàn thành: {len(master):,} MSISDN từ {len(frames)} file"

    return {
        "success": True,
        "count": len(master),
        "output_file": config.file_kho_sach,
        "errors": errors,
        "message": msg,
    }


# ── PRIVATE ────────────────────────────────────────────────────────────────────

def _list_files(config: Config) -> list[str]:
    return [
        f for f in glob.glob(os.path.join(config.thu_muc_kho, "*.xlsx"))
        if not os.path.basename(f).startswith("~$")
        and os.path.basename(f) != config.ten_file_kho_sach
    ]


def _read_file(fp: str, config: Config, ten_kho_override: str = ""):
    name = Path(fp).stem

    try:
        try:
            xls = pd.ExcelFile(fp, engine=config.kho_engine)
        except Exception:
            xls = pd.ExcelFile(fp, engine="openpyxl")

        sheets = xls.sheet_names
        sheets_lower = {s.strip().lower(): s for s in sheets}  
        chuan = config.kho_sheet_chuan.strip().lower()
        du_phong = [s.strip().lower() for s in config.kho_sheet_du_phong]

        if chuan in sheets_lower:
            to_read = [sheets_lower[chuan]]
        else:
            to_read = [sheets_lower[s] for s in du_phong if s in sheets_lower]

        if not to_read:
            logger.warning(f"⚠ {name}: không khớp sheet nào. Sheets có: {sheets}")
            return None

        parts = [_read_sheet(xls, s, ten_kho_override or name) for s in to_read]
        parts = [p for p in parts if p is not None]

        return pd.concat(parts, ignore_index=True) if parts else None

    except Exception:
        return None


def _read_sheet(xls, sheet: str, kho_name: str):
    header = _find_header(xls, sheet, "msisdn")

    if header is None:
        return None

    df = pd.read_excel(xls, sheet_name=sheet, header=header, dtype=str)
    df.columns = df.columns.str.strip().str.lower()

    if "msisdn" not in df.columns:
        return None

    df = df.dropna(subset=["msisdn"])

    df["msisdn"] = (
        df["msisdn"]
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.strip()
    )

    df = df[~df["msisdn"].str.lower().isin(["nan", "none", ""])]
    df = df[df["msisdn"].str.fullmatch(r"\d{9,11}")]

    if df.empty:
        return None

    # FIX: không tạo cột tùy ý trong sheet layer (an toàn hơn)
    df["tên kho"] = kho_name

    return df


def _find_header(xls, sheet: str, keyword: str):
    for h in range(8):
        try:
            tmp = pd.read_excel(xls, sheet_name=sheet, header=h, nrows=0)
            if keyword in tmp.columns.str.strip().str.lower().tolist():
                return h
        except Exception:
            continue
    return None


def _format_output(df: pd.DataFrame):
    df = df.drop(columns=["stt"], errors="ignore")

    df.insert(0, "stt", range(1, len(df) + 1))

    # FIX quan trọng: bỏ duplicate columns
    df = df.loc[:, ~df.columns.duplicated()]

    # đảm bảo đủ cột
    for col in _COT_OUTPUT:
        if col not in df.columns:
            df[col] = None

    df = df[_COT_OUTPUT].rename(columns=_RENAME)

    return df


def _fail(msg: str, errors: list):
    return {
        "success": False,
        "count": 0,
        "output_file": None,
        "errors": errors,
        "message": msg,
    }