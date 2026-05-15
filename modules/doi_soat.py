"""
modules/doi_soat.py — Đối soát Kho_Sach.xlsx với So_Ban.xlsx.

Public API:
    run(config: Config) -> dict
"""
import logging
import os

import pandas as pd

from config import Config
from utils.phone import clean_phone

logger = logging.getLogger(__name__)


# ── PUBLIC ─────────────────────────────────────────────────────────────────────

def run_from_df(df_kho: pd.DataFrame, df_ban: pd.DataFrame) -> dict:
    errors: list[str] = []
    df_kho = df_kho.copy()
    df_ban = df_ban.copy()

    # 1. Chuẩn hóa tên cột
    df_kho.columns = df_kho.columns.str.strip().str.lower()
    df_ban.columns = df_ban.columns.str.strip().str.lower()

    # 2. Tạo khóa phụ _p để đối soát
    df_kho['_p'] = df_kho['msisdn'].apply(clean_phone)
    phone_col = 'msisdn' if 'msisdn' in df_ban.columns else 'stb'
    df_ban['_p'] = df_ban[phone_col].apply(clean_phone)

    # 3. Chuẩn bị dữ liệu để Map từ Số Bán sang Kho
    idx_ban = df_ban.set_index('_p')
    
    # Lấy thông tin ngày, lô và ghi chú từ file Số Bán
    map_ngay = idx_ban['ngày bán'].to_dict() if 'ngày bán' in idx_ban.columns else \
               idx_ban['ngay_ban'].to_dict() if 'ngay_ban' in idx_ban.columns else {}
    map_lo   = idx_ban['mã lô'].to_dict()   if 'mã lô'   in idx_ban.columns else \
               idx_ban['ma_lo'].to_dict()   if 'ma_lo'   in idx_ban.columns else {}
    map_note_ban = idx_ban['ghi chú'].to_dict() if 'ghi chú' in idx_ban.columns else {}

    # 4. Xử lý cột trên df_kho
    # Đổi tên cột 'ghi chú' của kho thành 'Ghi Chú Kho' để phân biệt
    if 'ghi chú' in df_kho.columns:
        df_kho = df_kho.rename(columns={'ghi chú': 'Ghi Chú Kho'})
    else:
        df_kho['Ghi Chú Kho'] = ""

    # Map các thông tin bán hàng vào bảng kho
    df_kho['ngày bán']  = df_kho['_p'].map(map_ngay)
    df_kho['mã lô bán'] = df_kho['_p'].map(map_lo)
    df_kho['Ghi Chú Bán'] = df_kho['_p'].map(map_note_ban).fillna("")
    
    # Phân loại trạng thái đối soát
    df_kho['_loai'] = df_kho['ngày bán'].notna().map({True: 'co_trong_so_ban', False: 'khong_trong_so_ban'})

    matched   = int(df_kho['ngày bán'].notna().sum())
    unmatched = int(df_kho['ngày bán'].isna().sum())

    # 5. Xử lý "Số ngoài kho" (Có trong số bán nhưng không có trong kho)
    kho_phones = set(df_kho['_p'].dropna())
    df_ng = df_ban[~df_ban['_p'].isin(kho_phones)].copy()
    ngoai_kho = len(df_ng)

    if ngoai_kho:
        # Xác định các cột lấy dữ liệu từ df_ng
        ngay_col = 'ngày bán' if 'ngày bán' in df_ng.columns else 'ngay_ban' if 'ngay_ban' in df_ng.columns else None
        lo_col   = 'mã lô'   if 'mã lô'   in df_ng.columns else 'ma_lo'   if 'ma_lo'   in df_ng.columns else None
        note_col = 'ghi chú' if 'ghi chú' in df_ng.columns else None

        extra = pd.DataFrame({
            'msisdn':     df_ng['_p'],
            'tên kho':    'NGOÀI KHO', # Đánh dấu để dễ thấy
            'Ghi Chú Kho': '',
            'Ghi Chú Bán': df_ng[note_col] if note_col else '',
            'ngày bán':   df_ng[ngay_col] if ngay_col else '',
            'mã lô bán':  df_ng[lo_col]   if lo_col   else '',
            '_loai':      'ngoai_kho',
        })
        
        # Kết hợp kho và ngoài kho
        df_result = pd.concat([df_kho, extra], ignore_index=True)
    else:
        df_result = df_kho.copy()

    # 6. Dọn dẹp và định dạng hiển thị
    # Xóa khóa phụ _p
    if '_p' in df_result.columns: df_result = df_result.drop(columns=['_p'])
    
    # Định dạng msisdn: 84 -> 0
    df_result['msisdn'] = df_result['msisdn'].astype(str).str.replace(r'^(84|\+84)', '0', regex=True)
    
    # Đặt lại STT
    df_result = df_result.drop(columns=['stt', 'STT'], errors='ignore')
    df_result = df_result.reset_index(drop=True)
    df_result.insert(0, 'STT', range(1, len(df_result) + 1))
    
    # Chuẩn hóa tên cột để in ra (Viết hoa chữ đầu)
    df_result.columns = [
        'STT' if c.lower() == 'stt' else 
        'Phân Loại' if c == '_loai' else 
        c.title() for c in df_result.columns
    ]

    # Xóa cột trùng nếu có
    df_result = df_result.loc[:, ~df_result.columns.duplicated()]

    msg = f"Hoàn thành: {matched:,}-Bán | {unmatched:,}-Tồn | {ngoai_kho:,}-Lỗi ngoài kho"
    return {
        "success": True, "df_result": df_result,
        "matched": matched, "unmatched": unmatched, "ngoai_kho": ngoai_kho,
        "errors": errors, "message": msg,
    }


def run(config: Config) -> dict:
    """
    Match kho × số bán. Điền Ngày Bán + Mã Lô Bán vào các dòng khớp.

    Returns dict:
        success, count, matched, unmatched, output_file, errors, message
    """
    logger.info("═" * 50)
    logger.info("DOI_SOAT | Bắt đầu")

    errors: list[str] = []

    # Kiểm tra file đầu vào
    for fp, label in [(config.file_kho_sach, "Kho_Sach"), (config.file_so_ban, "So_Ban")]:
        if not os.path.exists(fp):
            return _fail(f"Không tìm thấy {label}: {fp}", errors)

    df_kho = pd.read_excel(config.file_kho_sach, dtype=str)
    df_ban = pd.read_excel(config.file_so_ban,   dtype=str)
    df_kho.columns = df_kho.columns.str.strip().str.lower()
    df_ban.columns = df_ban.columns.str.strip().str.lower()

    logger.info(f"DOI_SOAT | Kho: {len(df_kho):,} dòng | So_Ban: {len(df_ban):,} dòng")

    # Chuẩn hóa số điện thoại
    df_kho['_phone'] = df_kho['msisdn'].apply(clean_phone)
    df_ban['_phone'] = df_ban['stb'].apply(clean_phone)

    # Dedup So_Ban — giữ bản đầu tiên, tránh ghi đè sai
    before = len(df_ban)
    df_ban = df_ban.drop_duplicates(subset='_phone', keep='first')
    dup = before - len(df_ban)
    if dup:
        msg = f"So_Ban: bỏ {dup:,} dòng trùng số điện thoại"
        logger.warning(f"  ⚠ {msg}")
        errors.append(msg)

    # Lookup dicts
    idx      = df_ban.set_index('_phone')
    map_ngay = idx['ngày bán'].to_dict() if 'ngày bán' in idx.columns else {}
    map_lo   = idx['mã lô'].to_dict()    if 'mã lô'   in idx.columns else {}

    # Đảm bảo cột tồn tại trong kho trước khi fillna
    for col in ['ngày bán', 'mã lô bán', 'ghi chú']:
        if col not in df_kho.columns:
            df_kho[col] = None

    mapped_ngay = df_kho['_phone'].map(map_ngay)
    mapped_lo   = df_kho['_phone'].map(map_lo)

    # Chỉ ghi đè khi map có giá trị (không xoá dữ liệu cũ đã có)
    df_kho['ngày bán']  = mapped_ngay.where(mapped_ngay.notna(), df_kho['ngày bán'])
    df_kho['mã lô bán'] = mapped_lo.where(mapped_lo.notna(),     df_kho['mã lô bán'])

    matched   = int(mapped_ngay.notna().sum())
    unmatched = int(mapped_ngay.isna().sum())
    logger.info(f"DOI_SOAT | Khớp: {matched:,} | Chưa khớp: {unmatched:,}")

    df_kho = df_kho.drop(columns=['_phone'])
    df_kho.columns = df_kho.columns.str.title()

    config.ensure_dirs()
    df_kho.to_excel(config.file_doi_soat, index=False)

    msg = f"Hoàn thành: {matched:,} khớp / {len(df_kho):,} tổng"
    logger.info(f"DOI_SOAT | {msg}")
    return {
        "success": True, "count": len(df_kho),
        "matched": matched, "unmatched": unmatched,
        "output_file": config.file_doi_soat,
        "errors": errors, "message": msg,
    }


def _fail(msg: str, errors: list) -> dict:
    logger.error(f"DOI_SOAT | {msg}")
    return {"success": False, "count": 0, "matched": 0, "unmatched": 0,
            "output_file": None, "errors": errors, "message": msg}
