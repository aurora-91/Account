"""
utils/phone.py — Chuẩn hóa số điện thoại Việt Nam về dạng 10 chữ số.
"""
import re

import pandas as pd


def clean_phone(value) -> str:
    """
    Chuẩn hóa số điện thoại. Trả về chuỗi 10 chữ số hoặc '' nếu không hợp lệ.

    Các dạng đầu vào được hỗ trợ:
        "0901234567"         → "0901234567"
        "0901234567.0"       → "0901234567"   (float từ Excel)
        "84901234567"        → "0901234567"   (prefix quốc tế)
        "+84 901 234 567"    → "0901234567"   (prefix + khoảng trắng)
        "901234567"          → "0901234567"   (9 số, format cũ)
        "090-123-4567"       → "0901234567"   (gạch ngang)
    """
    if pd.isna(value):
        return ""

    s = str(value).strip()

    # Xử lý float từ Excel: "0901234567.0" → "0901234567"
    if s.endswith('.0'):
        s = s[:-2]

    # Chỉ giữ chữ số
    s = re.sub(r'\D', '', s)
    if not s:
        return ""

    # Prefix quốc tế: +840... (12 số) hoặc 84... (11 số)
    if s.startswith('840') and len(s) == 12:
        s = '0' + s[3:]
    elif s.startswith('84') and len(s) == 11:
        s = '0' + s[2:]

    # Format 9 số cũ → thêm 0 đầu
    if len(s) == 9:
        s = '0' + s

    return s if len(s) == 10 else ""
