"""
utils/date_extractor.py — Trích xuất ngày/tháng từ tên file, tên folder, hoặc nhập tay.

Thứ tự ưu tiên:
  1. manual_override (nhập tay)
  2. Tên file (bỏ phần mở rộng)
  3. Tên folder
  4. '' nếu không tìm được

Output format:
  "DD/MM/YYYY"  — khi có ngày đầy đủ
  "MM/YYYY"     — khi chỉ có tháng + năm
  "T{M}"        — khi chỉ có tháng (không năm)
"""
import re
from typing import Optional


# ── PATTERNS (thứ tự: cụ thể → chung) ───────────────────────────────────────
# Ghi chú kỹ thuật: không dùng \b trước/sau số vì dấu _ là word-char trong regex,
# dẫn đến "BB_T3_2025" không match \bT3\b. Dùng lookahead/lookbehind thay thế.

_PATTERNS: list[tuple[str, str]] = [
    # DD/MM/YYYY  hoặc  DD-MM-YYYY  (dấu phân cách bất kỳ)
    (r'(\d{1,2})[/\-_\.](\d{1,2})[/\-_\.](\d{4})', 'dmy'),

    # YYYYMMDD compact — phải test trước YYYYMM để tránh nuốt 2 số cuối
    (r'(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])(?!\d)', 'ymd8'),

    # YYYY/MM/DD  hoặc  YYYY-MM-DD
    (r'(20\d{2})[/\-_\.](0?[1-9]|1[0-2])[/\-_\.](0?[1-9]|[12]\d|3[01])', 'ymd'),

    # "Tháng 01 2025" / "thang 1 2025"
    (r'[Tt]h[áa]ng[_\s\-]*(0?[1-9]|1[0-2])[_\s\-/]*(20\d{2})', 'my'),

    # "T3_2025" / "T03-2025" — lookahead/lookbehind thay vì \b
    (r'(?<![A-Za-z\d])T(0?[1-9]|1[0-2])[_\-/\s](20\d{2})(?!\d)', 'my'),

    # "01-2025" / "01/2025" — không có chữ số liền trước
    (r'(?<!\d)(0?[1-9]|1[0-2])[/\-](20\d{2})(?!\d)', 'my'),

    # "202501" compact YYYYMM (chỉ khớp khi KHÔNG đi kèm 2 chữ số ngày ở sau)
    (r'(20\d{2})(0[1-9]|1[0-2])(?!\d)', 'ym_compact'),

    # "Tháng 3" (không kèm năm)
    (r'[Tt]h[áa]ng[_\s\-]*(0?[1-9]|1[0-2])(?!\d)', 'month_only'),

    # "T3" đứng độc lập — không kề chữ/số khác
    (r'(?<![A-Za-z\d])T(0?[1-9]|1[0-2])(?![A-Za-z\d])', 'month_only'),
]


# ── PUBLIC ────────────────────────────────────────────────────────────────────

def extract_date(
    filename: str = "",
    folder_name: str = "",
    manual_override: Optional[str] = None,
) -> str:
    """
    Trả về chuỗi ngày/tháng theo thứ tự ưu tiên.
    Trả về '' nếu không tìm được và không có manual_override.
    """
    if manual_override and manual_override.strip():
        return manual_override.strip()

    # Bỏ phần mở rộng file
    stem = re.sub(r'\.[A-Za-z]{2,5}$', '', filename)

    return _parse(stem) or _parse(folder_name) or ""


# ── PRIVATE ───────────────────────────────────────────────────────────────────

def _parse(text: str) -> str:
    if not text:
        return ""
    for pattern, fmt in _PATTERNS:
        m = re.search(pattern, text, re.UNICODE | re.IGNORECASE)
        if not m:
            continue
        g = m.groups()
        try:
            if fmt == 'dmy':
                d, mo, y = g
                return f"{int(d):02d}/{int(mo):02d}/{y}"
            elif fmt in ('ymd', 'ymd8'):
                y, mo, d = g
                return f"{int(d):02d}/{int(mo):02d}/{y}"
            elif fmt == 'my':
                mo, y = g
                return f"{int(mo):02d}/{y}"
            elif fmt == 'ym_compact':
                y, mo = g
                return f"{int(mo):02d}/{y}"
            elif fmt == 'month_only':
                return f"T{int(g[0])}"
        except (ValueError, IndexError):
            continue
    return ""
