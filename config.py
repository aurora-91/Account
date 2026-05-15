"""
config.py — Toàn bộ cài đặt nằm ở đây.
Khi gắn vào web: tạo Config() rồi ghi đè các field trước khi truyền vào module.
"""
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class Config:
    # ── ĐƯỜNG DẪN ────────────────────────────────────────────────────────────
    thu_muc_kho: str      = r'/mnt/d/sim/ds_kho'
    thu_muc_bienban: str  = r'/mnt/d/sim/Biên bản đối soát'
    output_folder: str    = r'/mnt/d/sim/ket_qua'

    # ── TÊN FILE OUTPUT ───────────────────────────────────────────────────────
    ten_file_kho_sach: str = 'Kho_Sach.xlsx'
    ten_file_so_ban: str   = 'So_Ban.xlsx'
    ten_file_doi_soat: str = 'Kho_Sach_da_doi_soat.xlsx'

    # ── CẤU HÌNH KHO ─────────────────────────────────────────────────────────
    kho_sheet_chuan: str        = 'Kho'
    kho_sheet_du_phong: list    = field(default_factory=lambda: ['Kho thường', 'Kho định dạng'])
    kho_engine: str             = 'calamine'   # fallback tự động về openpyxl nếu chưa cài

    # ── CẤU HÌNH SỐ BÁN ──────────────────────────────────────────────────────
    # Từ khóa duy nhất, chữ thường, không phân biệt hoa/thường
    so_ban_tu_khoa_cot: str  = 'stb'
    so_ban_so_file_test: int = 0        # 0 = chạy toàn bộ

    # Từ khóa nhận diện sheet tổng hợp lô (thứ tự ưu tiên)
    lo_sheet_keywords: list = field(default_factory=lambda: [
        'tổng hợp lô', 'tong hop lo', 'danh sách lô', 'ds lô', 'lô', 'lo'
    ])
    # Từ khóa nhận diện sheet chi tiết
    chitiet_keywords: list = field(default_factory=lambda: ['chi tiết', 'chi tiet'])

    # ── ĐƯỜNG DẪN ĐẦY ĐỦ (computed) ─────────────────────────────────────────
    @property
    def file_kho_sach(self) -> str:
        return str(Path(self.thu_muc_kho) / self.ten_file_kho_sach)

    @property
    def file_so_ban(self) -> str:
        return str(Path(self.output_folder) / self.ten_file_so_ban)

    @property
    def file_doi_soat(self) -> str:
        return str(Path(self.output_folder) / self.ten_file_doi_soat)

    def ensure_dirs(self):
        Path(self.output_folder).mkdir(parents=True, exist_ok=True)
