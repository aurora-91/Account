"""
web_app.py — SIM Tool · Không lưu trữ, không database.

Dùng: upload file → chạy → xem → tải xuống. Đóng tab là hết.

Chạy: streamlit run web_app.py
"""
import io
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))
from config import Config
from modules import kho as mod_kho
from modules import so_ban as mod_so_ban
from modules.doi_soat import run_from_df
from utils.date_extractor import extract_date

st.set_page_config(page_title="SIM Tool", page_icon="📱", layout="wide")

st.markdown("""
<style>
#MainMenu, footer { visibility: hidden }
[data-testid="metric-container"] {
    border: 1px solid rgba(128,128,128,.15);
    border-radius: 10px;
    padding: 12px 16px;
}
</style>
""", unsafe_allow_html=True)


# ── HELPERS ───────────────────────────────────────────────────────────────────

def to_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def show_errors(errors: list, label="cảnh báo"):
    if errors:
        with st.expander(f"⚠ {len(errors)} {label}"):
            for e in errors:
                st.warning(e)

# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    st.title("📱 SIM Tool")
    st.caption("Upload file kho + biên bản → xem kết quả → tải xuống. Đóng tab là sạch.")

    # ── Upload ────────────────────────────────────────────────────────────────
    col_kho, col_bb = st.columns(2)

    with col_kho:
        st.subheader("📦 File kho")
        message_kho = """
        Định dạng file kho hợp lệ:
        - File Excel .xlsx
        - Có cột MSISDN bắt buộc
        - Header nằm trong 8 dòng đầu
        - Có thể thêm các cột:
        Hạng số, Trạng thái, Ngày bán, Mã lô bán, Ghi chú
        - Mỗi file = 1 kho
        - Tên file sẽ là tên kho nếu không nhập tay
        """
        st.info(message_kho)
        files_kho = st.file_uploader(
            "Chọn file kho (.xlsx)",
            type=["xlsx"],
            accept_multiple_files=True,
            key="up_kho",
        )

        if files_kho:
            st.session_state["files_kho"] = files_kho
        else:
            files_kho = st.session_state.get("files_kho", [])
        names: dict = {}
        if files_kho:
            st.markdown("**Tên kho** · để trống = lấy tên file:")
            for f in files_kho:
                names[f.name] = st.text_input(
                    f.name,
                    placeholder=Path(f.name).stem,
                    key=f"kn_{f.name}",
                    label_visibility="collapsed",
                )

    with col_bb:
        st.subheader("📋 File biên bản")
        message_bb = """
        Định dạng file biên bản hợp lệ:
        - File Excel .xlsx
        - Nên có sheet "Chi tiết" (hoặc tên gần giống)
        - Có cột STB chứa số điện thoại
        - Header nằm trong 10 dòng đầu
        - Có thể có nhiều sheet lô, nếu muốn đọc STB trong sheet lô cần tạo 1 cột Mã lô có tên tương ứng
        - Tên file hoặc thư mục nên chứa tháng/năm (vd: 05_2025)
        - Nếu không nhận diện được ngày, nhập tay bên dưới
        """
        st.info(message_bb)
        files_bb = st.file_uploader(
            "Chọn file biên bản (.xlsx)",
            type=["xlsx"],
            accept_multiple_files=True,
            key="up_bb",
        )

        if files_bb:
            st.session_state["files_bb"] = files_bb
        else:
            files_bb = st.session_state.get("files_bb", [])
        dates: dict = {}
        missing_dates: list = []
        if files_bb:
            st.markdown("**Ngày/tháng** · tự detect, có thể sửa:")
            for f in files_bb:
                detected = extract_date(filename=f.name)
                c1, c2 = st.columns([3, 2])
                c1.markdown(f"`{f.name}`")
                val = c2.text_input(
                    f.name,
                    value=detected,
                    placeholder="vd: 05/2025",
                    key=f"bd_{f.name}",
                    label_visibility="collapsed",
                )
                dates[f.name] = val.strip()
                if not dates[f.name]:
                    missing_dates.append(f.name)

            if missing_dates:
                st.warning(f"⚠ Chưa có ngày: {', '.join(missing_dates)}")

    # ── Nút chạy ──────────────────────────────────────────────────────────────
    st.divider()
    can_run = bool(files_kho) and bool(files_bb) and not missing_dates

    if st.button("▶  Chạy đối soát", type="primary", disabled=not can_run):
        _chay(files_kho, names, files_bb, dates)

    if not can_run and (files_kho or files_bb):
        if missing_dates:
            st.caption("Điền ngày/tháng cho tất cả file biên bản trước khi chạy.")
        else:
            st.caption("Cần upload đủ cả file kho lẫn biên bản.")

    # ── Kết quả ───────────────────────────────────────────────────────────────
    if "result" in st.session_state:
        _show_result()


# ── XỬ LÝ ────────────────────────────────────────────────────────────────────

def _chay(files_kho, names, files_bb, dates):
    with st.status("Đang xử lý...", expanded=True) as status:
        status.write(f"Đọc {len(files_kho)} file kho...")
        r_kho = mod_kho.process_files(files_kho, names)
        if not r_kho["success"]:
            status.update(label="Lỗi đọc kho", state="error")
            st.error(r_kho["message"])
            return
        
        status.write(f"Đọc {len(files_bb)} file biên bản...")
        r_bb = mod_so_ban.process_files(files_bb, dates)
        if not r_bb["success"]:
            status.update(label="Lỗi đọc biên bản", state="error")
            st.error(r_bb["message"])
            return

        status.write("Ghép kho × biên bản...")
        result = run_from_df(r_kho["df"], r_bb["df"])
        if not result["success"]:
            status.update(label="Lỗi đối soát", state="error")
            st.error(result["message"])
            return
        
        status.update(label="✅ Hoàn thành!", state="complete")

    # LƯU DỮ LIỆU VÀO SESSION ĐỂ TẢI XUỐNG
    st.session_state["result"] = result
    # Lưu file đối soát tổng
    st.session_state["excel_all"] = to_excel(result["df_result"].drop(columns=["Phân Loại"], errors="ignore"))
    # Lưu file Kho sạch (đã gộp và lọc trùng)
    st.session_state["excel_kho_clean"] = to_excel(r_kho["df"])
    # Lưu file Bán sạch (đã gộp và lọc trùng)
    st.session_state["excel_ban_clean"] = to_excel(r_bb["df"])
    
    st.rerun()


# ── HIỂN THỊ KẾT QUẢ ─────────────────────────────────────────────────────────

def _show_result():
    result = st.session_state["result"]
    df: pd.DataFrame = result["df_result"]

    st.divider()
    st.subheader("📊 Kết quả đối soát")

    # 1. Chỉ số Metrics
    c1, c2, c3, c4 = st.columns(4)
    tong = result["matched"] + result["unmatched"]
    c1.metric("Tổng trong kho", f"{tong:,}")
    c2.metric("Khớp Kho & Bán", f"{result['matched']:,}")
    c3.metric("Tồn Kho (Chưa bán)", f"{result['unmatched']:,}")
    c4.metric("Lỗi ngoài kho", f"{result['ngoai_kho']:,}")

    dup_kho_count = 0
    dup_ban_count = 0

    if "Ghi Chú Kho" in df.columns:
        dup_kho_count = df["Ghi Chú Kho"].fillna("").ne("").sum()

    if "Ghi Chú Bán" in df.columns:
        dup_ban_count = df["Ghi Chú Bán"].fillna("").ne("").sum()

    e1, e2 = st.columns(2)

    e1.error(f"⚠ Trùng trong kho: {dup_kho_count:,} số")
    e2.warning(f"⚠ Trùng trong số bán: {dup_ban_count:,} số")

    # 2. Bộ lọc (Filters)
    f1, f2, f3 = st.columns(3)
    kho_opts = sorted(df["Tên Kho"].dropna().unique()) if "Tên Kho" in df.columns else []
    ngay_opts = sorted(df["Ngày Bán"].dropna().unique()) if "Ngày Bán" in df.columns else []

    f_kho = f1.multiselect("Lọc Kho", kho_opts)
    f_ngay = f2.multiselect("Lọc Tháng bán", ngay_opts)
    
    # THÊM BỘ LỌC TRÙNG Ở ĐÂY
    is_filter_dup = st.checkbox("🔍 Chỉ hiện các số bị trùng (Kho hoặc Bán)", value=False)

    dv = df.copy()
    if f_kho: dv = dv[dv["Tên Kho"].isin(f_kho)]
    if f_ngay: dv = dv[dv["Ngày Bán"].isin(f_ngay)]
    
    if is_filter_dup:
        cond_kho = (
            dv["Ghi Chú Kho"].fillna("").ne("")
            if "Ghi Chú Kho" in dv.columns
            else False
        )

        cond_ban = (
            dv["Ghi Chú Bán"].fillna("").ne("")
            if "Ghi Chú Bán" in dv.columns
            else False
        )

        dv = dv[cond_kho | cond_ban]
    
    # Logic lọc trùng
    if is_filter_dup:
        cond_kho = (dv["Ghi Chú Kho"] != "") & (dv["Ghi Chú Kho"].notna())
        cond_ban = (dv["Ghi Chú Bán"] != "") & (dv["Ghi Chú Bán"].notna())
        dv = dv[cond_kho | cond_ban]

    # 3. HÀM TÔ MÀU (Highlight)
    def style_row(row):
        styles = [''] * len(row)
        for i, col in enumerate(row.index):
            if col == "Ghi Chú Kho" and row[col]:
                styles[i] = 'background-color: #FFCC80; color: black' # Cam
            elif col == "Ghi Chú Bán" and row[col]:
                styles[i] = 'background-color: #B2EBF2; color: black' # Xanh
        return styles

    # 4. Hiển thị bảng
    show_cols = [c for c in dv.columns if c not in ("Phân Loại", "STT")]
    view_df = dv[show_cols]

    # Chỉ tô màu khi dữ liệu nhỏ
    if len(view_df) <= 5000:
        st.dataframe(
            view_df.style.apply(style_row, axis=1),
            use_container_width=True,
            height=500,
            hide_index=True
        )
    else:
        st.info("Dữ liệu lớn → tắt tô màu để tăng tốc hiển thị.")

        st.dataframe(
            view_df,
            use_container_width=True,
            height=500,
            hide_index=True
        )
    st.caption(f"Đang hiển thị: {len(dv):,} / {len(df):,} dòng")

    # 5. NÚT TẢI FILE (Nằm kế bên nhau)
    st.markdown("### 📥 Tải xuống dữ liệu sạch")
    d1, d2, d3, _ = st.columns([1, 1, 1, 2])
    
    d1.download_button(
        "⬇️ Tải Đối Soát Tổng",
        st.session_state["excel_all"],
        "doi_soat_day_du.xlsx",
        help="File kết quả đối soát cuối cùng"
    )
    
    d2.download_button(
        "⬇️ Tải Kho Tổng Hợp",
        st.session_state["excel_kho_clean"],
        "kho_tong_hop_sach.xlsx",
        help="File gộp tất cả kho đã xử lý trùng"
    )
    
    d3.download_button(
        "⬇️ Tải Số Bán Tổng Hợp",
        st.session_state["excel_ban_clean"],
        "so_ban_tong_hop_sach.xlsx",
        help="File gộp tất cả biên bản đã xử lý trùng"
    )


if __name__ == "__main__":
    main()
