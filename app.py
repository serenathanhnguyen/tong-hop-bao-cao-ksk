"""Streamlit app: tổng hợp Bảng phân loại sức khỏe từ file dữ liệu khám sức
khỏe thô (Medinet M03, sheet ThongTinHanhChinh).

Chạy: streamlit run app.py
"""

from __future__ import annotations

import io
from datetime import date

import streamlit as st

from ksk_report.pipeline import run_pipeline
from ksk_report.workbook_builder import ReportMeta

st.set_page_config(page_title="Tổng hợp báo cáo KSK", page_icon="🩺", layout="centered")

st.title("🩺 Tổng hợp Bảng phân loại sức khỏe")
st.caption(
    "Upload file dữ liệu khám sức khỏe thô (Medinet, mẫu M03 — sheet "
    "`ThongTinHanhChinh`) để tạo Bảng tổng hợp phân loại sức khỏe theo đúng "
    "khung mẫu của Trạm Y tế. Toàn bộ quy tắc áp dụng nằm ở "
    "[`docs/QUY_TAC.md`](docs/QUY_TAC.md)."
)

with st.form("meta_form"):
    st.subheader("1. Thông tin đơn vị / đợt khám")
    col1, col2 = st.columns(2)
    with col1:
        ten_phuong = st.text_input("Tên phường/xã (dòng UBND)", value="PHƯỜNG NHIÊU LỘC")
        ten_tram_yte = st.text_input("Tên trạm y tế", value="TRẠM Y TẾ")
        nam = st.text_input("Năm khám", value=str(date.today().year))
    with col2:
        doi_tuong = st.text_input(
            "Đối tượng khám (in hoa trên tiêu đề bảng)",
            placeholder="VD: CB-NV CÔNG TY TNHH CÁ SẤU VIỆT PHONG",
        )
        nguoi_lap_bang = st.text_input("Người lập bảng", value="Nguyễn Thị Ngọc Sương")
        giam_doc = st.text_input("Tên Giám đốc (để trống nếu ký tay, không in tên)", value="")

    dien_ngay_ky = st.checkbox("Điền sẵn ngày ký cụ thể (mặc định: để trống dạng ‘Ngày … tháng … năm …’)")
    ngay_ky = None
    if dien_ngay_ky:
        ngay_ky = st.date_input("Ngày ký", value=date.today())

    st.subheader("2. File dữ liệu thô")
    uploaded = st.file_uploader("File Excel Medinet M03 (.xlsx)", type=["xlsx"])

    submitted = st.form_submit_button("Tạo bảng tổng hợp", type="primary")

if submitted:
    if uploaded is None:
        st.error("Vui lòng chọn file dữ liệu thô (.xlsx) trước.")
        st.stop()

    meta = ReportMeta(
        ten_phuong=ten_phuong.strip() or "PHƯỜNG NHIÊU LỘC",
        ten_tram_yte=ten_tram_yte.strip() or "TRẠM Y TẾ",
        doi_tuong=doi_tuong.strip(),
        nam=nam.strip(),
        nguoi_lap_bang=nguoi_lap_bang.strip() or "Nguyễn Thị Ngọc Sương",
        giam_doc=giam_doc.strip(),
        ngay_ky=ngay_ky,
    )

    tmp_path = "/tmp/_ksk_raw_upload.xlsx"
    with open(tmp_path, "wb") as f:
        f.write(uploaded.getbuffer())

    try:
        with st.spinner("Đang đọc dữ liệu và dựng bảng tổng hợp..."):
            result = run_pipeline(tmp_path, meta)
    except Exception as e:  # noqa: BLE001 — hiển thị lỗi thân thiện cho người dùng cuối
        st.error(f"Không tạo được bảng tổng hợp: {e}")
        st.stop()

    st.success(f"Đã tổng hợp xong cho {result.so_nguoi} người.")

    if result.cot_da_bo:
        st.info(
            "Đã bỏ hẳn các cột cận lâm sàng không có dữ liệu ở bất kỳ ai (QUY TẮC 1): "
            + ", ".join(result.cot_da_bo)
        )

    if result.verify_issues:
        st.warning(
            f"⚠️ Phát hiện {len(result.verify_issues)} ô lệch định dạng in đậm/nghiêng — "
            "đã được sửa lại đúng trong file, nhưng nên xem lại danh sách bên dưới."
        )
        with st.expander("Chi tiết ô lệch định dạng"):
            for issue in result.verify_issues:
                st.write(f"`{issue.cell}` — {issue.message}")
    else:
        st.caption("✅ Đã kiểm tra chéo: không có ô nào lệch định dạng in đậm/in nghiêng so với bảng tham chiếu.")

    st.warning(
        "Cột **CẢNH BÁO** ở cuối bảng liệt kê mọi nhận định do công cụ tự suy ra "
        "(Xếp loại, bt/x của CTM/nước tiểu/siêu âm, chẩn đoán sơ bộ chưa xác định...). "
        "Rà soát lại rồi mới xóa cột này trước khi phát hành bảng chính thức."
    )

    buf = io.BytesIO()
    result.workbook.save(buf)
    buf.seek(0)

    file_name = f"BTH_KSK_{meta.nam or 'nam'}_{(meta.doi_tuong or 'doi_tuong').replace(' ', '_')[:40]}.xlsx"
    st.download_button(
        "⬇️ Tải file Bảng tổng hợp (.xlsx)",
        data=buf,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
