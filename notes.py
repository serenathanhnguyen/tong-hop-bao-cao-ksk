"""Dựng nội dung 2 cột "Ghi chú" (chỉ chép nguyên văn dữ liệu gốc) và
"Cảnh báo" (mọi nhận định/suy luận của công cụ, để người lập bảng rà soát
lại rồi xóa cột này đi) — xem QUY TẮC QUAN TRỌNG 2 trong docs/QUY_TAC.md.

Văn phong câu chữ (nối câu bằng dấu chấm + khoảng trắng, không xuống dòng;
gộp nhiều chỉ số cận lâm sàng lệch cùng chiều vào 1 câu; nhãn "X-quang phổi:"
...) khớp đúng theo file khung mẫu thật Jo cung cấp
(BTH_KSK_2026_CB-NV_CONG_TY_CA_SAU_VIET_PHONG), không phải suy đoán.
"""

from __future__ import annotations

from . import config
from .classify import PersonDerived
from .raw_reader import Person


def _fmt_num(v: float) -> str:
    if v == int(v):
        return str(int(v))
    return f"{v:g}"


def _join_vn(items: list[str]) -> str:
    """Nối danh sách kiểu tiếng Việt: 'A, B và C' (khớp 'ALT (46) và AST (44)')."""
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    return ", ".join(items[:-1]) + " và " + items[-1]


def _clinical_group_notes(p: Person) -> tuple[list[str], list[str]]:
    """Trả về (ghi_chu_items, canh_bao_items) cho các mã chẩn đoán lâm sàng.

    Quy tắc (theo khung mẫu thật): nếu có "chẩn đoán xác định" -> ghi thẳng
    "{Chuyên khoa}: {mã}." không kèm nhãn phụ. Chỉ khi CHƯA có xác định mà chỉ
    có "chẩn đoán sơ bộ" thì mới gắn nhãn "(chẩn đoán sơ bộ)" sau tên chuyên
    khoa, và luôn kèm cảnh báo "... dùng mã chẩn đoán SƠ BỘ (chưa có chẩn đoán
    xác định) — kiểm tra lại."
    """
    ghi_chu: list[str] = []
    canh_bao: list[str] = []
    for group in config.CLINICAL_GROUPS:
        r = p.clinical.get(group.key)
        if r is None:
            continue
        if r.chan_doan_xac_dinh:
            ghi_chu.append(f"{group.label}: {r.chan_doan_xac_dinh}.")
            if r.chan_doan_so_bo and r.chan_doan_so_bo != r.chan_doan_xac_dinh:
                canh_bao.append(
                    f"{group.label}: còn có mã chẩn đoán sơ bộ '{r.chan_doan_so_bo}' "
                    f"khác với chẩn đoán xác định — kiểm tra lại."
                )
        elif r.chan_doan_so_bo:
            ghi_chu.append(f"{group.label} (chẩn đoán sơ bộ): {r.chan_doan_so_bo}.")
            canh_bao.append(
                f"{group.label} dùng mã chẩn đoán SƠ BỘ (chưa có chẩn đoán xác định) — kiểm tra lại."
            )
        if r.tu_choi_kham:
            ghi_chu.append(f"{group.label}: Từ chối khám ({r.tu_choi_kham}).")
    return ghi_chu, canh_bao


def build_ghi_chu_canh_bao(p: Person, d: PersonDerived) -> tuple[str, str]:
    ghi_chu_items: list[str] = []
    canh_bao_items: list[str] = []

    clinical_notes, clinical_warnings = _clinical_group_notes(p)
    ghi_chu_items.extend(clinical_notes)

    if p.us_general_text.strip():
        ghi_chu_items.append(f"Siêu âm tổng quát: {p.us_general_text.strip().rstrip('.')}.")
    if p.xquang_text.strip():
        ghi_chu_items.append(f"X-quang phổi: {p.xquang_text.strip().rstrip('.')}.")
    if p.us_breast_text.strip():
        ghi_chu_items.append(f"Siêu âm tuyến vú: {p.us_breast_text.strip().rstrip('.')}.")
    if p.ket_luan_chung.strip():
        ghi_chu_items.append(f"Kết luận: {p.ket_luan_chung.strip().rstrip('.')}.")

    # Cảnh báo: khớp đúng thứ tự trong file khung mẫu thật — Xếp loại tự suy
    # ra LUÔN đứng đầu, rồi mới đến các cảnh báo chẩn đoán sơ bộ/chỉ số lệch.
    if d.xep_loai_note:
        canh_bao_items.append(d.xep_loai_note)
    canh_bao_items.extend(clinical_warnings)

    # Cảnh báo: gộp các chỉ số máu in đậm (cao) / in nghiêng (thấp) mỗi loại thành 1 câu
    bold_items = []
    italic_items = []
    for template_col, flag in d.lab_flags.items():
        if flag.value is None:
            continue
        short = config.CHEM_SHORT_NAME.get(template_col, template_col)
        if flag.bold:
            bold_items.append(f"{short} ({_fmt_num(flag.value)})")
        elif flag.italic:
            italic_items.append(f"{short} ({_fmt_num(flag.value)})")
    if bold_items:
        canh_bao_items.append(
            f"{_join_vn(bold_items)} in đậm do vượt giá trị tham chiếu — kiểm tra lại."
        )
    if italic_items:
        canh_bao_items.append(
            f"{_join_vn(italic_items)} in nghiêng do thấp hơn giá trị tham chiếu — kiểm tra lại."
        )

    if d.cbc_out_of_range:
        canh_bao_items.append(
            f"Tổng phân tích tế bào máu = 'x' (suy luận từ bảng tham chiếu 7b): "
            + "; ".join(d.cbc_out_of_range) + " — kiểm tra lại."
        )
    if d.urinalysis_out_of_range:
        canh_bao_items.append(
            f"Tổng phân tích nước tiểu = 'x' (suy luận từ bảng tham chiếu 7): "
            + "; ".join(d.urinalysis_out_of_range) + " — kiểm tra lại."
        )
    if d.us_general_inferred:
        canh_bao_items.append(
            f"Siêu âm tổng quát = '{d.us_general_result}' — suy luận tự động từ mô tả văn bản, kiểm tra lại."
        )
    if d.us_breast_inferred:
        canh_bao_items.append(
            f"Siêu âm tuyến vú = '{d.us_breast_result}' — suy luận tự động từ mô tả văn bản, kiểm tra lại."
        )

    return " ".join(ghi_chu_items), " ".join(canh_bao_items)
