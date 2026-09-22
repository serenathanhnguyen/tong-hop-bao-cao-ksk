"""Dựng nội dung 2 cột "Ghi chú" (chỉ chép nguyên văn dữ liệu gốc) và
"Cảnh báo" (mọi nhận định/suy luận của công cụ, để người lập bảng rà soát
lại rồi xóa cột này đi) — xem QUY TẮC QUAN TRỌNG 2 trong docs/QUY_TAC.md.
"""

from __future__ import annotations

from . import config
from .classify import PersonDerived
from .raw_reader import Person


def _clinical_group_notes(p: Person) -> tuple[list[str], list[str]]:
    """Trả về (ghi_chu_items, canh_bao_items) cho các mã chẩn đoán lâm sàng."""
    ghi_chu: list[str] = []
    canh_bao: list[str] = []
    for group in config.CLINICAL_GROUPS:
        r = p.clinical.get(group.key)
        if r is None:
            continue
        parts = []
        if r.chan_doan_so_bo:
            parts.append(f"chẩn đoán sơ bộ '{r.chan_doan_so_bo}'")
        if r.chan_doan_xac_dinh:
            parts.append(f"chẩn đoán xác định '{r.chan_doan_xac_dinh}'")
        if r.tu_choi_kham:
            parts.append(f"từ chối khám ({r.tu_choi_kham})")
        if parts:
            ghi_chu.append(f"{group.label}: " + "; ".join(parts))
        if r.chan_doan_so_bo and not r.chan_doan_xac_dinh:
            canh_bao.append(
                f"{group.label}: mới chỉ có chẩn đoán sơ bộ '{r.chan_doan_so_bo}', "
                f"chưa xác định — kiểm tra lại."
            )
    return ghi_chu, canh_bao


def build_ghi_chu_canh_bao(p: Person, d: PersonDerived) -> tuple[str, str]:
    ghi_chu_items: list[str] = []
    canh_bao_items: list[str] = []

    clinical_notes, clinical_warnings = _clinical_group_notes(p)
    ghi_chu_items.extend(clinical_notes)
    canh_bao_items.extend(clinical_warnings)

    if p.us_general_text.strip():
        ghi_chu_items.append(f"Siêu âm tổng quát: {p.us_general_text.strip()}")
    if p.xquang_text.strip():
        ghi_chu_items.append(f"X-quang: {p.xquang_text.strip()}")
    if p.us_breast_text.strip():
        ghi_chu_items.append(f"Siêu âm tuyến vú: {p.us_breast_text.strip()}")
    if p.ket_luan_chung.strip():
        ghi_chu_items.append(f"Kết luận: {p.ket_luan_chung.strip()}")

    # Cảnh báo: Xếp loại tự suy ra
    if d.xep_loai_note:
        canh_bao_items.append(d.xep_loai_note)

    # Cảnh báo: các chỉ số máu in đậm/in nghiêng
    for template_col, flag in d.lab_flags.items():
        if flag.out_of_range_label:
            style = "in đậm (cao)" if flag.bold else "in nghiêng (thấp)"
            canh_bao_items.append(f"{template_col}: {flag.out_of_range_label} → {style}, xác nhận lại định dạng.")

    if d.cbc_out_of_range:
        canh_bao_items.append(
            "Tổng phân tích tế bào máu = 'x' (suy luận từ bảng tham chiếu 7b): "
            + "; ".join(d.cbc_out_of_range)
        )
    if d.urinalysis_out_of_range:
        canh_bao_items.append(
            "Tổng phân tích nước tiểu = 'x' (suy luận từ bảng tham chiếu 7): "
            + "; ".join(d.urinalysis_out_of_range)
        )
    if d.us_general_inferred:
        canh_bao_items.append(
            f"Siêu âm tổng quát = '{d.us_general_result}' — suy luận tự động từ mô tả văn bản, kiểm tra lại."
        )
    if d.us_breast_inferred:
        canh_bao_items.append(
            f"Siêu âm tuyến vú = '{d.us_breast_result}' — suy luận tự động từ mô tả văn bản, kiểm tra lại."
        )

    return "\n".join(ghi_chu_items), "\n".join(canh_bao_items)
