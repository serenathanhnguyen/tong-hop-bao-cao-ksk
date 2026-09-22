"""Orchestrator: file M03 thô -> workbook "Bảng tổng hợp phân loại sức khỏe".

Đây là hàm chính app.py (Streamlit) gọi vào.
"""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl import Workbook

from . import columns as columns_mod
from . import raw_reader
from .classify import PersonDerived, derive_person
from .notes import build_ghi_chu_canh_bao
from .verify import VerifyIssue, verify_chem_formatting
from .workbook_builder import DATA_START_ROW, ReportMeta, build_column_defs, build_workbook


@dataclass
class PipelineResult:
    workbook: Workbook
    so_nguoi: int
    cot_da_bo: list[str]  # cận lâm sàng columns đã lược bỏ (QUY TẮC 1)
    verify_issues: list[VerifyIssue]
    canh_bao_global: list[str]


def run_pipeline(raw_path: str, meta: ReportMeta) -> PipelineResult:
    people = raw_reader.read_raw_workbook(raw_path)
    if not people:
        raise ValueError(
            "Không đọc được người nào có Họ tên trong sheet ThongTinHanhChinh — "
            "kiểm tra lại file đầu vào (đúng sheet, đúng vị trí bắt đầu row 5)."
        )

    derived: list[PersonDerived] = [derive_person(p) for p in people]

    active_cls_cols, dropped = columns_mod.active_cls_columns(people, derived)

    ghi_chu_list: list[str] = []
    canh_bao_list: list[str] = []
    for p, d in zip(people, derived):
        gc, cb = build_ghi_chu_canh_bao(p, d)
        ghi_chu_list.append(gc)
        canh_bao_list.append(cb)

    wb = build_workbook(meta, people, derived, active_cls_cols, ghi_chu_list, canh_bao_list)

    cols = build_column_defs(active_cls_cols)
    col_index = {c.key: i + 1 for i, c in enumerate(cols)}
    last_data_row = DATA_START_ROW + len(people) - 1
    issues = verify_chem_formatting(wb["Sheet1"], active_cls_cols, col_index, DATA_START_ROW, last_data_row)

    canh_bao_global = []
    if dropped:
        canh_bao_global.append(
            "Đã bỏ hẳn các cột cận lâm sàng không có dữ liệu ở bất kỳ ai: " + ", ".join(dropped)
        )
    if issues:
        canh_bao_global.append(
            f"Phát hiện {len(issues)} ô lệch định dạng in đậm/nghiêng sau khi rà soát — xem chi tiết."
        )

    return PipelineResult(
        workbook=wb,
        so_nguoi=len(people),
        cot_da_bo=dropped,
        verify_issues=issues,
        canh_bao_global=canh_bao_global,
    )
