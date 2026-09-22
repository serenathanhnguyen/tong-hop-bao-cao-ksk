"""QUY TẮC QUAN TRỌNG 1: quyết định cột cận lâm sàng nào được giữ lại trong
bảng tổng hợp — chỉ giữ cột có ít nhất 1 người có dữ liệu thực tế, xóa hẳn
cột nào trống ở TẤT CẢ mọi người."""

from __future__ import annotations

from . import config
from .classify import PersonDerived
from .raw_reader import Person


def _column_has_data(col: dict, people: list[Person], derived: list[PersonDerived]) -> bool:
    kind = col["kind"]
    if kind == "cbc":
        return any(d.cbc_result for d in derived)
    if kind == "chem":
        template_col = col["label"]
        return any(
            d.lab_flags.get(template_col) is not None and d.lab_flags[template_col].value is not None
            for d in derived
        )
    if kind == "urinalysis":
        return any(d.urinalysis_result for d in derived)
    if kind == "us_general":
        return any(p.us_general_text.strip() for p in people)
    if kind == "us_breast":
        return any(p.us_breast_text.strip() for p in people)
    return False


def active_cls_columns(people: list[Person], derived: list[PersonDerived]) -> list[dict]:
    """Trả về danh sách cột cận lâm sàng (thứ tự T..AE gốc) đã lược bỏ cột trống."""
    active = []
    dropped = []
    for col in config.FULL_CLS_COLUMNS:
        if _column_has_data(col, people, derived):
            active.append(col)
        else:
            dropped.append(col["label"])
    return active, dropped
