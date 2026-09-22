"""Bước 8 trong quy trình: kiểm tra chéo toàn bộ ô số cận lâm sàng trước khi
gửi file — so giá trị với bảng tham chiếu và đối chiếu với định dạng
bold/italic THỰC TẾ của ô trong workbook đã dựng xong (không tin vào cờ đã
tính lúc build), để bắt các trường hợp lệch định dạng trước khi gửi cho Jo.
"""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl import Workbook

from . import config


@dataclass
class VerifyIssue:
    cell: str
    message: str


def verify_chem_formatting(
    ws,
    active_cls_cols: list[dict],
    col_index: dict[str, int],
    data_start_row: int,
    last_data_row: int,
) -> list[VerifyIssue]:
    issues: list[VerifyIssue] = []
    if last_data_row < data_start_row:
        return issues

    chem_cols = [c for c in active_cls_cols if c["kind"] == "chem"]
    for c in chem_cols:
        rng = config.BLOOD_CHEM_RANGES[c["chem_name"]]
        idx = col_index[f"cls_{c['key']}"]
        for row in range(data_start_row, last_data_row + 1):
            cell = ws.cell(row=row, column=idx)
            value = cell.value
            if not isinstance(value, (int, float)):
                continue
            expect_bold = rng.high is not None and value > rng.high
            expect_italic = rng.low is not None and value < rng.low
            actual_bold = bool(cell.font.bold)
            actual_italic = bool(cell.font.italic)
            if actual_bold != expect_bold or actual_italic != expect_italic:
                issues.append(
                    VerifyIssue(
                        cell=cell.coordinate,
                        message=(
                            f"{c['label']} = {value} — kỳ vọng bold={expect_bold}, "
                            f"italic={expect_italic} nhưng ô đang bold={actual_bold}, "
                            f"italic={actual_italic}."
                        ),
                    )
                )
    return issues
