"""Đọc file dữ liệu thô Medinet M03 (sheet ThongTinHanhChinh) thành danh sách Person."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet

from . import config


def _col_idx(letter: str) -> int:
    return column_index_from_string(letter)


def _cell(ws: Worksheet, row: int, col_letter: str) -> Any:
    return ws.cell(row=row, column=_col_idx(col_letter)).value


def parse_vn_number(value: Any) -> float | None:
    """Đổi số dạng Việt Nam (dấu phẩy thập phân, vd '4,71') thành float.

    Chấp nhận cả số/float sẵn có (Excel đôi khi đã parse sẵn), chuỗi có dấu
    chấm phân cách hàng nghìn, hoặc rỗng -> None.
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s:
        return None
    # loại bỏ khoảng trắng, đơn vị dính kèm (vd "68 U/L")
    s = s.replace(" ", "")
    # nếu có cả dấu chấm và dấu phẩy -> dấu chấm là phân cách nghìn, dấu phẩy là thập phân
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


@dataclass
class ClinicalResult:
    chua_phat_hien_bt: str = ""
    chan_doan_so_bo: str = ""
    chan_doan_xac_dinh: str = ""
    phan_loai: int | None = None
    tu_choi_kham: str = ""


@dataclass
class Person:
    row: int
    cccd: str = ""
    ho_ten: str = ""
    ngay_sinh: Any = None
    gioi_tinh_raw: Any = None
    chieu_cao: float | None = None
    can_nang: float | None = None
    ha_tam_thu: Any = None
    ha_tam_truong: Any = None
    mat_phai: Any = None
    mat_trai: Any = None
    clinical: dict[str, ClinicalResult] = field(default_factory=dict)
    labs: dict[str, Any] = field(default_factory=dict)
    us_general_text: str = ""
    xquang_text: str = ""
    us_breast_text: str = ""
    ket_luan_chung: str = ""

    @property
    def gioi_tinh(self) -> str:
        g = parse_vn_number(self.gioi_tinh_raw)
        if g == 1:
            return "Nam"
        if g == 2:
            return "Nữ"
        return ""

    @property
    def is_nam(self) -> bool:
        return self.gioi_tinh == "Nam"

    @property
    def huyet_ap_text(self) -> str:
        thu = parse_vn_number(self.ha_tam_thu)
        truong = parse_vn_number(self.ha_tam_truong)
        if thu is None or truong is None:
            return ""
        return f"{int(thu)}/{int(truong)}"


def _read_clinical_group(ws: Worksheet, row: int, group: config.ClinicalGroup) -> ClinicalResult:
    start = _col_idx(group.start_col)
    vals = [ws.cell(row=row, column=start + i).value for i in range(group.ncols)]
    result = ClinicalResult(
        chua_phat_hien_bt=_text(vals[0]) if len(vals) > 0 else "",
        chan_doan_so_bo=_text(vals[1]) if len(vals) > 1 else "",
        chan_doan_xac_dinh=_text(vals[2]) if len(vals) > 2 else "",
    )
    phan_loai_raw = vals[3] if len(vals) > 3 else None
    pl = parse_vn_number(phan_loai_raw)
    result.phan_loai = int(pl) if pl is not None else None
    if group.has_tu_choi_kham and len(vals) > 4:
        result.tu_choi_kham = _text(vals[4])
    return result


def read_raw_workbook(path: str) -> list[Person]:
    """Đọc toàn bộ dòng có dữ liệu (có Họ tên) trong sheet ThongTinHanhChinh.

    Đọc đến hết sheet thực tế (ws.max_row), KHÔNG dừng sớm khi gặp 1-2 dòng
    trống liên tiếp ở giữa (đã có lần bỏ sót người thật vì dừng đọc sớm) —
    chỉ dừng hẳn khi không còn dòng nào có tên phía sau.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    if config.RAW_SHEET_NAME not in wb.sheetnames:
        raise ValueError(
            f"Không tìm thấy sheet '{config.RAW_SHEET_NAME}' trong file. "
            f"Các sheet có sẵn: {wb.sheetnames}"
        )
    ws = wb[config.RAW_SHEET_NAME]

    # Tìm dòng cuối cùng thực sự có tên (quét toàn bộ, không giả định số người)
    last_row_with_name = None
    for r in range(config.RAW_DATA_START_ROW, ws.max_row + 1):
        if _text(_cell(ws, r, config.COL_HO_TEN)):
            last_row_with_name = r
    if last_row_with_name is None:
        return []

    people: list[Person] = []
    for r in range(config.RAW_DATA_START_ROW, last_row_with_name + 1):
        ho_ten = _text(_cell(ws, r, config.COL_HO_TEN))
        if not ho_ten:
            # dòng trống xen giữa -> bỏ qua nhưng KHÔNG dừng đọc (xem docstring)
            continue

        p = Person(
            row=r,
            cccd=_text(_cell(ws, r, config.COL_CCCD)),
            ho_ten=ho_ten,
            ngay_sinh=_cell(ws, r, config.COL_NGAY_SINH),
            gioi_tinh_raw=_cell(ws, r, config.COL_GIOI_TINH),
            chieu_cao=parse_vn_number(_cell(ws, r, config.COL_CHIEU_CAO)),
            can_nang=parse_vn_number(_cell(ws, r, config.COL_CAN_NANG)),
            ha_tam_thu=_cell(ws, r, config.COL_HA_TAM_THU),
            ha_tam_truong=_cell(ws, r, config.COL_HA_TAM_TRUONG),
            mat_phai=_cell(ws, r, config.COL_MAT_KHONG_KINH_PHAI),
            mat_trai=_cell(ws, r, config.COL_MAT_KHONG_KINH_TRAI),
            us_general_text=_text(_cell(ws, r, config.COL_SIEU_AM_BUNG)),
            xquang_text=_text(_cell(ws, r, config.COL_XQUANG)),
            ket_luan_chung=_text(_cell(ws, r, config.COL_KET_LUAN_CHUNG)),
        )
        # siêu âm tuyến vú: 2 cột (X-quang nhũ + siêu âm 2 tuyến vú) -> gộp text
        xquang_nhu = _text(_cell(ws, r, config.COL_XQUANG_NHU))
        sa_tuyen_vu = _text(_cell(ws, r, config.COL_SIEU_AM_TUYEN_VU))
        p.us_breast_text = " | ".join(t for t in [xquang_nhu, sa_tuyen_vu] if t)

        for group in config.CLINICAL_GROUPS:
            p.clinical[group.key] = _read_clinical_group(ws, r, group)

        for lab_key, col in {
            "hc": config.COL_HC_SL,
            "hgb": config.COL_HGB,
            "hct": config.COL_HCT,
            "mcv": config.COL_MCV,
            "mch": config.COL_MCH,
            "mchc": config.COL_MCHC,
            "rdw": config.COL_RDW,
            "bach_cau": config.COL_BACH_CAU,
            "tieu_cau": config.COL_TIEU_CAU,
            "glucose": config.COL_GLUCOSE,
            "ure": config.COL_URE,
            "creatinin": config.COL_CREATININ,
            "ast": config.COL_AST,
            "alt": config.COL_ALT,
            "ti_trong_nt": config.COL_TI_TRONG_NUOC_TIEU,
            "ph_nt": config.COL_PH_NUOC_TIEU,
            "acid_uric": config.COL_ACID_URIC,
            "cholesterol": config.COL_CHOLESTEROL,
            "triglycerid": config.COL_TRIGLYCERID,
            "hdl": config.COL_HDL,
            "ldl": config.COL_LDL,
        }.items():
            p.labs[lab_key] = parse_vn_number(_cell(ws, r, col))

        people.append(p)

    return people
