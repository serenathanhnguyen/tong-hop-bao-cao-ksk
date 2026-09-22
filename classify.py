"""Suy ra các giá trị dẫn xuất: phân loại lâm sàng theo cột khung mẫu, Xếp loại,
so sánh với bảng tham chiếu (in đậm/in nghiêng), và kết luận bt/x cho CTM /
nước tiểu / siêu âm.

QUAN TRỌNG: mọi giá trị suy ra trong module này (Xếp loại, bt/x của CTM/nước
tiểu/siêu âm) là suy luận của công cụ, KHÔNG có sẵn trong dữ liệu gốc dưới
dạng đã phân loại — theo đúng quy tắc trong docs/QUY_TAC.md, các suy luận này
phải được ghi vào cột "Cảnh báo" của bảng tổng hợp để người lập bảng rà soát
lại, không phải là sự thật tuyệt đối.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from . import config
from .raw_reader import Person


@dataclass
class LabFlag:
    value: float | None
    bold: bool = False  # cao hơn cận trên
    italic: bool = False  # thấp hơn cận dưới
    out_of_range_label: str | None = None  # vd "ALT 68 U/L (> 41)"


def flag_value(name: str, value: float | None, rng: config.RefRange) -> LabFlag:
    if value is None:
        return LabFlag(value=None)
    bold = rng.high is not None and value > rng.high
    italic = rng.low is not None and value < rng.low
    label = None
    if bold:
        label = f"{name} = {value:g} {rng.unit} (> {rng.high:g})"
    elif italic:
        label = f"{name} = {value:g} {rng.unit} (< {rng.low:g})"
    return LabFlag(value=value, bold=bold, italic=italic, out_of_range_label=label)


@dataclass
class PersonDerived:
    # phân loại 7 cột lâm sàng M..S (None nếu bỏ trống, vd Sản phụ khoa của nam)
    clinical_cols: dict[str, int | None] = field(default_factory=dict)
    xep_loai: int | None = None
    xep_loai_note: str = ""

    # lab flags theo tên cột khung mẫu (ALT(GPT), AST(GOT), ...)
    lab_flags: dict[str, LabFlag] = field(default_factory=dict)

    cbc_result: str = ""  # "bt" / "x" / ""
    cbc_out_of_range: list[str] = field(default_factory=list)

    urinalysis_result: str = ""
    urinalysis_out_of_range: list[str] = field(default_factory=list)

    us_general_result: str = ""
    us_general_inferred: bool = False

    us_breast_result: str = ""
    us_breast_inferred: bool = False


def _max_phan_loai(*results) -> int | None:
    vals = [r.phan_loai for r in results if r is not None and r.phan_loai is not None]
    return max(vals) if vals else None


def compute_clinical_columns(p: Person) -> dict[str, int | None]:
    """Map các nhóm lâm sàng thô -> 7 cột M..S của khung mẫu."""
    cols: dict[str, int | None] = {}

    # Nội = max(8 nhóm con)
    noi_results = [p.clinical.get(k) for k in config.NOI_SUBGROUPS]
    cols["Nội"] = _max_phan_loai(*noi_results)

    # Ngoại, Da liễu, Mắt, TMH, RHM = map trực tiếp 1-1
    for group_key, template_col in config.GROUP_TO_TEMPLATE_COL.items():
        cols[template_col] = _max_phan_loai(p.clinical.get(group_key))

    # Sản phụ khoa = max(Sản khoa, Phụ khoa), bỏ trống nếu nam
    if p.is_nam:
        cols["Sản phụ khoa"] = None
    else:
        cols["Sản phụ khoa"] = _max_phan_loai(
            p.clinical.get("san_khoa"), p.clinical.get("phu_khoa")
        )

    return cols


def compute_xep_loai(clinical_cols: dict[str, int | None]) -> tuple[int | None, str]:
    """Xếp loại = max(Nội..RHM). Đây là quy tắc TỰ SUY RA (không có sẵn trong
    dữ liệu thô M03), nên luôn trả về ghi chú cảnh báo kèm theo."""
    vals = {k: v for k, v in clinical_cols.items() if v is not None}
    if not vals:
        return None, ""
    xep_loai = max(vals.values())
    detail = ", ".join(f"{k}={v}" for k, v in vals.items())
    note = f"Xếp loại tự tính = max({detail}) = {xep_loai} — kiểm tra lại."
    return xep_loai, note


def compute_lab_flags(p: Person) -> dict[str, LabFlag]:
    raw_by_col = {
        config.COL_ALT: p.labs.get("alt"),
        config.COL_AST: p.labs.get("ast"),
        config.COL_CREATININ: p.labs.get("creatinin"),
        config.COL_GLUCOSE: p.labs.get("glucose"),
        config.COL_CHOLESTEROL: p.labs.get("cholesterol"),
        config.COL_TRIGLYCERID: p.labs.get("triglycerid"),
        config.COL_HDL: p.labs.get("hdl"),
        config.COL_LDL: p.labs.get("ldl"),
    }
    flags: dict[str, LabFlag] = {}
    for template_col, chem_name, raw_col in config.BLOOD_CHEM_TEMPLATE_COLUMNS:
        value = raw_by_col.get(raw_col)
        rng = config.BLOOD_CHEM_RANGES[chem_name]
        flags[template_col] = flag_value(chem_name, value, rng)
    return flags


def compute_cbc(p: Person) -> tuple[str, list[str]]:
    """Tổng phân tích tế bào máu: 'bt' nếu mọi chỉ số có dữ liệu đều trong
    khoảng tham chiếu 7b, 'x' nếu có ít nhất 1 chỉ số ngoài khoảng."""
    raw_map = {
        "RBC (Hồng cầu)": p.labs.get("hc"),
        "HGB (Huyết sắc tố)": p.labs.get("hgb"),
        "HCT (Hematocrit)": p.labs.get("hct"),
        "MCV": p.labs.get("mcv"),
        "MCH": p.labs.get("mch"),
        "MCHC": p.labs.get("mchc"),
        "RDW-CV": p.labs.get("rdw"),
        "WBC (Bạch cầu)": p.labs.get("bach_cau"),
        "PLT (Tiểu cầu)": p.labs.get("tieu_cau"),
    }
    have_any = any(v is not None for v in raw_map.values())
    if not have_any:
        return "", []
    out_of_range = []
    for name, value in raw_map.items():
        if value is None:
            continue
        rng = config.CBC_RANGES[name]
        f = flag_value(name, value, rng)
        if f.out_of_range_label:
            out_of_range.append(f.out_of_range_label)
    return ("x" if out_of_range else "bt"), out_of_range


def compute_urinalysis(p: Person) -> tuple[str, list[str]]:
    ti_trong = p.labs.get("ti_trong_nt")
    ph = p.labs.get("ph_nt")
    have_any = ti_trong is not None or ph is not None
    if not have_any:
        return "", []
    out_of_range = []
    if ti_trong is not None:
        f = flag_value("Specific Gravity", ti_trong, config.URINALYSIS_RANGES["Specific Gravity"])
        if f.out_of_range_label:
            out_of_range.append(f.out_of_range_label)
    if ph is not None:
        f = flag_value("pH", ph, config.URINALYSIS_RANGES["pH"])
        if f.out_of_range_label:
            out_of_range.append(f.out_of_range_label)
    return ("x" if out_of_range else "bt"), out_of_range


def _infer_bt_x_from_text(text: str) -> str:
    """Suy đoán bt/x từ mô tả siêu âm dạng văn bản tự do — LUÔN là suy luận,
    phải được đánh dấu vào Cảnh báo, không được coi là dữ liệu gốc đã phân loại."""
    lowered = text.lower()
    if any(kw in lowered for kw in config.NORMAL_KEYWORDS):
        return "bt"
    return "x"


def compute_ultrasound(text: str) -> tuple[str, bool]:
    """Trả về (bt/x, inferred). inferred=True nghĩa là kết quả suy ra từ text,
    cần đưa vào Cảnh báo để người tổng hợp xác nhận lại."""
    text = text.strip()
    if not text:
        return "", False
    return _infer_bt_x_from_text(text), True


def derive_person(p: Person) -> PersonDerived:
    d = PersonDerived()
    d.clinical_cols = compute_clinical_columns(p)
    d.xep_loai, d.xep_loai_note = compute_xep_loai(d.clinical_cols)
    d.lab_flags = compute_lab_flags(p)
    d.cbc_result, d.cbc_out_of_range = compute_cbc(p)
    d.urinalysis_result, d.urinalysis_out_of_range = compute_urinalysis(p)
    d.us_general_result, d.us_general_inferred = compute_ultrasound(p.us_general_text)
    d.us_breast_result, d.us_breast_inferred = compute_ultrasound(p.us_breast_text)
    return d
