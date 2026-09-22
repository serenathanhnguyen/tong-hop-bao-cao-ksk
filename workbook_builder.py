"""Dựng file Excel "Bảng tổng hợp phân loại sức khỏe" theo đúng khung mẫu
mô tả trong docs/QUY_TAC.md (mục 1-6).

Toàn bộ font/size/màu/độ rộng cột/merge/offset dòng trong file này đã được
đối chiếu trực tiếp với file khung mẫu thật do Jo cung cấp
(BTH_KSK_2026_CB-NV_CONG_TY_CA_SAU_VIET_PHONG) — không phải suy đoán từ mô
tả chữ. Khi cần đối chiếu lại, xem ghi chú "khớp file thật" cạnh từng đoạn.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from . import config
from .classify import PersonDerived
from .raw_reader import Person

FONT_NAME = "Times New Roman"  # khớp file thật: toàn bộ sheet dùng Times New Roman

THIN = Side(style="thin", color="000000")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
CENTER_NOWRAP = Alignment(horizontal="center", vertical="center", wrap_text=False)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
VERTICAL_HEADER = Alignment(horizontal="center", vertical="center", wrap_text=True, text_rotation=90)


def F(*, bold: bool = False, italic: bool = False, size: float = 12) -> Font:
    return Font(name=FONT_NAME, bold=bold, italic=italic, size=size)


@dataclass
class ReportMeta:
    ten_phuong: str = "PHƯỜNG NHIÊU LỘC"
    ten_tram_yte: str = "TRẠM Y TẾ"
    doi_tuong: str = ""  # vd "CB-NV CÔNG TY TNHH CÁ SẤU VIỆT PHONG"
    nam: str = ""
    nguoi_lap_bang: str = config.DEFAULT_NGUOI_LAP_BANG
    giam_doc: str = ""  # để trống mặc định, ký tay
    ngay_ky: date | None = None  # None -> để trống dạng "Ngày ... tháng ... năm ..."


@dataclass
class ColumnDef:
    key: str
    label: str
    group: str | None = None  # "mat" | "lam_sang" | "can_lam_sang" | None


# Nhãn header khớp nguyên văn file khung mẫu thật (kể cả xuống dòng \n)
GROUP_LABELS = {
    "mat": "MẮT 10/10",
    "lam_sang": "KHÁM LÂM SÀNG\n(Phân loại từ 1 đến 5)",
    "can_lam_sang": "CẬN LÂM SÀNG",
}


def build_column_defs(active_cls_cols: list[dict]) -> list[ColumnDef]:
    cols: list[ColumnDef] = [
        ColumnDef("stt", "STT"),
        ColumnDef("ho_ten", "HỌ & TÊN"),
        ColumnDef("ten", "TÊN"),
        ColumnDef("cccd", "CCCD"),
        ColumnDef("ngay_sinh", "NGÀY SINH"),
        ColumnDef("gioi_tinh", "GIỚI TÍNH"),
        ColumnDef("chieu_cao", "Chiều cao\n(Cm)"),
        ColumnDef("can_nang", "Cân nặng\n(Kg)"),
        ColumnDef("bmi", "BMI"),
        ColumnDef("huyet_ap", "Huyết áp mmHg"),
        ColumnDef("mat_p", "P", group="mat"),
        ColumnDef("mat_t", "T", group="mat"),
    ]
    for label in config.TEMPLATE_CLINICAL_COLS:
        cols.append(ColumnDef(f"ls_{label}", label, group="lam_sang"))
    for c in active_cls_cols:
        cols.append(ColumnDef(f"cls_{c['key']}", c["label"], group="can_lam_sang"))
    cols.append(ColumnDef("xep_loai", "Xếp loại"))
    cols.append(ColumnDef("ghi_chu", "GHI CHÚ\n(CK: Chuyên khoa; KPK: Khám phụ khoa; SA: Siêu âm)"))
    cols.append(ColumnDef("canh_bao", "CẢNH BÁO\n(để kiểm tra lại, xóa cột sau khi đã rà soát)"))
    return cols


ROW_TITLE1, ROW_TITLE2 = 1, 2
ROW_TITLE4, ROW_TITLE5 = 4, 5
ROW_HEADER_GROUP, ROW_HEADER_SUB = 7, 8
DATA_START_ROW = 9

# Cột nào (theo key) dùng header xoay dọc 90° trong khung mẫu thật (M-S, cận
# lâm sàng) — chữ dài trong cột hẹp.
_ROTATED_GROUPS = {"lam_sang", "can_lam_sang"}

_VN_MONTHS_RE = None  # placeholder (không dùng, giữ chỗ cho rõ ý không cần regex)


def _parse_ngay_sinh(v: Any) -> tuple[Any, str]:
    """Trả về (value_to_write, number_format). Khớp file thật: NGÀY SINH lưu
    dạng ngày thật (không phải text) để Excel sort/lọc được, số format
    dd/mm/yyyy cho rõ ràng theo cách đọc ngày-tháng-năm của VN."""
    if v is None or v == "":
        return "", "General"
    if isinstance(v, (datetime, date)):
        return v, "dd/mm/yyyy"
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt), "dd/mm/yyyy"
        except ValueError:
            continue
    # không parse được -> giữ nguyên text gốc, không ép định dạng ngày
    return s, "General"


def _set(ws: Worksheet, row: int, col: int, value: Any, *, font: Font | None = None,
         align: Alignment | None = None, border: Border | None = BOX,
         number_format: str | None = None):
    cell = ws.cell(row=row, column=col, value=value)
    # QUY TẮC QUAN TRỌNG 3: luôn set lại font tường minh, không dựa vào mặc định của ô.
    cell.font = font if font is not None else F()
    if align is not None:
        cell.alignment = align
    if border is not None:
        cell.border = border
    if number_format is not None:
        cell.number_format = number_format
    return cell


def build_workbook(
    meta: ReportMeta,
    people: list[Person],
    derived: list[PersonDerived],
    active_cls_cols: list[dict],
    ghi_chu_list: list[str],
    canh_bao_list: list[str],
) -> Workbook:
    cols = build_column_defs(active_cls_cols)
    ncols = len(cols)
    last_col_idx = ncols
    last_col_letter = get_column_letter(last_col_idx)

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # ---- Mục 1: tiêu đề (khớp file thật: dòng 1-2 bold size 13, dòng 4-5 bold size 14) ----
    ws.merge_cells(f"A{ROW_TITLE1}:G{ROW_TITLE1}")
    _set(ws, ROW_TITLE1, 1, "ỦY BAN NHÂN DÂN " + meta.ten_phuong.upper(), font=F(bold=True, size=13), align=CENTER, border=None)
    ws.merge_cells(f"T{ROW_TITLE1}:{last_col_letter}{ROW_TITLE1}")
    _set(ws, ROW_TITLE1, 20, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", font=F(bold=True, size=13), align=CENTER, border=None)

    ws.merge_cells(f"A{ROW_TITLE2}:G{ROW_TITLE2}")
    _set(ws, ROW_TITLE2, 1, meta.ten_tram_yte.upper(), font=F(bold=True, size=13), align=CENTER, border=None)
    ws.merge_cells(f"T{ROW_TITLE2}:{last_col_letter}{ROW_TITLE2}")
    _set(ws, ROW_TITLE2, 20, "Độc lập - Tự do - Hạnh phúc", font=F(bold=True, size=13), align=CENTER, border=None)

    # row 3: dòng đệm trống (khớp file thật — có merge sẵn nhưng không có nội dung)
    ws.merge_cells(f"B3:F3")
    ws.merge_cells(f"N3:{last_col_letter}3")

    doi_tuong = meta.doi_tuong.strip() or "..."
    ws.merge_cells(f"A{ROW_TITLE4}:{last_col_letter}{ROW_TITLE4}")
    _set(ws, ROW_TITLE4, 1, f"BẢNG TỔNG HỢP PHÂN LOẠI SỨC KHỎE {doi_tuong}".upper(),
         font=F(bold=True, size=14), align=CENTER, border=None)

    nam = meta.nam.strip() or "..."
    ws.merge_cells(f"A{ROW_TITLE5}:{last_col_letter}{ROW_TITLE5}")
    _set(ws, ROW_TITLE5, 1, f"KHÁM SỨC KHỎE ĐỊNH KỲ NĂM {nam}".upper(),
         font=F(bold=True, size=14), align=CENTER, border=None)

    # ---- Mục 2: header cột (row 7-8) ---------------------------------------
    group_spans: dict[str, list[int]] = {}
    for idx, c in enumerate(cols, start=1):
        if c.group:
            group_spans.setdefault(c.group, []).append(idx)

    for idx, c in enumerate(cols, start=1):
        letter = get_column_letter(idx)
        if c.group:
            span = group_spans[c.group]
            sub_align = CENTER_NOWRAP if c.group == "mat" else VERTICAL_HEADER
            if idx == span[0]:
                first_l, last_l = get_column_letter(span[0]), get_column_letter(span[-1])
                ws.merge_cells(f"{first_l}{ROW_HEADER_GROUP}:{last_l}{ROW_HEADER_GROUP}")
                _set(ws, ROW_HEADER_GROUP, span[0], GROUP_LABELS[c.group], font=F(bold=True, size=10), align=CENTER)
                for j in span[1:]:
                    _set(ws, ROW_HEADER_GROUP, j, None, align=CENTER)
            _set(ws, ROW_HEADER_SUB, idx, c.label, font=F(bold=True, size=10), align=sub_align)
        else:
            ws.merge_cells(f"{letter}{ROW_HEADER_GROUP}:{letter}{ROW_HEADER_SUB}")
            _set(ws, ROW_HEADER_GROUP, idx, c.label, font=F(bold=True, size=10), align=CENTER)
            _set(ws, ROW_HEADER_SUB, idx, None, align=CENTER)

    # ---- Mục 3: vùng dữ liệu -----------------------------------------------
    col_index = {c.key: i + 1 for i, c in enumerate(cols)}

    def col(key: str) -> int:
        return col_index[key]

    for i, (p, d) in enumerate(zip(people, derived)):
        row = DATA_START_ROW + i
        ghi_chu, canh_bao = ghi_chu_list[i], canh_bao_list[i]

        _set(ws, row, col("stt"), i + 1, align=CENTER_NOWRAP)
        # khớp file thật: Họ tên căn giữa (không phải trái)
        _set(ws, row, col("ho_ten"), p.ho_ten, align=CENTER_NOWRAP)
        ten_rieng = p.ho_ten.strip().split(" ")[-1] if p.ho_ten.strip() else ""
        _set(ws, row, col("ten"), ten_rieng, align=CENTER_NOWRAP)
        _set(ws, row, col("cccd"), p.cccd, align=CENTER_NOWRAP)
        ngay_sinh_val, ngay_sinh_fmt = _parse_ngay_sinh(p.ngay_sinh)
        _set(ws, row, col("ngay_sinh"), ngay_sinh_val, align=CENTER_NOWRAP, number_format=ngay_sinh_fmt)
        _set(ws, row, col("gioi_tinh"), p.gioi_tinh, align=CENTER_NOWRAP)
        _set(ws, row, col("chieu_cao"), p.chieu_cao, align=CENTER_NOWRAP)
        _set(ws, row, col("can_nang"), p.can_nang, align=CENTER_NOWRAP)
        g_letter, h_letter = get_column_letter(col("chieu_cao")), get_column_letter(col("can_nang"))
        _set(ws, row, col("bmi"), f"={h_letter}{row}/({g_letter}{row}*{g_letter}{row})*10000",
             align=CENTER_NOWRAP, number_format="0.0")
        _set(ws, row, col("huyet_ap"), p.huyet_ap_text, align=CENTER_NOWRAP)
        _set(ws, row, col("mat_p"), p.mat_phai, align=CENTER_NOWRAP)
        _set(ws, row, col("mat_t"), p.mat_trai, align=CENTER_NOWRAP)

        for label in config.TEMPLATE_CLINICAL_COLS:
            val = d.clinical_cols.get(label)
            _set(ws, row, col(f"ls_{label}"), val if val is not None else "", align=CENTER_NOWRAP)

        for c in active_cls_cols:
            key = f"cls_{c['key']}"
            kind = c["kind"]
            if kind == "cbc":
                _set(ws, row, col(key), d.cbc_result, align=CENTER_NOWRAP)
            elif kind == "chem":
                flag = d.lab_flags[c["label"]]
                font = F(bold=flag.bold, italic=flag.italic)
                _set(ws, row, col(key), flag.value, font=font, align=CENTER_NOWRAP)
            elif kind == "urinalysis":
                _set(ws, row, col(key), d.urinalysis_result, align=CENTER_NOWRAP)
            elif kind == "us_general":
                _set(ws, row, col(key), d.us_general_result, align=CENTER_NOWRAP)
            elif kind == "us_breast":
                _set(ws, row, col(key), d.us_breast_result, align=CENTER_NOWRAP)

        _set(ws, row, col("xep_loai"), d.xep_loai if d.xep_loai is not None else "", font=F(size=12), align=CENTER_NOWRAP)
        # khớp file thật: Ghi chú/Cảnh báo căn GIỮA (không phải trái) + wrap
        _set(ws, row, col("ghi_chu"), ghi_chu, align=CENTER)
        _set(ws, row, col("canh_bao"), canh_bao, font=F(size=12), align=CENTER)

    last_data_row = DATA_START_ROW + len(people) - 1 if people else DATA_START_ROW

    # ---- Mục 4: dòng TỔNG CỘNG ----------------------------------------------
    tong_cong_row = last_data_row + 1
    ws.merge_cells(f"A{tong_cong_row}:E{tong_cong_row}")
    _set(ws, tong_cong_row, 1, "TỔNG CỘNG", font=F(bold=True, size=13), align=CENTER)
    if people:
        for c in cols:
            if c.key in ("stt", "ho_ten", "ten", "cccd", "ngay_sinh", "ghi_chu", "canh_bao"):
                continue
            idx = col_index[c.key]
            letter = get_column_letter(idx)
            _set(ws, tong_cong_row, idx,
                 f"=COUNTA({letter}{DATA_START_ROW}:{letter}{last_data_row})",
                 font=F(bold=True, size=12), align=CENTER_NOWRAP)

    # ---- Mục 5: khối thống kê xếp loại (khớp file thật: nhãn cột B, giá trị
    # cột C, đơn vị cột E, bắt đầu ngay tong_cong_row+2, không viền) ----------
    stats_start = tong_cong_row + 2
    xep_loai_letter = get_column_letter(col_index["xep_loai"])
    _set(ws, stats_start, 2, "Tổng số:", font=F(bold=True, size=12), align=LEFT, border=None)
    _set(ws, stats_start, 3, f"=SUM(C{stats_start + 1}:C{stats_start + 5})", font=F(bold=True, size=12), align=CENTER, border=None)
    _set(ws, stats_start, 5, "Người", font=F(bold=True, size=12), align=LEFT, border=None)
    for i in range(1, 6):
        r = stats_start + i
        _set(ws, r, 2, f"Loại {i}:", font=F(size=12), align=LEFT, border=None)
        if people:
            _set(ws, r, 3,
                 f'=COUNTIF(${xep_loai_letter}${DATA_START_ROW}:${xep_loai_letter}${last_data_row},"{i}")',
                 font=F(size=12), align=CENTER, border=None)
        else:
            _set(ws, r, 3, 0, font=F(size=12), align=CENTER, border=None)
        _set(ws, r, 5, "Người", font=F(size=12), align=LEFT, border=None)

    # Ghi chú cố định: 1 ô B, không merge, không viền, in nghiêng (khớp file thật)
    note_row = stats_start + 6
    _set(ws, note_row, 2,
         "Ghi chú: chỉ số cận lâm sàng in đậm = cao hơn giá trị tham chiếu; "
         "in nghiêng = thấp hơn giá trị tham chiếu.",
         font=F(italic=True, size=11), align=Alignment(horizontal="left", vertical="center"), border=None)

    # ---- Mục 6: phần ký tên (khớp file thật: date_row=note_row+1,
    # label_row=date_row+1, name_row=label_row+6) -----------------------------
    sign_col_start = max(1, last_col_idx - 6)
    sign_col_letter_start = get_column_letter(sign_col_start)

    date_row = note_row + 1
    ws.merge_cells(f"{sign_col_letter_start}{date_row}:{last_col_letter}{date_row}")
    if meta.ngay_ky is not None:
        ngay_text = (f"Ngày {meta.ngay_ky.day} tháng {meta.ngay_ky.month} "
                     f"năm {meta.ngay_ky.year}")
    else:
        # QUY TẮC QUAN TRỌNG 4: giữ dạng chấm chấm để trống, không tự điền ngày.
        ngay_text = "Ngày ......... tháng ......... năm ........."
    _set(ws, date_row, sign_col_start, ngay_text, font=F(italic=True, size=12), align=CENTER, border=None)

    label_row = date_row + 1
    # khớp file thật: cả 2 nhãn đều merge rộng 6 cột (B:G bên trái, sign_col:cuối bên phải)
    ws.merge_cells(f"B{label_row}:G{label_row}")
    _set(ws, label_row, 2, "Người Lập Bảng", font=F(bold=True, size=12), align=CENTER, border=None)
    ws.merge_cells(f"{sign_col_letter_start}{label_row}:{last_col_letter}{label_row}")
    _set(ws, label_row, sign_col_start, "GIÁM ĐỐC", font=F(bold=True, size=12), align=CENTER, border=None)

    name_row = label_row + 6
    ws.merge_cells(f"B{name_row}:G{name_row}")
    _set(ws, name_row, 2, meta.nguoi_lap_bang, font=F(bold=True, size=12), align=CENTER, border=None)
    # khớp file thật: ô tên Giám đốc luôn merge sẵn (để trống chờ ký), không chỉ khi có tên
    ws.merge_cells(f"{sign_col_letter_start}{name_row}:{last_col_letter}{name_row}")
    _set(ws, name_row, sign_col_start, meta.giam_doc.strip(), font=F(bold=True, size=12), align=CENTER, border=None)

    # ---- Độ rộng cột (khớp file thật cho A-L; cột lâm sàng/cận lâm sàng hẹp
    # vì header xoay dọc; Ghi chú/Cảnh báo rộng vừa phải — trang in đã bật
    # fitToPage nên không cần ép hẹp như bản gốc) ------------------------------
    fixed_widths = {
        "stt": 4.5, "ho_ten": 30, "ten": 10.5, "cccd": 15, "ngay_sinh": 11,
        "gioi_tinh": 7, "chieu_cao": 9, "can_nang": 9, "bmi": 8.5, "huyet_ap": 10,
        "mat_p": 4.5, "mat_t": 7,
    }
    for key, width in fixed_widths.items():
        ws.column_dimensions[get_column_letter(col_index[key])].width = width
    for label in config.TEMPLATE_CLINICAL_COLS:
        ws.column_dimensions[get_column_letter(col_index[f"ls_{label}"])].width = 5.5
    for c in active_cls_cols:
        ws.column_dimensions[get_column_letter(col_index[f"cls_{c['key']}"])].width = 6
    ws.column_dimensions[get_column_letter(col_index["xep_loai"])].width = 7
    ws.column_dimensions[get_column_letter(col_index["ghi_chu"])].width = 32
    ws.column_dimensions[get_column_letter(col_index["canh_bao"])].width = 32

    # ---- Chiều cao dòng header (chữ xoay dọc + wrap cần dòng cao; dữ liệu để
    # Excel tự autofit theo wrap text, không ép cứng vì độ dài Ghi chú/Cảnh
    # báo thay đổi theo từng người) --------------------------------------------
    ws.row_dimensions[ROW_HEADER_GROUP].height = 30
    ws.row_dimensions[ROW_HEADER_SUB].height = 85

    # ---- Đóng băng: giữ cố định các cột định danh + mắt phải khi cuộn ngang
    # (khớp file thật: freeze ngay trước cột "T" - mắt trái) -------------------
    freeze_col_letter = get_column_letter(col_index["mat_t"])
    ws.freeze_panes = f"{freeze_col_letter}{DATA_START_ROW}"

    # ---- Thiết lập trang in: A4 ngang, tự co vừa trang (khớp file thật) -----
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = 0.57
    ws.page_margins.right = 0.41
    ws.page_margins.top = 0.39
    ws.page_margins.bottom = 0.44
    last_row_for_print = name_row
    ws.print_area = f"A1:{last_col_letter}{last_row_for_print}"

    return wb
