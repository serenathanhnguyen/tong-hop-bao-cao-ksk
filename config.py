"""Hằng số cấu hình: map cột dữ liệu thô (Medinet M03, sheet ThongTinHanhChinh)
và bảng giá trị tham chiếu xét nghiệm.

Toàn bộ vị trí cột / khoảng tham chiếu lấy từ đặc tả "Khung mẫu Bảng tổng hợp
phân loại sức khỏe" — xem docs/QUY_TAC.md trong repo này để đọc đầy đủ quy tắc
diễn giải. File này chỉ chứa dữ liệu cấu hình thuần, không có logic.
"""

from __future__ import annotations

from dataclasses import dataclass, field

# ---------------------------------------------------------------------------
# Vùng dữ liệu thô (sheet ThongTinHanhChinh)
# ---------------------------------------------------------------------------

RAW_SHEET_NAME = "ThongTinHanhChinh"
RAW_DATA_START_ROW = 5  # row 1-4 là hướng dẫn / tên cột kỹ thuật

# Cột hành chính cơ bản
COL_CCCD = "E"
COL_HO_TEN = "F"
COL_NGAY_SINH = "G"
COL_GIOI_TINH = "H"  # 1 = Nam, 2 = Nữ
COL_CHIEU_CAO = "BA"
COL_CAN_NANG = "BB"
COL_HA_TAM_THU = "BE"
COL_HA_TAM_TRUONG = "BF"

# Thị lực không kính
COL_MAT_KHONG_KINH_PHAI = "DR"
COL_MAT_KHONG_KINH_TRAI = "DS"

# ---------------------------------------------------------------------------
# Khám lâm sàng: mỗi nhóm là 4 cột liên tiếp
#   [0] Chưa phát hiện bất thường  [1] Chẩn đoán sơ bộ
#   [2] Chẩn đoán xác định         [3] Phân loại (1-5)
# Sản khoa / Phụ khoa có thêm cột thứ 5 "Từ chối khám".
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ClinicalGroup:
    key: str
    label: str  # tên chuyên khoa hiển thị trong Ghi chú/Cảnh báo
    start_col: str
    ncols: int = 4
    has_tu_choi_kham: bool = False


CLINICAL_GROUPS: list[ClinicalGroup] = [
    ClinicalGroup("tuan_hoan", "Tuần Hoàn", "BH"),
    ClinicalGroup("ho_hap", "Hô Hấp", "BL"),
    ClinicalGroup("tieu_hoa", "Tiêu Hóa", "BP"),
    ClinicalGroup("than_tiet_nieu", "Thận - Tiết Niệu", "BT"),
    ClinicalGroup("noi_tiet", "Nội Tiết", "BX"),
    ClinicalGroup("co_xuong_khop", "Cơ Xương Khớp", "CB"),
    ClinicalGroup("than_kinh", "Thần Kinh", "CF"),
    ClinicalGroup("tam_than", "Tâm Thần", "CJ"),
    ClinicalGroup("ngoai_khoa", "Ngoại Khoa", "CN"),
    ClinicalGroup("da_lieu", "Da Liễu", "CR"),
    ClinicalGroup("san_khoa", "Sản Khoa", "CV", ncols=5, has_tu_choi_kham=True),
    ClinicalGroup("phu_khoa", "Phụ Khoa", "DA", ncols=5, has_tu_choi_kham=True),
    ClinicalGroup("mat", "Mắt", "DF"),
    ClinicalGroup("tmh", "Tai Mũi Họng", "DJ"),
    ClinicalGroup("rhm", "Răng Hàm Mặt", "DN"),
]

# Nhóm nào gộp (max phân loại) vào cột "Nội" của khung mẫu
NOI_SUBGROUPS = [
    "tuan_hoan",
    "ho_hap",
    "tieu_hoa",
    "than_tiet_nieu",
    "noi_tiet",
    "co_xuong_khop",
    "than_kinh",
    "tam_than",
]

# Map 1 nhóm lâm sàng thô -> cột khung mẫu (M..S); None nghĩa là gộp riêng (Nội, Sản phụ khoa)
GROUP_TO_TEMPLATE_COL = {
    "ngoai_khoa": "Ngoại",
    "da_lieu": "Da liễu",
    "mat": "Mắt",
    "tmh": "TMH",
    "rhm": "RHM",
}

# Thứ tự 7 cột lâm sàng M..S trong khung mẫu (dùng để tính Xếp loại = max)
TEMPLATE_CLINICAL_COLS = ["Nội", "Ngoại", "Da liễu", "Sản phụ khoa", "Mắt", "TMH", "RHM"]

# ---------------------------------------------------------------------------
# Xét nghiệm máu / nước tiểu / CTM / siêu âm (vùng cột ED trở đi)
# ---------------------------------------------------------------------------

COL_HC_SL = "EE"  # Số lượng hồng cầu (RBC)
COL_HGB = "EF"  # Huyết sắc tố
COL_HCT = "EG"  # Hematocrit
COL_MCV = "EH"
COL_MCH = "EI"
COL_MCHC = "EJ"
COL_RDW = "EK"
COL_BACH_CAU = "EL"  # WBC tổng
COL_TIEU_CAU = "ER"  # PLT
COL_GLUCOSE = "ES"
COL_URE = "ET"
COL_CREATININ = "EU"
COL_AST = "EV"  # GOT
COL_ALT = "EW"  # GPT
COL_TI_TRONG_NUOC_TIEU = "EX"
COL_PH_NUOC_TIEU = "EY"
COL_ACID_URIC = "FI"
COL_CHOLESTEROL = "FJ"
COL_TRIGLYCERID = "FK"
COL_HDL = "FL"
COL_LDL = "FM"
COL_SIEU_AM_BUNG = "FN"
COL_XQUANG = "FO"
COL_XQUANG_NHU = "FT"
COL_SIEU_AM_TUYEN_VU = "FU"
COL_KET_LUAN_CHUNG = "FV"

# ---------------------------------------------------------------------------
# Bảng tham chiếu 7: sinh hóa máu + nước tiểu (mục 7 trong đặc tả)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class RefRange:
    low: float | None
    high: float | None
    unit: str = ""


# Chỉ số nào nằm trong 4 cột số học U-X / Y-AB của khung mẫu (in đậm/nghiêng)
BLOOD_CHEM_RANGES: dict[str, RefRange] = {
    "ALT (GPT)": RefRange(0, 41, "U/L"),
    "AST (GOT)": RefRange(0, 40, "U/L"),
    "Creatinin": RefRange(46, 106, "µmol/L"),
    "Glucose": RefRange(3.90, 6.20, "mmol/L"),
    "Cholesterol toàn phần": RefRange(0.50, 1.70, "mmol/L"),
    "Triglycerid": RefRange(0.90, 6.00, "mmol/L"),
    "HDL-Cholesterol": RefRange(0, 3.34, "mmol/L"),
    "LDL-Cholesterol": RefRange(0, 3.34, "mmol/L"),
}

# Chỉ số tham chiếu nhưng KHÔNG có cột riêng trong khung mẫu (Urê, Acid Uric) —
# vẫn giữ ở đây để dùng khi cần, nhưng workbook_builder không tạo cột cho chúng.
BLOOD_CHEM_RANGES_EXTRA: dict[str, RefRange] = {
    "Urê": RefRange(2.5, 7.5, "mmol/L"),
    "Acid Uric": RefRange(3.90, 5.90, "mmol/L"),
}

# Mỗi entry: (tên cột khung mẫu, tên chỉ số trong BLOOD_CHEM_RANGES, cột dữ liệu thô)
BLOOD_CHEM_TEMPLATE_COLUMNS = [
    ("ALT(GPT)", "ALT (GPT)", COL_ALT),
    ("AST(GOT)", "AST (GOT)", COL_AST),
    ("Creatinin", "Creatinin", COL_CREATININ),
    ("Glucose", "Glucose", COL_GLUCOSE),
    ("Cholesterol toàn phần", "Cholesterol toàn phần", COL_CHOLESTEROL),
    ("Triglycerid", "Triglycerid", COL_TRIGLYCERID),
    ("HDL-Cholesterol", "HDL-Cholesterol", COL_HDL),
    ("LDL-Cholesterol", "LDL-Cholesterol", COL_LDL),
]

URINALYSIS_RANGES: dict[str, RefRange] = {
    "Specific Gravity": RefRange(1.000, 1.030),
    "pH": RefRange(4.8, 8.0),
}

# ---------------------------------------------------------------------------
# Bảng tham chiếu 7b: tổng phân tích tế bào máu ngoại vi
# ---------------------------------------------------------------------------

CBC_RANGES: dict[str, RefRange] = {
    "WBC (Bạch cầu)": RefRange(3.5, 10.0, "x10^9/L"),
    "RBC (Hồng cầu)": RefRange(3.50, 6.00, "x10^12/L"),
    "HGB (Huyết sắc tố)": RefRange(11.0, 17.5, "g/dL"),
    "HCT (Hematocrit)": RefRange(34, 54, "%"),
    "MCV": RefRange(78, 100, "fL"),
    "MCH": RefRange(24, 33, "Pg"),
    "MCHC": RefRange(31, 37, "g/dL"),
    "RDW-CV": RefRange(7, 16, "%"),
    "PLT (Tiểu cầu)": RefRange(150, 400, "x10^9/L"),
}

# map tên chỉ số CBC -> cột dữ liệu thô tương ứng (chỉ những chỉ số M03 thường có)
CBC_RAW_COLS = {
    "RBC (Hồng cầu)": COL_HC_SL,
    "HGB (Huyết sắc tố)": COL_HGB,
    "HCT (Hematocrit)": COL_HCT,
    "MCV": COL_MCV,
    "MCH": COL_MCH,
    "MCHC": COL_MCHC,
    "RDW-CV": COL_RDW,
    "WBC (Bạch cầu)": COL_BACH_CAU,
    "PLT (Tiểu cầu)": COL_TIEU_CAU,
}

# ---------------------------------------------------------------------------
# Cột "cận lâm sàng" đầy đủ theo khung mẫu gốc (T..AE), trước khi lược bỏ cột trống.
# key: mã cột nội bộ; label: tiêu đề hiển thị trong header.
# kind: "cbc" | "chem" | "urinalysis" | "us_general" | "us_breast"
# ---------------------------------------------------------------------------

FULL_CLS_COLUMNS = [
    {"key": "cbc", "label": "Tổng phân tích tế bào máu", "kind": "cbc"},
    {"key": "alt", "label": "ALT(GPT)", "kind": "chem", "chem_name": "ALT (GPT)", "raw_col": COL_ALT},
    {"key": "ast", "label": "AST(GOT)", "kind": "chem", "chem_name": "AST (GOT)", "raw_col": COL_AST},
    {"key": "creatinin", "label": "Creatinin", "kind": "chem", "chem_name": "Creatinin", "raw_col": COL_CREATININ},
    {"key": "glucose", "label": "Glucose", "kind": "chem", "chem_name": "Glucose", "raw_col": COL_GLUCOSE},
    {"key": "cholesterol", "label": "Cholesterol toàn phần", "kind": "chem", "chem_name": "Cholesterol toàn phần", "raw_col": COL_CHOLESTEROL},
    {"key": "triglycerid", "label": "Triglycerid", "kind": "chem", "chem_name": "Triglycerid", "raw_col": COL_TRIGLYCERID},
    {"key": "hdl", "label": "HDL-Cholesterol", "kind": "chem", "chem_name": "HDL-Cholesterol", "raw_col": COL_HDL},
    {"key": "ldl", "label": "LDL-Cholesterol", "kind": "chem", "chem_name": "LDL-Cholesterol", "raw_col": COL_LDL},
    {"key": "urinalysis", "label": "Tổng phân tích nước tiểu", "kind": "urinalysis"},
    {"key": "us_general", "label": "Siêu âm tổng quát", "kind": "us_general", "raw_col": COL_SIEU_AM_BUNG},
    {"key": "us_breast", "label": "Siêu âm tuyến vú", "kind": "us_breast", "raw_col": COL_SIEU_AM_TUYEN_VU},
]

# Từ khoá coi là "bình thường" khi đọc kết quả siêu âm dạng văn bản tự do
NORMAL_KEYWORDS = [
    "bình thường",
    "chưa phát hiện bất thường",
    "chưa ghi nhận bất thường",
    "không phát hiện bất thường",
]

DEFAULT_NGUOI_LAP_BANG = "Nguyễn Thị Ngọc Sương"
