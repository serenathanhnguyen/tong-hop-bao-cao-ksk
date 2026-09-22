import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.make_fixture import build_fixture  # noqa: E402

from ksk_report.pipeline import run_pipeline  # noqa: E402
from ksk_report.workbook_builder import ReportMeta, DATA_START_ROW  # noqa: E402


def test_pipeline_end_to_end():
    with tempfile.TemporaryDirectory() as tmp:
        raw_path = os.path.join(tmp, "fixture_m03.xlsx")
        build_fixture(raw_path)

        meta = ReportMeta(
            ten_phuong="PHƯỜNG NHIÊU LỘC",
            ten_tram_yte="TRẠM Y TẾ",
            doi_tuong="CB-NV LỚP MẪU GIÁO HOA TÌNH THƯƠNG",
            nam="2026",
        )
        result = run_pipeline(raw_path, meta)

        # 1. Đọc đủ 3 người, kể cả người sau dòng trống xen giữa
        assert result.so_nguoi == 3

        ws = result.workbook["Sheet1"]

        # 2. QUY TẮC 1: Cholesterol/Triglycerid/HDL/LDL/Siêu âm tuyến vú trống ở
        #    toàn bộ 3 người trong fixture -> phải bị lược bỏ khỏi bảng.
        assert "Cholesterol toàn phần" in result.cot_da_bo
        assert "Triglycerid" in result.cot_da_bo
        assert "HDL-Cholesterol" in result.cot_da_bo
        assert "LDL-Cholesterol" in result.cot_da_bo
        assert "Siêu âm tuyến vú" in result.cot_da_bo
        # ALT/AST/Creatinin/Glucose có dữ liệu -> phải được giữ lại
        assert "ALT(GPT)" not in result.cot_da_bo

        # 3. Không có ô nào lệch định dạng bold/italic sau khi build
        assert result.verify_issues == []

        # 4. Header hàng 7/8 tồn tại, không rỗng
        header_row8_values = [c.value for c in ws[8] if c.value]
        assert any("ALT" in str(v) for v in header_row8_values)

        # 5. Dữ liệu người 2 (Trần Thị Bình, ALT=68 > 41) phải được in đậm
        found_bold_alt = False
        for row in ws.iter_rows(min_row=DATA_START_ROW, max_row=DATA_START_ROW + 2):
            for cell in row:
                if cell.value == 68.0 or cell.value == 68:
                    found_bold_alt = cell.font.bold
        assert found_bold_alt is True

        # 6. Dòng ngày ký vẫn còn dạng chấm chấm để trống (không tự điền ngày)
        found_ngay_placeholder = False
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("Ngày ....."):
                    found_ngay_placeholder = True
        assert found_ngay_placeholder is True

        # 7. Cảnh báo có ghi nhận "chẩn đoán sơ bộ ... chưa xác định" cho người 2
        # (nhãn "CẢNH BÁO" là cột không group -> nằm ở row 7, merge dọc xuống row 8)
        canh_bao_col = None
        for idx, cell in enumerate(list(ws[7]), start=1):
            if cell.value and "CẢNH BÁO" in str(cell.value):
                canh_bao_col = idx
        assert canh_bao_col is not None
        canh_bao_values = [
            ws.cell(row=r, column=canh_bao_col).value or ""
            for r in range(DATA_START_ROW, DATA_START_ROW + 3)
        ]
        assert any("chưa có chẩn đoán xác định" in v for v in canh_bao_values)


if __name__ == "__main__":
    test_pipeline_end_to_end()
    print("OK")
