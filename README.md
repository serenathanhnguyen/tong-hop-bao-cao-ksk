# tong-hop-bao-cao-ksk

Công cụ tổng hợp **Bảng tổng hợp phân loại sức khỏe** cho Trạm Y tế phường
Nhiêu Lộc từ file dữ liệu khám sức khỏe thô (Medinet, mẫu M03).

Upload file `.xlsx` thô (sheet `ThongTinHanhChinh`) → app tự map dữ liệu vào
đúng khung mẫu, tính Xếp loại, so sánh với bảng tham chiếu xét nghiệm (in
đậm/in nghiêng), lược bỏ cột không có dữ liệu, và xuất file Excel sẵn sàng in
theo đúng thể thức hành chính.

Toàn bộ quy tắc nghiệp vụ — và cách từng quy tắc được cài đặt ở đâu trong mã
nguồn — nằm ở [`docs/QUY_TAC.md`](docs/QUY_TAC.md). Đọc file đó trước khi sửa
logic.

## Chạy app

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Mở trình duyệt theo địa chỉ Streamlit in ra (mặc định `http://localhost:8501`),
upload file dữ liệu thô, điền tên đơn vị/đối tượng/năm, bấm **Tạo bảng tổng
hợp**, rồi tải file kết quả về.

## Cấu trúc mã nguồn

```
app.py                     Giao diện Streamlit (upload -> download)
ksk_report/
  config.py                Vị trí cột dữ liệu thô + bảng giá trị tham chiếu
  raw_reader.py             Đọc file M03 thô -> danh sách Person
  classify.py                Suy ra Xếp loại, so sánh tham chiếu, bt/x
  columns.py                  QUY TẮC 1: quyết định cột cận lâm sàng nào giữ lại
  notes.py                     QUY TẮC 2: tách Ghi chú (nguyên văn) / Cảnh báo (suy luận)
  workbook_builder.py           Dựng file Excel theo khung mẫu (mục 1-6)
  verify.py                      Bước 8: kiểm tra chéo định dạng trước khi gửi
  pipeline.py                     Orchestrator: raw path -> PipelineResult
docs/QUY_TAC.md             Đặc tả nghiệp vụ đầy đủ (nguồn quy tắc)
tests/
  make_fixture.py           Sinh file M03 giả lập để test (không có dữ liệu thật)
  test_pipeline.py           Test end-to-end
```

## Chạy test

```bash
pip install -r requirements.txt pytest
python3 -m pytest tests/ -q
```

Test dùng dữ liệu giả lập (`tests/make_fixture.py`) vì repo không chứa dữ
liệu khám sức khỏe thật — không có PII nào được commit vào repo này.

## Lưu ý quan trọng khi dùng

- Cột **CẢNH BÁO** ở cuối bảng liệt kê mọi thứ do công cụ **tự suy ra** (Xếp
  loại, bt/x của CTM/nước tiểu/siêu âm, chẩn đoán sơ bộ chưa xác định...).
  Đây không phải kết luận y khoa — rà soát lại rồi mới xóa cột này trước khi
  phát hành bảng chính thức.
- Dòng "Ngày ... tháng ... năm ..." mặc định để trống, tự điền tay sau khi
  in (trừ khi bạn tick điền ngày cụ thể trong app).
- Phần bố cục chữ ký cuối bảng là dựng lại theo mô tả bằng chữ (repo không có
  sẵn file khung mẫu gốc để copy y nguyên style) — xem lại vị trí trước khi
  in nếu cần khớp tuyệt đối với mẫu giấy tờ cũ.
