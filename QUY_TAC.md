# Khung mẫu "Bảng tổng hợp phân loại sức khỏe" + Bảng giá trị tham chiếu xét nghiệm

Đây là đặc tả gốc mà `ksk_report/` cài đặt. Dùng khi Jo upload file Excel dữ
liệu khám sức khỏe thô và yêu cầu làm bảng tổng hợp — bảng tổng hợp ra phải
theo đúng khung mẫu này (tên cột, thứ tự cột, công thức, phần chân trang).

Nguồn khung mẫu: file `BTH_KSK_2026_CB-NV_LOP_MAU_GIAO_HOA_TINH_THUONG_note.xlsx`
(sheet `Sheet1`), đơn vị mẫu: Trạm Y tế phường Nhiêu Lộc, đối tượng CB-NV Lớp
Mẫu giáo Hoa Tình Thương, khám sức khỏe định kỳ.

> Lưu ý khi đọc file này cùng mã nguồn: bản cài đặt trong `ksk_report/` được
> viết lại hoàn toàn từ đặc tả bằng chữ dưới đây (không có sẵn file khung mẫu
> gốc `.xlsx` để copy y nguyên style/merge), nên phần bố cục chữ ký cuối bảng
> (mục 6) là dựng lại hợp lý theo mô tả, không phải pixel-perfect với bản gốc
> — Jo nên xem qua và chỉnh vị trí nếu cần trước khi in.

## QUY TẮC QUAN TRỌNG 1: bỏ cột không có dữ liệu

Nếu một cột cận lâm sàng (Cholesterol toàn phần, Triglycerid, HDL-Cholesterol,
LDL-Cholesterol, Siêu âm tổng quát, Siêu âm tuyến vú, hay bất kỳ cột nào khác
trong khung mẫu gốc) TRỐNG ở TẤT CẢ mọi người trong dữ liệu thô (tức xét
nghiệm/hạng mục đó không được thực hiện cho đợt khám này) → XÓA HẲN cột đó
khỏi bảng tổng hợp, không chỉ để trống. Chỉ giữ lại các cột thực sự có dữ
liệu. Áp dụng nguyên tắc này cho mọi lần làm bảng tổng hợp về sau, không
riêng cột nào — kiểm tra từng cột trước khi xuất bảng.

Khi xóa cột: phải cập nhật lại toàn bộ các merge cell liên quan (tiêu đề
A4/A5, "CẬN LÂM SÀNG", "Xếp loại", "Ghi chú", "Cảnh báo", các dòng ký tên) và
VIẾT LẠI TAY các công thức COUNTA (dòng TỔNG CỘNG) và COUNTIF (khối thống kê
xếp loại, tham chiếu cột Xếp loại) theo cột/vị trí mới — openpyxl không tự
sửa text công thức khi xóa cột.

_(`ksk_report/columns.py` quyết định cột nào giữ lại; `workbook_builder.py`
dựng toàn bộ header/formula theo đúng vị trí cột động — không copy dòng mẫu
có sẵn nên không có rủi ro công thức trỏ sai vị trí cũ.)_

## QUY TẮC QUAN TRỌNG 2: cột GHI CHÚ chỉ chép nguyên văn từ dữ liệu gốc — không tự diễn giải/thêm nhận định

Cột "GHI CHÚ" (cuối bảng, ngay sau "Xếp loại") CHỈ được lấy/chép lại đúng
những gì đã có sẵn trong file dữ liệu thô của người đó (ví dụ: mã chẩn
đoán/ICD ở các cột "Chẩn đoán sơ bộ"/"Chẩn đoán xác định" theo từng chuyên
khoa, kết quả X-quang/siêu âm, "Kết luận" chung...). KHÔNG tự thêm nhận định,
diễn giải lâm sàng, hay kết luận riêng (ví dụ không tự viết "tăng huyết áp độ
1", "thừa cân", "men gan tăng nhẹ" nếu những cụm từ này không có sẵn trong dữ
liệu gốc — dù có thể suy ra từ số liệu). Ghi rõ tên chuyên khoa/hạng mục
trước mỗi mã/kết quả để dễ đọc (việc ghi tên chuyên khoa là trình bày, không
phải diễn giải). Nếu một người không có chẩn đoán/ghi chú gì trong dữ liệu
gốc, để trống hoặc chỉ ghi lại đúng phần "Kết luận" nếu có.

Mọi nhận định/tính toán riêng cần người tổng hợp kiểm tra lại (ví dụ: cách
tính cột Xếp loại, chỉ số nào được in đậm/in nghiêng do so với bảng tham
chiếu, chẩn đoán mới ở dạng "sơ bộ" chưa xác định...) → đưa vào một cột RIÊNG
có tên "CẢNH BÁO" ở cuối cùng bảng (sau cột Ghi chú), tiêu đề: "CẢNH BÁO\n(để
kiểm tra lại, xóa cột sau khi đã rà soát)". Cột này liệt kê ngắn gọn những
điểm người tổng hợp (Jo/trạm y tế) nên đối chiếu lại thủ công, và được thiết
kế để XÓA ĐI sau khi đã rà soát xong — không phải một phần vĩnh viễn của
bảng.

_(`ksk_report/notes.py` cài đặt đúng phân tách này: `ghi_chu` chỉ ghép các
chuỗi lấy nguyên văn từ ô dữ liệu gốc; mọi suy luận — Xếp loại tự tính,
lý do in đậm/nghiêng, bt/x suy ra từ text siêu âm, chẩn đoán sơ bộ chưa xác
định — đều đi vào `canh_bao`.)_

## QUY TẮC QUAN TRỌNG 3 (kỹ thuật khi dựng file bằng openpyxl): luôn RESET định dạng bold/italic khi ghi giá trị mới vào ô

Nếu chỉ gán `.value` cho ô mà không gán lại `.font`, định dạng in đậm/nghiêng
CŨ (nếu ô được tái sử dụng từ dòng mẫu có sẵn) sẽ dính lại dù giá trị mới đã
đổi — gây in đậm SAI cho các chỉ số thực ra nằm trong khoảng tham chiếu. Khi
ghi dữ liệu cho các ô chỉ số số học so với bảng tham chiếu (ALT, AST,
Creatinin, Glucose, Cholesterol, Triglycerid, HDL, LDL), LUÔN chủ động set
lại `Font(bold=..., italic=...)` theo đúng kết quả so sánh, KHÔNG dựa vào
định dạng có sẵn của ô. Sau khi dựng xong file, chạy một bước kiểm tra chéo:
với mỗi ô trong các cột này, in đậm chỉ khi giá trị > cận trên, in nghiêng
chỉ khi giá trị < cận dưới, còn lại phải là định dạng thường — nếu sai thì
sửa lại trước khi gửi.

_(`workbook_builder.py` set `Font(bold=flag.bold, italic=flag.italic)` tường
minh cho từng ô khi ghi dữ liệu — không có bước "copy dòng mẫu" nên không
dính định dạng cũ. `verify.py` + bước 8 trong `pipeline.py` chạy lại đúng
bước kiểm tra chéo này sau khi build, so giá trị thực tế trong ô với bảng
tham chiếu và font thực tế của ô, trả `verify_issues` nếu lệch.)_

## QUY TẮC QUAN TRỌNG 4: phần ngày ký cuối bảng — GIỮ dòng chữ "Ngày ... tháng ... năm ..." nhưng KHÔNG điền số cụ thể

Dòng ký tên cuối bảng (bên phải, phía trên "GIÁM ĐỐC") vẫn phải có chữ "Ngày
......... tháng ......... năm ........." (dạng có dấu chấm để trống, giống
mẫu giấy tờ hành chính) — KHÔNG được xóa hẳn ô này. Nhưng KHÔNG điền số
ngày/tháng/năm cụ thể vào đó (không tự lấy ngày khám hay bất kỳ ngày nào) để
Jo/trạm y tế tự điền tay sau khi in. Chỉ điền số cụ thể nếu Jo nói rõ ngày ký
là ngày nào.

_(`app.py` mặc định để trống — chỉ điền ngày cụ thể khi Jo chủ động tick
"Điền sẵn ngày ký cụ thể" và chọn ngày.)_

## 1. Phần tiêu đề (rows 1–5)

* A1:G1 = "ỦY BAN NHÂN DÂN PHƯỜNG NHIÊU LỘC", T1:{cột cuối}1 = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
* A2:G2 = "TRẠM Y TẾ", T2:{cột cuối}2 = "Độc lập - Tự do - Hạnh phúc"
* A4:{cột cuối}4 = "BẢNG TỔNG HỢP PHÂN LOẠI SỨC KHỎE {ĐỐI TƯỢNG}"
* A5:{cột cuối}5 = "KHÁM SỨC KHỎE ĐỊNH KỲ NĂM {năm}"
* Merge tới cột cuối cùng của bảng — cột cuối này thay đổi tùy có bao nhiêu cột cận lâm sàng thực tế + có thêm cột Cảnh báo.

## 2. Header cột (row 7–8) — thứ tự CHUẨN đầy đủ (trước khi lược bỏ cột trống)

* A: STT
* B: HỌ & TÊN
* C: TÊN (chỉ tên riêng, để sort)
* D: CCCD
* E: NGÀY SINH
* F: GIỚI TÍNH
* G: Chiều cao (Cm)
* H: Cân nặng (Kg)
* I: BMI — công thức `=H{row}/(G{row}*G{row})*10000`
* J: Huyết áp mmHg (dạng text "110/60")
* K–L: MẮT 10/10 → K = "P" (mắt phải), L = "T" (mắt trái) — điểm thị lực KHÔNG KÍNH
* M–S: KHÁM LÂM SÀNG (phân loại 1–5 từng chuyên khoa): M=Nội, N=Ngoại, O=Da liễu, P=Sản phụ khoa, Q=Mắt, R=Tai mũi họng, S=Răng hàm mặt
* T–AE: CẬN LÂM SÀNG, thứ tự đầy đủ: T=Tổng phân tích tế bào máu, U=ALT(GPT), V=AST(GOT), W=Creatinin, X=Glucose, Y=Cholesterol toàn phần, Z=Triglycerid, AA=HDL-Cholesterol, AB=LDL-Cholesterol, AC=Tổng phân tích nước tiểu, AD=Siêu âm tổng quát, AE=Siêu âm tuyến vú — CHỈ GIỮ các cột trong nhóm này có dữ liệu thực tế
* Xếp loại: kết quả phân loại sức khỏe tổng hợp 1–5
* Ghi chú (CK: Chuyên khoa; KPK: Khám phụ khoa; SA: Siêu âm) — chỉ chép nguyên văn từ dữ liệu gốc
* Cảnh báo (cột cuối cùng) — nhận định/điểm cần kiểm tra lại của người lập bảng

Cột Tổng phân tích tế bào máu, Tổng phân tích nước tiểu, Siêu âm (nếu còn
giữ) nhận giá trị "bt" (bình thường) hoặc "x" (bất thường). Các cột lâm sàng
(M–S) và Mắt (K–L) là số/điểm theo thang phân loại, không phải văn bản.

**Quy tắc tính Xếp loại** (suy ra, không có sẵn trong dữ liệu thô M03) = giá
trị LỚN NHẤT (max) trong 7 cột phân loại lâm sàng M–S của người đó. Không
tính cột cận lâm sàng vào Xếp loại. Nam giới bỏ qua cột Sản phụ khoa (để
trống). LUÔN ghi vào cột Cảnh báo: "Xếp loại tự tính = max(...) = {n} — kiểm
tra lại."

## 3. Vùng dữ liệu (từ row 9, mỗi người 1 dòng)

STT tăng dần từ 1; dữ liệu lấy từ phiếu khám lâm sàng + kết quả xét nghiệm +
siêu âm của từng người.

## 4. Dòng TỔNG CỘNG

* Cột A = "TỔNG CỘNG"
* Từ cột F đến cột Xếp loại: `=COUNTA({cột}{dòng đầu}:{cột}{dòng cuối})` — chỉ áp dụng cho các cột còn giữ lại, KHÔNG tính cho Ghi chú/Cảnh báo.

## 5. Khối thống kê xếp loại

* "Tổng số:" = tổng các dòng Loại 1..5, đơn vị "Người"
* "Loại 1:" đến "Loại 5:" = `=COUNTIF($<cột Xếp loại>${dòng đầu}:$<cột Xếp loại>${dòng cuối},"{n}")`, đơn vị "Người"
* Ghi chú cố định: "Ghi chú: chỉ số cận lâm sàng in đậm = cao hơn giá trị tham chiếu; in nghiêng = thấp hơn giá trị tham chiếu."

## 6. Phần ký tên

* "Ngày ......... tháng ......... năm ........." (căn phải) — xem QUY TẮC QUAN TRỌNG 4.
* Bên trái: "Người Lập Bảng" + tên (mặc định "Nguyễn Thị Ngọc Sương" trừ khi Jo nói khác).
* Bên phải: "GIÁM ĐỐC" (để trống tên, ký tay, trừ khi Jo cung cấp tên).

## 7. Bảng giá trị tham chiếu xét nghiệm sinh hóa máu + nước tiểu

Nước tiểu (Tổng phân tích nước tiểu): Leukocytes/Ketone/Nitrite/Bilirubin/
Protein/Blood/Ascorbic acid = Negative; Urobilinogen/Glucose = Normal;
Specific Gravity 1.000–1.030; pH 4.8–8.0. Dữ liệu thô M03 thường chỉ có tỉ
trọng và pH → nếu cả hai trong khoảng: "bt", ngược lại "x".

Xét nghiệm máu (mmol/L trừ khi ghi chú khác):

| Chỉ số | Khoảng tham chiếu |
|---|---|
| ALT (GPT) | 0–41 U/L |
| AST (GOT) | 0–40 U/L |
| Urê [Máu] | 2.5–7.5 mmol/L |
| Creatinin | 46–106 µmol/L |
| Glucose [Máu] | 3.90–6.20 mmol/L |
| Cholesterol toàn phần [Máu] | 0.50–1.70 mmol/L |
| Triglycerid [Máu] | 0.90–6.00 mmol/L |
| HDL-Cholesterol [Máu] | 0–3.34 mmol/L |
| LDL-Cholesterol [Máu] | 0–3.34 mmol/L |
| Acid Uric [Máu] | 3.90–5.90 mmol/L |

Cao hơn cận trên → in đậm; thấp hơn cận dưới → in nghiêng; trong khoảng →
định dạng thường. Chỉ số nào hoàn toàn không có dữ liệu ở bất kỳ ai → bỏ cột
đó (QUY TẮC QUAN TRỌNG 1). Lưu ý: Urê và Acid Uric có trong bảng tham chiếu
nhưng KHÔNG có cột riêng trong khung mẫu T–AE, nên không tạo cột cho 2 chỉ số
này dù có dữ liệu.

## 7b. Bảng giá trị tham chiếu Tổng phân tích tế bào máu ngoại vi

| Chỉ số | Khoảng tham chiếu |
|---|---|
| WBC (Bạch cầu) | 3.5–10.0 x10⁹/L |
| RBC (Hồng cầu) | 3.50–6.00 x10¹²/L |
| HGB (Huyết sắc tố) | 11.0–17.5 g/dL |
| HCT (Hematocrit) | 34–54 % |
| MCV | 78–100 fL |
| MCH | 24–33 Pg |
| MCHC | 31–37 g/dL |
| RDW-CV | 7–16 % |
| PLT (Tiểu cầu) | 150–400 x10⁹/L |

Dữ liệu thô (M03) thường chỉ có: HC (RBC), Huyết sắc tố (HGB), Hematocrit
(HCT), MCV, MCH, MCHC, RDW, Bạch cầu (WBC tổng), Tiểu cầu (PLT) — so các chỉ
số này với bảng trên để quyết định "bt" (tất cả trong khoảng) hay "x" (có ít
nhất 1 chỉ số ngoài khoảng — chi tiết luôn ghi vào Cảnh báo vì đây là nhận
định tự suy ra).

## Nguồn dữ liệu thô: file Mẫu 03 (Medinet, sheet `ThongTinHanhChinh`)

Dữ liệu từng người nằm ở sheet `ThongTinHanhChinh`, mỗi người 1 dòng bắt đầu
từ row 5 (row 1–4 là hướng dẫn/tên cột kỹ thuật). Toàn bộ vị trí cột (E=CCCD,
F=Họ tên, ..., các nhóm khám lâm sàng BH..DQ, xét nghiệm ED trở đi...) được
khai báo tường minh trong `ksk_report/config.py` — đọc file đó để tra vị trí
cột chính xác thay vì đoán lại từ file này.

Decimal Việt Nam dùng dấu phẩy (ví dụ "4,71" = 4.71) — `raw_reader.parse_vn_number`
tự đổi dấu phẩy thành dấu chấm khi tính toán/so sánh số.

## Quy trình khi Jo upload file dữ liệu thô và yêu cầu làm bảng tổng hợp

1. Đọc TOÀN BỘ các dòng có dữ liệu trong sheet dữ liệu thô — kiểm tra kỹ đến
   dòng cuối cùng thực sự có tên, không dựa vào giả định số người
   (`raw_reader.read_raw_workbook` quét hết `ws.max_row`, không dừng sớm khi
   gặp dòng trống xen giữa).
2. Map dữ liệu thô vào đúng cột theo `config.py`.
3. Rà soát từng cột cận lâm sàng: cột nào trống ở TẤT CẢ mọi người thì loại
   bỏ hẳn khỏi bảng (`columns.active_cls_columns`).
4. Tạo file Excel mới theo đúng cấu trúc mục 1–6 (`workbook_builder.build_workbook`).
5. Điền cột Ghi chú CHỈ bằng nội dung nguyên văn từ dữ liệu gốc
   (`notes.build_ghi_chu_canh_bao`).
6. Áp dụng bảng giá trị tham chiếu ở mục 7 và 7b để quyết định "bt"/"x" và in
   đậm/in nghiêng (`classify.py`), LUÔN reset font chủ động cho từng ô; mọi
   nhận định/tính toán riêng → cột Cảnh báo.
7. Điền dòng TỔNG CỘNG và khối thống kê xếp loại, phần ký tên — giữ dòng
   "Ngày ... tháng ... năm ..." dạng chấm chấm để trống trừ khi Jo nói rõ
   ngày ký cụ thể.
8. TRƯỚC KHI GỬI FILE: chạy lại một bước kiểm tra chéo toàn bộ các ô số ở
   cột cận lâm sàng — so giá trị với bảng tham chiếu và đối chiếu với định
   dạng bold/italic thực tế của ô (`verify.verify_chem_formatting`, chạy tự
   động trong `pipeline.run_pipeline` và hiển thị trong `app.py` nếu có ô
   lệch).
