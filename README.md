# QC Generator V7 – Báo cáo kiểm tra xuất hàng

Ứng dụng web tĩnh chạy hoàn toàn trong trình duyệt. Không cần cài Python, không cần server xử lý dữ liệu và không tải file của người dùng lên Internet.

## Chức năng chính

- Đọc file DG theo cấu trúc: `Code | Tên khách hàng | Điểm giao hàng | Ship`.
- Đọc nhiều file sản phẩm 50 / 54 / 58 theo cấu trúc: `Code | Sản phẩm`.
- Lọc một mã Ship hoặc xử lý tất cả Ship.
- Gom theo **Tên khách hàng + Điểm giao hàng**.
- Cùng tên khách nhưng khác điểm giao vẫn tách riêng.
- Tùy chọn loại sản phẩm trùng và xóa chữ VietGap.
- Sắp xếp sản phẩm: `5kg → 2kg → 1kg → TF → 500g → loại khác`.
- Giữ nguyên một khách hàng trong cùng form nếu có tối đa 25 sản phẩm.
- Chỉ tách khách hàng khi có trên 25 sản phẩm.
- Xuất trực tiếp vào mẫu `BM-QC-26`, sheet `Xe to`.

## Điểm sửa quan trọng ở V7

File mẫu dùng **textbox/drawing** cho hai phần:

- `Người báo cáo (QC) / Reported by`
- `Kho & vận chuyển / Inspected by`

Nếu chỉ sao chép ô và định dạng, hai phần này sẽ bị mất. V7 không dùng cách chèn dòng thông thường. Ứng dụng chỉnh trực tiếp cấu trúc XLSX và sao chép đồng thời:

- Toàn bộ dòng của Form 2.
- Giá trị và style của ô.
- Merge cell.
- Chiều cao dòng.
- Page break và Print Area.
- Textbox `Người báo cáo`.
- Textbox `Kho & vận chuyển`.
- Phần `Người kiểm tra / Approved by` và tên người kiểm tra.

Form lặp được nhận diện từ file mẫu thay vì cố định vị trí. Với mẫu gốc trong dự án:

- Form 1 có 24 dòng sản phẩm.
- Form 2 và các form tự thêm có 25 dòng sản phẩm.
- Form tự thêm được nhân bản từ Form 2, từ dòng 37 đến dòng 72, tổng cộng 36 dòng.

## Cách sử dụng

1. Mở `index.html`.
2. Chọn file DG.
3. Chọn các file sản phẩm 50 / 54 / 58.
4. Chọn Ship cần xử lý hoặc bật `Xử lý tất cả Ship`.
5. Nhấn `Tạo báo cáo QC`.
6. Nhấn `Lưu / Chia sẻ Excel` trên iPhone/iPad hoặc `Tải Excel trực tiếp` trên máy tính.

Ứng dụng có sẵn mẫu QC tích hợp, nên không bắt buộc tải lại file mẫu.

## Kiểm tra demo

Mở file:

`sample_output/QC_demo_5_forms_CHU_KY_DAY_DU.xlsx`

Kiểm tra Form 3, Form 4 và Form 5. Mỗi form đều có:

- Người báo cáo (QC) / Reported by.
- Kho & vận chuyển / Inspected by.
- Người kiểm tra / Approved by.
- Quản Đắc Tân.

Form 3 bắt đầu tại dòng 73, dữ liệu bắt đầu tại J79/K79 và phần chữ ký nằm ở cuối form, khoảng dòng 106–108.

## Đưa lên GitHub Pages

1. Tạo repository mới trên GitHub.
2. Tải toàn bộ nội dung thư mục này lên nhánh `main`.
3. Vào `Settings → Pages`.
4. Chọn `Deploy from a branch`.
5. Chọn `main` và `/(root)`.
6. Nhấn `Save`.

## Cấu trúc thư mục

```text
QC_Generator_V7/
├── index.html
├── app.js
├── styles.css
├── template-base64.js
├── .nojekyll
├── libs/
│   ├── xlsx.full.min.js
│   └── jszip.min.js
├── template/
│   └── BM-QC-26.xlsx
├── sample_data/
├── sample_output/
└── docs/
```

## Lưu ý

- Không đổi tên sheet `Xe to` trong mẫu QC.
- Nếu dùng mẫu khác, mẫu đó cần giữ cấu trúc hai form gốc tương tự file tích hợp.
- Trên iPhone/iPad, nên dùng nút `Lưu / Chia sẻ Excel`, sau đó chọn `Lưu vào Tệp` hoặc Microsoft Excel.
