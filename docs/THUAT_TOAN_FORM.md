# Thuật toán nhân bản form QC

## Vấn đề của file mẫu

Hai phần `Người báo cáo` và `Kho & vận chuyển` không phải nội dung ô Excel. Chúng là đối tượng DrawingML dạng textbox nằm trong `xl/drawings/drawing1.xml`.

Do đó, cách sao chép `A:X` hoặc chèn thêm dòng chỉ giữ ô, border và merge nhưng không tạo textbox mới.

## Cách V7 xử lý

1. Mở XLSX như một gói ZIP bằng JSZip.
2. Tìm sheet `Xe to` từ `xl/workbook.xml` và relationships.
3. Đọc hai form gốc và tự nhận diện:
   - Dòng tiêu đề `Khách hàng / Customer`.
   - Dòng bắt đầu dữ liệu.
   - Dòng `Ghi chú`.
   - Dòng bắt đầu/kết thúc Form 2.
4. Dùng Form 2 làm nguồn vì Form 2 có đủ 25 dòng sản phẩm và đầy đủ phần chữ ký.
5. Khi cần Form 3 trở đi:
   - Nhân bản toàn bộ XML dòng của Form 2.
   - Dịch địa chỉ ô theo khoảng lặp.
   - Nhân bản các vùng merge.
   - Nhân bản hai textbox chữ ký và dịch vị trí neo theo số dòng.
   - Thêm page break.
   - Mở rộng Print Area.
6. Ghi tên khách vào cột J và sản phẩm vào cột K bằng inline string để không làm hỏng bảng sharedStrings.
7. Đóng gói lại thành XLSX.

## Cấu trúc được nhận diện trong mẫu gốc

- Form 1: dòng 2–36.
- Form 2: dòng 37–72.
- Form 3: dòng 73–108.
- Form 4: dòng 109–144.
- Form 5: dòng 145–180.
- Khoảng lặp Form 2 → Form 3: 36 dòng.
- Dữ liệu Form 2: J43/K43 đến dòng 67, tổng 25 dòng.
- Dữ liệu Form 3: J79/K79 đến dòng 103, tổng 25 dòng.
