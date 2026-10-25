# Excel Batch Tool (JavaScript / Node.js)

## Yêu cầu
- Node.js bản LTS
- File Excel `.xlsx`
- Cột tiêu đề chính xác: `Fulfillment: Tracking Number`

## Chạy trên Windows
1. Giải nén thư mục dự án.
2. Mở thư mục bằng VS Code.
3. Mở Terminal trong VS Code.
4. Chạy:
   ```bash
   npm install
   npm start
   ```
5. Mở trình duyệt tại http://localhost:3000
6. Chọn file `.xlsx`, nhập số lượng (mặc định 100), bấm **Tạo file ZIP**.

## Quy tắc
- Mỗi file đầu ra được đọc độc lập từ file gốc.
- Chỉ thay các ô không trống bên dưới tiêu đề `Fulfillment: Tracking Number`.
- Số tracking là số nguyên ngẫu nhiên từ 1.000.000 đến 10.000.000, bao gồm cả hai đầu.
- Tên file: `Order_001.xlsx`, `Order_002.xlsx`, ...
- Tải xuống một file `Orders.zip`.

## Lưu ý
ExcelJS hỗ trợ nhiều tính năng Excel phổ biến nhưng không bảo đảm giữ nguyên mọi tính năng đặc biệt/đối tượng nhúng của mọi workbook. Dùng bản sao của file gốc và kiểm tra đầu ra trước khi dùng trong quy trình quan trọng.
