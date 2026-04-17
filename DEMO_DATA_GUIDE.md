# DEMO DATA GUIDE

## Mục đích
Tài liệu này mô tả bộ dữ liệu demo kiểm thử cho workbook L&D v3.

## Cách dùng nhanh
1. Đảm bảo `Danh mục nhân viên` đã có dữ liệu thật.
2. Vào menu `L&D Vận hành -> Nạp dữ liệu demo`.
3. Xác nhận thao tác.
4. Hệ thống sẽ tự:
   - nạp `27` nhân viên demo từ snapshot HR chuẩn
   - nạp `5` khóa học demo
   - nạp `5` lớp đào tạo demo
   - sinh `25` dòng `Data raw học viên`
   - chạy `Làm mới báo cáo`

## Xóa bộ demo
- Để tránh bấm nhầm, chức năng xóa demo **không nằm trong menu**.
- Mở `Apps Script`, chọn hàm `clearDemoDataset()` và bấm `Run`.
- Hàm này chỉ xóa:
  - `Danh mục khóa học` demo
  - `Lớp đào tạo` demo
  - `Data raw học viên` demo
- `Danh mục nhân viên` được giữ nguyên.
- Sau khi xóa, hệ thống tự `Làm mới báo cáo`.

## Phạm vi dữ liệu demo
- `Danh mục nhân viên`: upsert theo `Mã nhân viên` từ file `GSH_HR Master Data_QTHC_02.03.2026 (1).xlsx`
- `Danh mục khóa học`: upsert theo `Mã khóa học`
- `Lớp đào tạo`: thay thế theo `Mã lớp`
- `Data raw học viên`: thay thế theo `Mã lớp`
- dữ liệu thật ngoài phạm vi demo sẽ được giữ nguyên
- `Staging nhân sự`, `Staging khóa học`, `Staging đào tạo`: không dùng trong bộ demo này

## Kết quả mong đợi
- `Bảng điều khiển`: có số liệu mới
- `Phân tích phòng ban`: xuất hiện ít nhất 4 phòng ban
- `Phòng ban theo khóa`: thấy khóa `GDP-GSP 2026` trên nhiều phòng ban
- `Phân tích xu hướng`: có ít nhất 2 tháng `2026-03` và `2026-04`

## Ghi chú
- hệ thống nạp một snapshot nhân viên demo cố định từ file HR `GSH_HR Master Data_QTHC_02.03.2026 (1).xlsx`
- dữ liệu demo được gắn tag `[DEMO v3.2.3]`
- bản vá hiệu năng cho thao tác demo nằm trong `v3.2.4`
- bộ seed nhân viên demo cố định theo file HR mới nằm trong `v3.2.5`
- hàm editor-only `clearDemoDataset()` để dọn demo an toàn nằm trong `v3.2.6`
