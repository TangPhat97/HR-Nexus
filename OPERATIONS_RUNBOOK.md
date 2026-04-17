# OPERATIONS RUNBOOK

## 1. Khởi tạo lần đầu
1. mở workbook
2. refresh Apps Script mới nhất
3. chạy `setupSystem()`
4. mở `Hướng dẫn nhập liệu`
5. kiểm tra các tab chính đã xuất hiện

## 2. Quy trình nhập dữ liệu chuẩn
1. cập nhật `Danh mục nhân viên`
2. cập nhật `Danh mục khóa học`
3. tạo `Lớp đào tạo`
4. paste `Data raw học viên`
5. chạy `Làm mới báo cáo`

## 3. Quy trình import file
1. mở `Trung tâm import`
2. chọn loại import
3. upload file
4. chạy QA
5. sửa lỗi FAIL
6. publish
7. làm mới báo cáo

## 4. Kiểm tra nhanh sau mỗi lần cập nhật
- chạy `Chạy test nhanh`
- kiểm tra `Bảng điều khiển`
- kiểm tra `Phân tích phòng ban`
- kiểm tra `Phòng ban theo khóa`

## 5. Backup
- chạy `Sao lưu` trước thay đổi lớn
- ghi nhận URL backup

## 6. Archive năm
1. backup
2. chạy `Tạo archive năm`
3. kiểm tra file archive mới
4. xác nhận workbook production chỉ giữ dữ liệu hoạt động

## 7. Khi có lỗi
- lỗi data: xem `Kết quả QA`
- lỗi runtime: xem `Nhật ký lỗi`
- lỗi thao tác: xem `Nhật ký tác động`
- lỗi schema: kiểm tra `VERSION_CURRENT.md` và `SAVEBRAIN.md`

## 8. Lưu ý hiện tại
- RBAC đang tạm tắt để test
- import `.xlsx` cần `Advanced Drive Service`
