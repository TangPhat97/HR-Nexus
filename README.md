# Hệ Thống L&D Master Data trên Google Sheets + Google Apps Script

Đây là bộ mã nguồn hiện tại của hệ thống `MASTER DATA L&D`, vận hành trên:
- `Google Sheets`
- `Google Apps Script`
- quản trị source bằng `clasp + Git`

Phiên bản hiện tại: `v3.2.8`

Hệ thống đã được tái thiết kế từ workbook analytics cũ sang mô hình vận hành rõ ràng hơn, gồm:
- `Danh mục nhân viên`
- `Danh mục khóa học`
- `Lớp đào tạo`
- `Data raw học viên`
- `Dữ liệu đào tạo`
- `Dashboard + Analytics`
- `QA / Logs / Backup / Archive`

## Bộ tài liệu nên đọc theo thứ tự
- [SYSTEM_SUMMARY.md](D:\clean-architecture-canonical\file excel tu dong\SYSTEM_SUMMARY.md): bảng tổng hợp hệ thống, sheet, menu và luồng dữ liệu
- [HANDOVER_GUIDE.md](D:\clean-architecture-canonical\file excel tu dong\HANDOVER_GUIDE.md): tài liệu bàn giao cho `Admin`, `Operator`, `Dev/AI`
- [VERSION_CURRENT.md](D:\clean-architecture-canonical\file excel tu dong\VERSION_CURRENT.md): version hiện tại, ý nghĩa release và lịch sử patch `v3.2.x`
- [DEMO_DATA_GUIDE.md](D:\clean-architecture-canonical\file excel tu dong\DEMO_DATA_GUIDE.md): cách nạp và dọn dữ liệu demo
- [SAVEBRAIN.md](D:\clean-architecture-canonical\file excel tu dong\SAVEBRAIN.md): memory file cho dev/AI khi quay lại project
- [OPERATIONS_RUNBOOK.md](D:\clean-architecture-canonical\file excel tu dong\OPERATIONS_RUNBOOK.md): runbook vận hành
- [CHANGELOG.md](D:\clean-architecture-canonical\file excel tu dong\CHANGELOG.md): lịch sử thay đổi chính
- [ARCHITECTURE_CONTRACT.md](D:\clean-architecture-canonical\file excel tu dong\ARCHITECTURE_CONTRACT.md): contract kiến trúc và public entrypoints
- [DATA_DICTIONARY.md](D:\clean-architecture-canonical\file excel tu dong\DATA_DICTIONARY.md): từ điển dữ liệu
- [RBAC_MATRIX.md](D:\clean-architecture-canonical\file excel tu dong\RBAC_MATRIX.md): ma trận quyền

## Cấu trúc thư mục chính
- `gas/`: toàn bộ mã Apps Script và HTML sidebar
- `docs/`: tài liệu lịch sử từ bản cũ, chỉ dùng để tham khảo
- `skills/`: tài nguyên nghiên cứu skill/UI dùng trong quá trình tái thiết kế
- `plans/`: kế hoạch cũ, không phải source of truth của bản hiện tại
- `.brain/`: metadata làm việc nội bộ, không phải tài liệu bàn giao chính
- `backups/`: backup cục bộ của workspace

## Luồng vận hành chuẩn
1. Chạy `setupSystem()`
2. Kiểm tra `Hướng dẫn nhập liệu`
3. Nạp `Danh mục nhân viên`
4. Nạp `Danh mục khóa học`
5. Tạo `Lớp đào tạo`
6. Paste `Data raw học viên`
7. Chạy `refreshAnalytics()` hoặc menu `Làm mới báo cáo`
8. Kiểm tra `Bảng điều khiển`, `Phân tích phòng ban`, `Phòng ban theo khóa`

## Public functions chính
- `setupSystem()`
- `openStartHere()`
- `openHandoverGuide()`
- `openInputGuide()`
- `loadDemoDataset()`
- `clearDemoDataset()` `editor-only`, không đưa vào menu để tránh bấm nhầm
- `openTrainingEntry()`
- `openTrainingSessionEntry()`
- `openImportCenter()`
- `runQaChecks()`
- `publishStaging()`
- `refreshAnalytics()`
- `backupSystem()`
- `archiveYear()`
- `restoreBackup()`
- `runFullTestSuite()`

## Lưu ý vận hành hiện tại
- `RBAC_ENABLED` đang để `false` trong config để thuận tiện test nội bộ.
- `Data raw học viên` là sheet nhập bulk chính.
- `Dữ liệu đào tạo` là fact table do hệ thống sinh ra, không nên nhập tay.
- Các sheet dashboard và analytics là output-only.
- Nếu import `.xlsx`, Apps Script cần bật `Advanced Drive Service`.
- `clearDemoDataset()` chỉ dùng trong Apps Script editor.

## Nguồn sự thật của bản hiện tại
Khi cần hiểu hệ thống hiện tại, ưu tiên đọc theo thứ tự:
1. [README.md](D:\clean-architecture-canonical\file excel tu dong\README.md)
2. [SYSTEM_SUMMARY.md](D:\clean-architecture-canonical\file excel tu dong\SYSTEM_SUMMARY.md)
3. [VERSION_CURRENT.md](D:\clean-architecture-canonical\file excel tu dong\VERSION_CURRENT.md)
4. [SAVEBRAIN.md](D:\clean-architecture-canonical\file excel tu dong\SAVEBRAIN.md)
5. thư mục [gas](D:\clean-architecture-canonical\file excel tu dong\gas)
## Cap nhat van hanh 2026-03-23
- `Lam moi bao cao` mac dinh chay quick sync cho `Data raw hoc vien -> Du lieu dao tao`; co the chay nhieu lan sau khi sua ma/dong du lieu.
- Them menu `Dong bo lai toan bo du lieu raw` de full rebuild khi can doi soat hoac cuu ho.
- Nhan su da nghi viec van nen giu `row_status = ACTIVE` va doi `employment_status = Nghi viec` de bao toan lich su dao tao.
- `Danh muc nhan vien` da bo sung cot `Dia ban`; gia tri nay duoc snapshot sang `Du lieu dao tao` de san sang cho filter/report sau.

## Cap nhat van hanh 2026-03-23 (2 buoc)
- Nguoi dung cuoi nen chay Lam moi data raw truoc de cap nhat ma raw, ma bam dong, ghi chu va trang thai cho Data raw hoc vien.
- Sau khi raw da on, chay Dong bo du lieu dao tao de cap nhat Du lieu dao tao, Phan tich phong ban va Phong ban theo khoa.
- efreshAnalytics() hien duoc giu lai de tuong thich nguoc va se goi luong dong bo du lieu dao tao.

