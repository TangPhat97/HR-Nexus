# VERSION CURRENT

## 1. Version hiện tại
- `System version`: `v3.2.8`
- `Stack`: `Google Sheets + Google Apps Script`
- `Apps Script project`: `MASTER DATA L&D`

## 2. Ý nghĩa của `v3.2.8`
`v3.2.8` là bản hiện tại của nhánh workbook vận hành L&D trên Google Sheets, với các năng lực chính:
- migrate mềm từ workbook đã có dữ liệu
- `Lớp đào tạo`
- `Data raw học viên`
- `Phòng ban theo khóa`
- `Hướng dẫn nhập liệu`
- sidebar `Thông tin bàn giao`
- menu `Nạp dữ liệu demo`
- sidebar `Lớp đào tạo nhanh`
- tối ưu hiệu năng khi `loadDemoDataset()` và `refreshAnalytics()`
- seed nhân viên demo cố định theo file `GSH_HR Master Data_QTHC_02.03.2026 (1).xlsx`
- hàm editor-only `clearDemoDataset()` để dọn bộ demo an toàn
- `soft-migrate schema` cho các sheet dashboard / analytics do hệ thống tự quản lý
- test nhanh tự self-heal schema cho dashboard / analytics trước khi chấm pass/fail
- dropdown / checkbox / conditional formatting cho sheet nhập liệu

## 3. Chính sách tăng version
- tăng `patch` khi sửa bug, tài liệu, wording, test, UX nhỏ hoặc tối ưu nội bộ
- tăng `minor` khi thêm sheet, sidebar, report hoặc workflow mới nhưng không phá dữ liệu cũ
- tăng `major` khi đổi schema lõi, đổi thứ tự cột, hoặc bắt buộc migrate workbook

## 4. Lịch sử patch quan trọng trong `v3.2.x`
### `v3.2.0`
- Thêm `Lớp đào tạo`
- Thêm `Data raw học viên`
- Thêm `Phòng ban theo khóa`
- Mở rộng analytics theo phòng ban và khóa học

### `v3.2.1`
- Thêm `Hướng dẫn nhập liệu`
- Thêm `SAVEBRAIN.md`, `SYSTEM_SUMMARY.md`, `HANDOVER_GUIDE.md`, `VERSION_CURRENT.md`
- Gắn note hướng dẫn trên header sheet nhập liệu

### `v3.2.2`
- Thêm sidebar `Thông tin bàn giao`
- Thêm menu mở thông tin bàn giao

### `v3.2.3`
- Thêm `loadDemoDataset()`
- Sinh bộ demo `5 khóa học + 5 lớp + 25 raw`
- Cập nhật `runFullTestSuite()` để nhận diện public function mới

### `v3.2.4`
- Tối ưu `loadDemoDataset()` và `refreshAnalytics()`
- Giảm re-apply UX sheet trong luồng refresh để hạn chế treo

### `v3.2.5`
- Seed `Danh mục nhân viên` demo từ snapshot HR cố định
- Demo không còn phụ thuộc vào dữ liệu HR đang có sẵn trong workbook

### `v3.2.6`
- Thêm `clearDemoDataset()` chỉ chạy từ Apps Script editor
- Dọn demo an toàn mà không đụng `Danh mục nhân viên`

### `v3.2.7`
- Thêm `soft-migrate schema` cho các sheet output hệ thống
- `clearDemoDataset()` và `refreshAnalytics()` không còn bị chặn bởi schema cũ của dashboard / analytics

### `v3.2.8`
- `runFullTestSuite()` tự gọi cơ chế self-heal cho các sheet output trước khi chấm schema
- Giảm false negative ở `Schema bảng điều khiển`

## 5. Trạng thái module hiện tại
| Module | Trạng thái | Ghi chú |
|---|---|---|
| `Constants.gs` | active | source of truth cho schema, labels, version |
| `Repository.gs` | active | bootstrap, ensure schema, soft-migrate output sheets |
| `SecurityService.gs` | active | RBAC đang tạm tắt để test |
| `GuideService.gs` | active | tài liệu sống trong workbook |
| `SheetUxService.gs` | active | dropdown, checkbox, formatting |
| `EmployeeService.gs` | active | employee master |
| `CourseCatalogService.gs` | active | course master |
| `TrainingSessionService.gs` | active | session / lớp đào tạo |
| `TrainingRecordService.gs` | active | sync raw sang fact table |
| `ValidationService.gs` | active | QA / validation |
| `ImportService.gs` | active | import file |
| `AnalyticsService.gs` | active | dashboard + analytics |
| `BackupService.gs` | active | backup |
| `ArchiveService.gs` | active | archive |
| `AppController.gs` | active | public entrypoints |
| `MenuService.gs` | active | menu UI |

## 6. Compatibility notes
- `MASTER_EMPLOYEES`, `MASTER_COURSES`, `TRAINING_SESSIONS`, `TRAINING_RAW_PARTICIPANTS`, `TRAINING_RECORDS` là các sheet đầu vào / dữ liệu lõi, không nên đổi thứ tự cột tùy ý
- Các sheet output như `DASHBOARD_EXEC`, `DASHBOARD_OPERATIONS`, `ANALYTICS_DEPARTMENT`, `ANALYTICS_DEPARTMENT_COURSE` đã có cơ chế soft-migrate schema
- `clearDemoDataset()` không xóa seed HR demo trong `Danh mục nhân viên`
- `RBAC_ENABLED=false` là giả định test hiện tại, không phải trạng thái production

## 7. Đã hoàn thành
- setup workbook topology
- alias rename cho tab cũ
- tô màu tab
- Việt hóa UI chính
- hướng dẫn nhập liệu trong workbook
- sync `Data raw học viên -> Dữ liệu đào tạo`
- analytics theo phòng ban và phòng ban theo khóa
- bộ demo có thể nạp và dọn riêng

## 8. Chưa hoàn thành hoàn toàn
- import chuyên biệt từ file `Course List 2024/2025`
- bật lại RBAC production
- báo cáo chi phí / ISO theo lớp ở mức hoàn chỉnh
- migration test đầy đủ trên workbook legacy

## 9. Khi nào cần tăng version tiếp
- `v3.2.9+` nếu chỉ sửa bug, wording, tài liệu, test, UX nhỏ
- `v3.3.0` nếu thêm report/import/sidebar đáng kể mà không phá schema cũ
- `v4.0.0` nếu thay đổi schema lõi hoặc bắt buộc migrate workbook
