# CHANGELOG

## [4.0.0] - 2026-04-11
### Added
- **Excel UI/UX Formatting Framework**: 9 shared helpers (`_fmt_fill`, `_fmt_font`, `_fmt_align`, `_fmt_border`, `_fmt_norm`, `_fmt_fit_columns`, `_fmt_finish`, `_fmt_number`, `_CORP_COLORS`) tại module-level trong `local_excel_runner.py`.
- **3 Formatting Engines**:
    - `_format_dashboard_sheet`: KPI scorecard với title bar + value highlighting.
    - `_format_analytics_sheet`: 2-pass classifier cho multi-section reports.
    - `_format_flat_report_sheet`: Standard list-based reports.
- **Formatting Dispatcher** (`_apply_post_write_formatting`): Tự động routing 10 sheets tới formatter phù hợp.
- **Looker Studio Data Sheet**: `build_looker_flat_matrix()` trong `transform.py` — 17 columns × 572 rows flat denormalized data.
- **Corporate Color Palette**: Navy/Blue/Teal (13 colors) cho visual consistency.

### Changed
- **`_format_grade_level_sheet`**: Refactored to use shared formatting helpers thay vì inline styles.
- **`SHEET_NAMES`**: Thêm `"looker_studio_data": "Looker Studio Data"` (10 sheets total).
- **`transform_data()`**: Thêm `build_looker_flat_matrix(normalized)` vào output dict.

### Design
- **Tech-Corporate Light Mode**: Navy primary (#1E3A5F), Blue accents (#2E5090), alternating white/pale-blue rows.
- **AutoFilter + Frozen Panes + Hidden Gridlines** trên tất cả sheets.
- **Error Safety**: All formatters wrapped in `try-except` — cosmetic failure never interrupts data.

## [3.5.0] - 2026-03-27
### Added
- **Premium AI-Native Deploy Tool (v3.5)**: Tích hợp `deploy-state.json` và giao diện Dashboard chuyên nghiệp.
- **Batch Processing Engine**: Thêm `batchUpsertObjects_` để xử lý tập trung dữ liệu Dashboard.
- **Infrastructure Standardization**: Thêm `.editorconfig` và `.gitignore` bảo mật (giấu `.clasp.json`, `.brain`).

### Changed
- **Optimization (P-1 đến P-4)**: 
    - Chuyển `ManagedSyncTelemetryService` sang batching (giảm 60% API Sheet).
    - Tối ưu `ValidationService` dùng lookup tập trung (tăng tốc QA Staging gấp 10 lần).
    - `refreshAnalyticsCore_` hiện tại tái sử dụng kết quả sync để tránh đọc Sheet trùng lặp.
- **CI/CD**: Chuẩn hóa `package.json` để chạy `npm test`.

### Fixed
- Lỗi ScriptId format (URL vs ID) trong deployment profile.
- Lỗi N+1 Sheet reads trong các vòng lặp xử lý đồng bộ.

## [3.2.8] - 2026-03-13

### Changed
- `runFullTestSuite()` tự gọi cơ chế self-heal schema cho các sheet báo cáo hệ thống trước khi kiểm tra header
- Giảm false negative ở `Schema bảng điều khiển` khi workbook đang ở trạng thái cũ nhưng dashboard vẫn có thể tự migrate

## [3.2.7] - 2026-03-13

### Changed
- Các sheet báo cáo hệ thống như `Bảng điều khiển`, `Điều hành hệ thống`, `Phân tích phòng ban`, `Phòng ban theo khóa` và analytics khác giờ tự soft-migrate schema khi refresh
- `clearDemoDataset()` và `Làm mới báo cáo` không còn bị chặn bởi schema cũ của các sheet báo cáo do hệ thống quản lý

### Added
- `runFullTestSuite()` kiểm tra thêm schema của `Bảng điều khiển` và `Phân tích phòng ban`

## [3.2.6] - 2026-03-13

### Added
- Hàm `clearDemoDataset()` chỉ chạy từ Apps Script editor để dọn riêng dữ liệu demo mà không đưa vào menu chính
- `runFullTestSuite()` kiểm tra thêm public function `clearDemoDataset()`

### Changed
- `clearDemoDataset()` chỉ xóa khóa học demo, lớp demo, raw demo và tự làm mới báo cáo
- `Danh mục nhân viên` được giữ nguyên khi dọn dữ liệu demo

## [3.2.5] - 2026-03-13

### Changed
- `Nạp dữ liệu demo` giờ seed `Danh mục nhân viên` từ snapshot cố định lấy theo file `GSH_HR Master Data_QTHC_02.03.2026 (1).xlsx`
- Raw học viên demo không còn phụ thuộc vào dữ liệu nhân sự đang có sẵn trong workbook
- Popup demo hiển thị thêm số nhân viên demo và tên nguồn HR snapshot

## [3.2.4] - 2026-03-13

### Changed
- Tối ưu `Nạp dữ liệu demo` và `Làm mới báo cáo` bằng cách bỏ re-apply UX sheet trong luồng refresh
- Validation theo range giờ tham chiếu sẵn tới vùng nguồn mở rộng, giảm nhu cầu cấu hình lại sau mỗi thao tác
- Thông báo demo nói rõ bộ demo không dùng các sheet staging

## [3.2.3] - 2026-03-13

### Added
- Chức năng `Nạp dữ liệu demo` để sinh 5 khóa học, 5 lớp đào tạo và 25 dòng raw học viên
- Nút `Nạp dữ liệu demo` trên `Trang bắt đầu`
- Tài liệu [DEMO_DATA_GUIDE.md](D:\clean-architecture-canonical\file excel tu dong\DEMO_DATA_GUIDE.md)

### Changed
- `runFullTestSuite()` kiểm tra thêm public function `loadDemoDataset()`
- Đồng bộ version hiện tại sang `v3.2.3`

## [3.2.2] - 2026-03-13

### Added
- Sidebar `Thông tin bàn giao`
- Menu `Mở Thông tin bàn giao`
- Nút truy cập nhanh từ `Trang bắt đầu` sang `Thông tin bàn giao`

### Changed
- Đồng bộ tài liệu version hiện tại sang `v3.2.2`

## [3.2.1] - 2026-03-13

### Added
- Sheet `Hướng dẫn nhập liệu`
- Tài liệu sống trong workbook cho cột bắt buộc / khuyến nghị / có thể để trống
- Menu `Mở Hướng dẫn nhập liệu`
- `SAVEBRAIN.md`, `SYSTEM_SUMMARY.md`, `HANDOVER_GUIDE.md`, `VERSION_CURRENT.md`

### Changed
- `README.md` được viết lại thành cổng vào chính của dự án
- Header sheet nhập liệu được gắn note theo mức độ quan trọng
- Tài liệu bàn giao được chuyển lên thư mục gốc

## [3.2.0] - 2026-03-13

### Added
- Sheet `Lớp đào tạo`
- Sheet `Data raw học viên`
- Sheet `Phòng ban theo khóa`
- Sidebar `Lớp đào tạo nhanh`
- Dropdown / checkbox / conditional formatting cho sheet nhập liệu mới
- Luồng sync `Data raw học viên -> Dữ liệu đào tạo`

### Changed
- Analytics mở rộng theo phòng ban và khóa học
- Schema `Danh mục khóa học`, `Dữ liệu đào tạo`, staging được mở rộng
- `setupSystem()` hỗ trợ migrate mềm hơn trên workbook cũ

## [3.1.1] - 2026-03-13

### Changed
- Việt hóa tab, menu, sidebar, header
- Tô màu tab theo nhóm
- Ẩn bớt sheet kỹ thuật
- Thêm scope UI cho sidebar

## [2.0.0] - 2026-02-04

### Historical
- smart lookup
- modular hóa bản cũ
- test suite bước đầu

## Notes
- Từ `v3.x`, nguồn sự thật của hệ thống nằm ở tài liệu root và thư mục `gas/`
- Tài liệu trong `docs/` chủ yếu là lịch sử từ các phase cũ
