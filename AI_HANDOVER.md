# 🧠 AI-TO-AI HANDOVER DOCUMENT: HR-NEXUS V4.0 (LOCAL ENGINE)

**Date of Handover:** 2026-04-12
**Current System State:** Phase 1 (Local POC) of Cloudization Strategy.
**Target Reader:** LLM / AI Assistant taking over the workspace.

---

## 1. HỆ THỐNG HR-NEXUS LÀ GÌ?
HR-NEXUS là một bộ công cụ báo cáo và đối soát dữ liệu đào tạo nội bộ (L&D).
Bộ Core ban đầu được viết hoàn toàn bằng **Google Apps Script (GAS)** nhúng trong Google Sheets, tuy nhiên đã gặp **Bottleneck về Timeout (6 phút)** và Rate Limit khi lượng Raw Data vượt mức 5.000 dòng.

Khách hàng (User) đã quyết định nâng cấp hệ thống lên cấu trúc Backend tách rời bằng **Python (Pandas, OpenPyXL)**. Hiện tại hệ thống đang ở **Phase 1: Local Python Engine**. File Excel nội bộ (Local) đóng vai trò Database tạm thời trước khi được mount hoàn toàn lên Google Cloud Run.

---

## 2. PROJECT ARCHITECTURE & FILE MAPPING

Bạn, AI Agent, cần làm quen với các file code quan trọng nhất trong thư mục `LD REPORT V2/`:

| Tên File | Vai trò Core | Chú thích cho AI |
| :--- | :--- | :--- |
| `local_excel_app.py` | UI/Frontend Tool | Tool giao diện Desktop (Tkinter) để Admin thao tác chạy tự động tại máy. Không chứa logic toán học. |
| `local_excel_runner.py` | Controller | Chịu trách nhiệm load/đọc file `.xlsx` bằng Pandas, điều hướng qua các module xử lý, sau đó ghi đè dữ liệu (Write-Back) bằng thư viện `openpyxl`. |
| `transform.py` | Data Core (Pipeline) | Chứa toàn bộ logic dọn dẹp (Sanitization) và biến đổi hình học Data (Pivots/GroupBys). Đây là linh hồn của hệ thống Báo cáo. |
| `local_fact_builder.py`| Reconciliation Engine | Bộ máy "Đối soát". Nó so sánh `Data raw học viên` với danh mục Master (Nhân viên, Khóa học, Lớp đào tạo). Trigger cờ `QA_FAILED` / `QA_WARNING` khi có dị thường. |
| `gsheet_sync.py` | API Layer | Module đồng bộ API trực tiếp lên Google Sheets (Gspread). Dùng để kết nối Local với Cloud (đã hoàn thiện cơ bản). |

---

## 3. NHỮNG LUẬT DATA (BUSINESS LOGIC) CỰC KỲ QUAN TRỌNG

Cảnh báo: Nếu bạn định sửa code ở các module xử lý (`transform.py`, `local_fact_builder.py`), bạn **PHẢI** tuân thủ các quy tắc sau:

### 3.1. Auto-Fill Master Data (Tính năng mới nhất)
* **Status:** Vừa triển khai thành công.
* **Logic:** Khi Admin cấu hình hệ thống bằng `Data raw học viên`, thường họ sẽ ghi thêm Email của "Thực tập sinh" vào Raw. Nhưng trong `Danh mục nhân viên`, Email của TTS bị bỏ trống.
* **Cơ chế:** Ở hàm `_build_local_training_sync()` trong `local_fact_builder.py`, sau khi dò map dữ liệu, nếu thấy **Raw có Email** nhưng **Master trống**, hệ thống sẽ cập nhật biến `employees_final` và đẩy thẳng về mảng `EMPLOYEES_SHEET` ở file `local_excel_runner.py` để Write-back vào Excel.
* **Lưu ý AI:** Đừng cản trở quy trình Write-Back Array của `EMPLOYEES_SHEET` ở file Runner.

### 3.2. Deterministic Hash Rules (Kế thừa từ GAS)
Để tương thích ngược với Database cũ, các khóa chính (ID) tự sinh dạng `SES-XXXX` và `RAW_ID` được băm bằng thuật toán SHA-256. 
* Hệ mã hóa dùng `json.dumps(obj, separators=(',', ':'), ensure_ascii=False)`.
* Dữ liệu đưa vào là list (hoặc dict) dạng string UTF-8 không dấu cách thừa.
* Script nằm ở: `transform.py` -> `_generate_gas_session_hash(...)`.

### 3.3. Xử lý "Số bị ép kiểu Float"
Pandas khi đọc tệp Excel (.xlsx) các mã nhân viên toàn số dài (vd `50003832`) thường bị ép kiểu sang dạng Float (`50003832.0`) nếu cột có ô trống.
* **Giải pháp:** Đã sử dụng `_sanitize_scalar` để ép `int(val)` trước khi string hóa. Không được dùng `astype(str)` bừa bãi kẻo làm chết khâu Reconciliation.
* Thêm vào đó, số `"0"` nhập từ Admin trong ô thẻ Email đã được `mask(eq("0"), "")` ép thành Blank tự động trên toàn line.

---

## 4. BỘ NHỚ LƯU TRỮ TRẠNG THÁI (BRAIN MEMORY)

Project được gắn với 1 framework quản lý memory nằm ở `.brain/`
* Lệnh `/save-brain` sẽ tự động tạo Handover.
* **`.brain/brain.json`**: Chứa thông tin Meta về cấu trúc, Schema, Tool paths (Static Info).
* **`.brain/session.json`**: Lịch sử công việc mới nhất. Bất kỳ AI nào khởi động session mới phải Check 2 file này đầu tiên để nắm ngữ cảnh. Nếu AI load lại Project này, lập tức scan nội dung thư mục `.brain`.

---

## 5. NHỮM VỤ / KẾ HOẠCH BỊ TREO TRONG TƯƠNG LAI (PHASE 2)
1. **Kiểm thử hiệu năng ngẫu nhiên:** Khi File Excel Test đầy Full 5000 records, kiểm tra xem OpenPyxl Save bị BottleNeck (Chạy quá 30 giây) không (Hiện tại trung bình < 3 giây).
2. **Cloud Run Jobs Deployment:** Đóng gói toàn bộ Script vào một `Dockerfile` và quăng lên Cloud. Sửa file GAS trong Google Sheet cũ gắn `UrlFetchApp` để call API của Cloud Run này. 

🚀 **Chúc bạn làm tốt phiên tiếp theo!** Hãy luôn quét qua `.brain/session.json`!
