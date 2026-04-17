# Hướng Dẫn Thiết Lập Hệ Thống HR-NEXUS (Data Report)
*Tài liệu bàn giao dành cho Đội ngũ Kỹ thuật / IT*

Hệ thống HR-NEXUS là một kiến trúc Hybrid Cloud, kết hợp ưu điểm giao diện trực quan của Google Sheets và sức mạnh xử lý khối lượng dữ liệu lớn của Python & Google Cloud Run.

Để thiết lập hệ thống từ trạng thái trắng (sạch code nội bộ cũ), hãy thực hiện tuần tự 3 giai đoạn sau:

---

## Giai đoạn 1: Chuẩn bị Google Cloud Platform (GCP)
Trái tim của hệ thống là `cloud_job_main.py` và `transform.py` được đóng gói qua Docker chạy trên Google Cloud Run Jobs.

1. **Tạo Project mới trên GCP**: Vào Google Cloud Console tạo 1 Project mới, nhận `PROJECT_ID`.
2. **Kích hoạt các API cần thiết**:
   - Google Drive API
   - Google Sheets API
   - Cloud Run Admin API
   - Artifact Registry API
3. **Tạo Service Account (SA)**:
   - Tạo Service Account mới (ví dụ: `hr-nexus-sa@<your-project-id>.iam.gserviceaccount.com`).
   - Tạo File Key định dạng `.json` và tải về máy.
   - **QUAN TRỌNG:** Phải sử dụng email của Service Account này share quyền Editor vào file Google Sheets chứa Bảng điều khiển (Dashboard).
4. **Deploy Cloud Run Job**:
   - Mở terminal tại thư mục gốc của project (nơi chứa `Dockerfile`).
   - Chạy lệnh submit build lên GCP:
     ```bash
     gcloud builds submit --tag gcr.io/<PROJECT_ID>/ld-report-image
     ```
   - Tạo Cloud Run Job từ image vừa build:
     ```bash
     gcloud run jobs create <JOB_NAME> --image gcr.io/<PROJECT_ID>/ld-report-image --region <REGION> --task-timeout 10m
     ```

---

## Giai đoạn 2: Cấu hình Local Project
Nếu muốn chạy báo cáo trực tiếp từ máy (Local) cho quá trình kiểm thử hoặc chạy server on-premise:

1. Copy file key JSON của Service Account (đã tạo ở Giai đoạn 1) đưa vào thư mục `credentials/`. Đổi tên hoặc giữ nguyên tùy ý.
2. Mở file `sync_config.json` và sửa lại các biến:
   - `spreadsheet_id`: Lấy ID từ URL của file Google Sheets đích.
   - `credentials_path`: Đường dẫn tới file JSON ở bước trên (vd: `credentials/key.json`).
   - `gcp_project_id`: Project ID mới tạo ở GCP.
3. Chạy `pip install -r requirements.txt`.
4. Nếu chỉ định chạy mượt trên Local Excel App bằng giao diện: khởi động app bằng lệnh `python local_excel_app.py` hoặc bấm vào file `.cmd` có sẵn.

---

## Giai đoạn 3: Triển khai UI Lên Google Sheets (Apps Script)
Để người dùng tự bấm nút đồng bộ trên Google Sheets mà không cần thao tác code:

1. **Push thư mục `gas` lên Apps Script:**
   - Cài đặt Google Clasp: `npm install -g @google/clasp`
   - Đăng nhập: `clasp login`
   - Mở terminal tại thư mục gốc và tạo một project Apps Script mới đính kèm với Spreadsheet đích:
     ```bash
     clasp create --type sheets --parentId [YOUR_SPREADSHEET_ID] --rootDir ./gas
     ```
   - Đẩy toàn bộ codebase `*.gs`, `*.html` lên Sheets:
     ```bash
     clasp push
     ```
2. **Khai báo Nút bấm Cloud:**
   - Sau khi đẩy script lên, mở file `gas_cloud_trigger.js` ở máy tính (nằm ngoài thư mục `gas/`).
   - Chỉnh sửa dòng `PROJECT_ID`, `LOCATION`, `JOB_NAME` khớp với thông tin Giai đoạn 1.
   - Mở Extensions -> Apps Script trên Google Sheets, Dán nội dung của `gas_cloud_trigger.js` vào file Script mới lưu tên `CloudTrigger.gs`. Đảm bảo gọi hàm `addCloudMenuItem(menu)` như hướng dẫn nội tuyến (inline guide) trong file.
3. Xin quyền (Authorize) tập lệnh lần đầu và sử dụng thành công qua Custom Menu.

---
**Chúc các bạn triển khai thành công!** Nếu có phát sinh, vui lòng check log trên Google Cloud Console > Cloud Run Jobs > Logs để khắc phục.
