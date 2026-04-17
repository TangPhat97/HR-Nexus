# ☁️ HR-NEXUS Cloud Operations - Cheat Sheet

> **Dùng khi:** Sửa code Python xong, muốn cập nhật lên mây.
> **Thư mục:** `O:\clean-architecture-canonical\LD REPORT V2\LD REPORT V2`

---

## 🔄 Quy trình cập nhật code lên Cloud (3 bước)

```powershell
# Bước 1: Build image mới (khoảng 1 phút)
gcloud builds submit --tag asia-southeast1-docker.pkg.dev/[YOUR_PROJECT_ID]/ld-reports/ld-report-engine

# Bước 2: Cập nhật Cloud Job để dùng image mới
gcloud run jobs update ld-report-job --image=asia-southeast1-docker.pkg.dev/[YOUR_PROJECT_ID]/ld-reports/ld-report-engine:latest --region=asia-southeast1

# Bước 3: Chạy thử (hoặc bấm nút trên Google Sheet)
gcloud run jobs execute ld-report-job --region=asia-southeast1 --wait
```

---

## 🚀 Chạy báo cáo

| Cách | Lệnh / Thao tác |
|------|-----------------|
| **Google Sheet** | Menu ☁️ L&D Cloud Engine → 🚀 Chạy báo cáo Cloud |
| **Terminal** | `gcloud run jobs execute ld-report-job --region=asia-southeast1 --wait` |

---

## 🔍 Kiểm tra & Debug

```powershell
# Xem trạng thái job
gcloud run jobs describe ld-report-job --region=asia-southeast1

# Xem log lần chạy gần nhất
gcloud logging read "resource.type=cloud_run_job AND resource.labels.job_name=ld-report-job" --limit=30 --format="value(textPayload)" --project=[YOUR_PROJECT_ID]

# Xem chi tiết 1 execution cụ thể (thay ID)
gcloud run jobs executions describe <EXECUTION_ID> --region=asia-southeast1
```

---

## ⚙️ Thay đổi cấu hình

```powershell
# Đổi Google Sheet ID
gcloud run jobs update ld-report-job --region=asia-southeast1 --set-env-vars="GS_SPREADSHEET_ID=<SHEET_ID_MỚI>,FISCAL_YEAR=2026"

# Đổi năm tài chính
gcloud run jobs update ld-report-job --region=asia-southeast1 --set-env-vars="GS_SPREADSHEET_ID=[YOUR_SPREADSHEET_ID],FISCAL_YEAR=2027"
```

---

## 📋 Thông tin quan trọng

| Key | Value |
|-----|-------|
| **Project ID** | `[YOUR_PROJECT_ID]` |
| **Region** | `asia-southeast1` |
| **Job Name** | `ld-report-job` |
| **Image** | `asia-southeast1-docker.pkg.dev/[YOUR_PROJECT_ID]/ld-reports/ld-report-engine` |
| **Test Sheet ID** | `[YOUR_SPREADSHEET_ID]` |
| **Service Account** | `hr-nexus-sync@[YOUR_PROJECT_ID].iam.gserviceaccount.com` |
| **Secret** | `gs-service-account:latest` |

---

## ⚠️ Lưu ý khi chuyển sang Live Sheet

1. Share **Live Sheet** cho Service Account (Editor):
   ```
   hr-nexus-sync@[YOUR_PROJECT_ID].iam.gserviceaccount.com
   ```
2. Cập nhật Sheet ID:
   ```powershell
   gcloud run jobs update ld-report-job --region=asia-southeast1 --set-env-vars="GS_SPREADSHEET_ID=<LIVE_SHEET_ID>,FISCAL_YEAR=2026"
   ```
3. Cài code GAS vào Live Sheet (copy từ `gas_cloud_trigger.js`)
