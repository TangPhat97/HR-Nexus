"""
HR-NEXUS Cloud Run Job Entry Point
====================================
File này chạy trên Google Cloud Run Jobs, không cần màn hình hay giao diện.
Đọc cấu hình từ biến môi trường, xử lý báo cáo và thoát.

Biến môi trường bắt buộc:
  GS_SPREADSHEET_ID         - ID của Google Sheet cần xử lý
  GS_SERVICE_ACCOUNT_JSON   - Nội dung JSON của Service Account (từ Secret Manager)

Biến môi trường tùy chọn:
  FISCAL_YEAR               - Năm tài chính (mặc định: năm hiện tại)
"""
import datetime
import json
import os
import sys

# ─── Validate môi trường ngay từ đầu ────────────────────────────────────────

spreadsheet_id = os.environ.get("GS_SPREADSHEET_ID", "").strip()
creds_json_str  = os.environ.get("GS_SERVICE_ACCOUNT_JSON", "").strip()
fiscal_year     = os.environ.get("FISCAL_YEAR", str(datetime.datetime.now().year)).strip()

if not spreadsheet_id:
    print("❌ LỖI: Thiếu biến môi trường GS_SPREADSHEET_ID")
    sys.exit(1)

if not creds_json_str:
    print("❌ LỖI: Thiếu biến môi trường GS_SERVICE_ACCOUNT_JSON (nội dung JSON của Service Account)")
    sys.exit(1)

# ─── Validate JSON credentials hợp lệ ──────────────────────────────────────

try:
    creds_data = json.loads(creds_json_str)
    if creds_data.get("type") != "service_account":
        print("❌ LỖI: GS_SERVICE_ACCOUNT_JSON không phải Service Account hợp lệ")
        sys.exit(1)
except json.JSONDecodeError as e:
    print(f"❌ LỖI: GS_SERVICE_ACCOUNT_JSON không phải JSON hợp lệ: {e}")
    sys.exit(1)

# ─── Ghi file cấu hình tạm thời ─────────────────────────────────────────────

CREDS_PATH  = "/tmp/service_account.json"
CONFIG_PATH = "/tmp/sync_config_cloud.json"

try:
    with open(CREDS_PATH, "w", encoding="utf-8") as f:
        json.dump(creds_data, f, ensure_ascii=False)

    config_data = {
        "spreadsheet_id": spreadsheet_id,
        "fiscal_year": int(fiscal_year),
        "credentials_path": CREDS_PATH,
        "backup_before_push": False,   # Cloud không có file Excel local để backup
    }
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config_data, f, ensure_ascii=False)

except OSError as e:
    print(f"❌ LỖI: Không thể ghi file tạm thời vào /tmp: {e}")
    sys.exit(1)

# ─── Kích hoạt Sync Engine ───────────────────────────────────────────────────

print("🚀 HR-NEXUS Cloud Engine đang khởi động...")
print(f"   📆 Năm tài chính : {fiscal_year}")
print(f"   🔗 Spreadsheet ID: {spreadsheet_id}")
print(f"   👤 Service Account: {creds_data.get('client_email', 'N/A')}")
print()

try:
    from gsheet_sync import SyncResult, run_sync

    result: SyncResult = run_sync(config_path=CONFIG_PATH, logger=print)

except ImportError as e:
    print(f"❌ LỖI IMPORT: Thiếu module - {e}")
    print("   → Kiểm tra lại Dockerfile có COPY đầy đủ file không?")
    sys.exit(1)

except Exception as e:
    print(f"💥 LỖI KHÔNG XÁC ĐỊNH: {e}")
    sys.exit(1)

# ─── Kết quả cuối ────────────────────────────────────────────────────────────

if result.success:
    print()
    print("✅ HOÀN TẤT: Báo cáo đã được cập nhật thành công lên Google Sheets!")
    sys.exit(0)
else:
    print()
    print("⚠️ THẤT BẠI: Quá trình xử lý gặp lỗi:")
    for err in result.errors:
        print(f"   • {err}")
    sys.exit(1)
