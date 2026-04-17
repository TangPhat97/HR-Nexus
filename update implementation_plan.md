# 🔬 So Sánh Kiến Trúc Đồng Bộ: Local vs Cloud Run

## Bối cảnh

HR-NEXUS cần "cầu nối" giữa Google Sheets (UI/nhập liệu) và Python Engine (xử lý nặng). Có 4 hướng tiếp cận khả thi. Phân tích dưới đây giúp anh chọn đúng.

---

## Ma Trận So Sánh Toàn Diện

### 4 Phương án

```
Option A: LOCAL PYTHON + gspread
Admin bấm lệnh trên máy → Python kéo data từ Sheet → xử lý → đẩy lên lại

Option B: CLOUD RUN JOBS + Cloud Scheduler  
Cloud Scheduler đặt giờ → Cloud Run bật container → Python xử lý → cập nhật Sheet → tự tắt

Option C: CLOUD RUN + GAS Trigger (nút bấm)
Admin bấm nút trên Sheet → GAS gửi webhook → Cloud Run bật lên → xử lý → cập nhật Sheet

Option D: STAGED (A → B) — HYBRID
Làm Option A trước → khi ổn định, Docker hóa + deploy lên Cloud Run
```

### Sơ đồ kiến trúc

```
OPTION A — Local Python                       OPTION B — Cloud Run + Scheduler
┌──────────┐   gspread   ┌──────────┐         ┌──────────┐  API   ┌─────────────┐
│  Google   │◄──────────►│  Local   │         │  Google   │◄─────►│  Cloud Run  │
│  Sheets   │  pull/push │  Python  │         │  Sheets   │       │  Job        │
│  (UI)     │            │  Engine  │         │  (UI)     │       │  (Docker)   │
└──────────┘            └──────────┘         └──────────┘       └──────┬──────┘
                         Admin bấm CMD                                  │
                                                            ┌──────────┴──────┐
                                                            │ Cloud Scheduler │
                                                            │  (2h sáng/ngày)  │
                                                            └─────────────────┘

OPTION C — Cloud Run + GAS (Nút bấm)          OPTION D — Staged (A rồi B)
┌──────────┐ webhook ┌─────────────┐          ┌───────────────────────────────┐
│  Google   │───────►│  Cloud Run  │          │  Phase 1: Local Python (A)   │
│  Sheets   │◄──────│  Job        │          │  ● Phát triển & test nhanh   │
│  (UI)     │  API  │  (Docker)   │          │  ● Chạy ổn định             │
└──────────┘       └─────────────┘          │         ↓ Docker hóa         │
     │                                       │  Phase 2: Cloud Run (B/C)    │
     │  Admin bấm Menu                       │  ● Tự động theo lịch        │
     └──► GAS: UrlFetchApp.fetch()           │  ● Không cần máy local      │
                                              └───────────────────────────────┘
```

---

### Bảng so sánh 12 tiêu chí

| # | Tiêu chí | 🅰️ Local gspread | 🅱️ Cloud Run + Scheduler | 🅲️ Cloud Run + GAS | 🅳️ Staged A→B |
|---|----------|:-----------------:|:------------------------:|:-------------------:|:-------------:|
| 1 | **Chi phí** | **$0** | ~$0.05/tháng | ~$0.05/tháng | $0 → $0.05 |
| 2 | **Thời gian setup** | **~2 giờ** | ~8 giờ | ~10 giờ | 2h + 4h sau |
| 3 | **Cần máy tính bật?** | ✅ Cần | ❌ Không | ❌ Không | Cần → Không |
| 4 | **Tự động theo lịch** | ❌ Thủ công | ✅ Scheduler | ✅ Nút bấm | ❌ → ✅ |
| 5 | **Tốc độ xử lý** | ⚡ Nhanh | ⚡ Nhanh | ⚡ Nhanh | ⚡ |
| 6 | **Bảo mật** | 🟡 Key local | 🟢 Secret Mgr | 🟢 Secret Mgr | 🟡 → 🟢 |
| 7 | **Scale (>5000 rows)** | 🟢 OK | 🟢 OK | 🟢 OK | 🟢 |
| 8 | **DevOps effort** | **Không** | Docker + GCP | Docker + GCP | Không → có |
| 9 | **Offline capability** | ✅ Có | ❌ Không | ❌ Không | Có → Không |
| 10 | **Reuse code hiện tại** | **100%** | **100%** | **100%** | 100% |
| 11 | **Nhiều người dùng** | 1 người | Không giới hạn | Không giới hạn | 1 → N |
| 12 | **Học thêm gì?** | gspread | Docker, GCP IAM, Cloud Run | + GAS webhook | Từ từ |

---

### Chi tiết từng phương án

#### 🅰️ Option A: Local Python + gspread

```python
# Luồng đơn giản nhất
python gsheet_sync.py sync

# Bên trong:
# 1. gspread kết nối Google Sheets
# 2. Đọc 11 sheet nguồn → DataFrame
# 3. transform_data() xử lý → 10 matrices
# 4. gspread ghi lên lại 10 sheet
# 5. Xong! ~30 giây
```

**Ưu điểm:**
- Làm được NGAY — chỉ cần `pip install gspread` + Service Account JSON
- Code engine giữ nguyên 100%, chỉ thêm 1 file `gsheet_sync.py`
- Debug dễ, chạy local thấy kết quả ngay
- $0 chi phí

**Nhược điểm:**
- Phải mở máy tính + bấm lệnh thủ công
- Chỉ 1 người chạy được (ai có credentials file)
- Không tự động được

---

#### 🅱️ Option B: Cloud Run Jobs + Cloud Scheduler

```
Cloud Scheduler (2h sáng mỗi ngày)
    → Gọi Cloud Run Jobs API
        → Container Docker thức dậy
            → Python: gspread đọc Sheet → xử lý → ghi lại
        → Container tự tắt ($0 khi không chạy)
```

**Ưu điểm:**
- ✅ Hoàn toàn tự động — không cần ai bấm
- ✅ Chạy bất kể máy tính có bật hay không
- ✅ Credentials an toàn trong Secret Manager
- ✅ Logs tập trung trên Cloud Logging

**Nhược điểm:**
- Cần biết Docker (viết Dockerfile, build image)
- Cần setup GCP Project: IAM, Secret Manager, Cloud Run, Scheduler
- Debug khó hơn (phải xem logs trên GCP Console)
- Mất ~8 giờ setup lần đầu

**Chi phí (ước tính HR-NEXUS):**
```
Cloud Run:   ~572 rows × 10 sheets = ~30 giây xử lý
             0.5 vCPU × 30s = 15 vCPU-seconds
             Free tier: 180,000 vCPU-seconds/tháng
             → $0 (dư sức)

Scheduler:   1 job/ngày = 31 jobs/tháng
             Free tier: 3 jobs miễn phí
             → $0

Secret Mgr:  1 secret × 31 accesses = 31
             Free tier: 10,000 accesses
             → $0

TOTAL:       ~$0/tháng (nằm trong free tier)
```

---

#### 🅲️ Option C: Cloud Run + GAS Menu Trigger

```
Admin mở Google Sheet → Bấm Menu "HR-NEXUS > Chạy Báo Cáo"
    → GAS: UrlFetchApp.fetch(CLOUD_RUN_URL, {method: "POST"})
        → Cloud Run thức dậy → xử lý → ghi lại Sheet
    → Sheet hiện: "✅ Báo cáo đã cập nhật!"
```

**Ưu điểm:**
- Trải nghiệm người dùng TỐT NHẤT — bấm 1 nút trên Sheet quen thuộc
- Kết hợp Cloud Run (mạnh) + GAS (UI quen thuộc)

**Nhược điểm:**
- Phức tạp nhất: Docker + GCP + GAS webhook + Authentication
- GAS cần JWT token để gọi Cloud Run (thêm code auth)
- Nếu Cloud Run lâu >30s, GAS có thể timeout (cần async pattern)

---

#### 🅳️ Option D: Staged Approach (ĐỀ XUẤT)

```
           HIỆN TẠI                    ĐÃ LÀM               SẮP LÀM (Phase 1)           TƯƠNG LAI (Phase 2)
┌──────────────────────┐    ┌──────────────────────┐    ┌──────────────────────┐    ┌──────────────────────┐
│  GAS trên Sheets     │    │  Python + openpyxl   │    │  Python + gspread    │    │  Cloud Run Jobs      │
│  (v3.5 - timeout)    │    │  (v4.0 - Excel only) │    │  (v4.1 - Sheet sync) │    │  (v5.0 - full cloud) │
│                      │ →  │                      │ →  │                      │ →  │                      │
│  ❌ Bị giới hạn 6min │    │  ✅ Xử lý mạnh      │    │  ✅ Sync 2 chiều     │    │  ✅ Tự động + scale  │
│                      │    │  ❌ Chỉ offline Excel │    │  ❌ Cần máy bật      │    │  ✅ Không cần máy    │
└──────────────────────┘    └──────────────────────┘    └──────────────────────┘    └──────────────────────┘
```

**Tại sao Staged là tốt nhất cho HR-NEXUS?**

1. **Code engine KHÔNG thay đổi** khi chuyển A→B — chỉ thay lớp "transport" (gspread local → gspread trong Docker)
2. **Test local trước** — chắc chắn logic đúng rồi mới lên cloud
3. **Không mất thời gian setup GCP** nếu chưa cần tự động
4. **Khi lên Cloud Run**, chỉ cần:
   - Viết `Dockerfile` (5 dòng)
   - Copy `gsheet_sync.py` + engine vào container
   - Chạy `gcloud run jobs deploy`
   - Tạo Cloud Scheduler trigger

---

## Em đề xuất: 🅳️ STAGED APPROACH

### Lý do:

| Câu hỏi | Trả lời |
|----------|---------|
| Hiện tại có cần tự động không? | Chưa — Admin bấm lệnh là được |
| Có nhiều người cần chạy không? | Chưa — chỉ Admin |
| Code engine đã sẵn sàng chưa? | ✅ Rồi — 100% đã test |
| Cần gì để sync Sheet? | Chỉ cần thêm `gspread` layer |
| Khi nào cần Cloud Run? | Khi muốn tự động chạy 2h sáng / nhiều người dùng |

### Lộ trình:

```
Tuần này (2h):   Phase 1 — Option A (Local gspread sync)
                 ├── pip install gspread
                 ├── Tạo Service Account
                 ├── Viết gsheet_sync.py
                 └── Test: python gsheet_sync.py sync ✅

Khi cần (4h):    Phase 2 — Nâng lên Cloud Run
                 ├── Viết Dockerfile
                 ├── gcloud run jobs deploy
                 ├── Cloud Scheduler (2h sáng)
                 └── (Optional) GAS menu trigger
```

---

## Open Questions

> [!IMPORTANT]
> **Anh xác nhận:**
> 1. **Đi "Staged" (A trước, B sau)** hay **đi thẳng Cloud Run (B)?**
>    - Staged: làm ngay trong 2h, test local, lên cloud sau
>    - Thẳng B: setup GCP + Docker ngay, mất ~8h
>
> 2. **Có cần tự động chạy theo lịch không?** (VD: 2h sáng mỗi ngày)
>    - Nếu CÓ → Cloud Run + Scheduler (Phase 2)
>    - Nếu CHƯA → Local gspread đủ rồi (Phase 1)
>
> 3. **Mấy người cần chạy báo cáo?**
>    - 1 Admin → Local đủ
>    - Nhiều người → Cloud Run
