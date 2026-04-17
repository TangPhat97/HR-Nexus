# 🔬 Nghiên Cứu Chuyên Sâu & Kế Hoạch Nâng Cấp V2

## 1. KẾT QUẢ AUDIT TOÀN BỘ HỆ THỐNG

### 1.1. Kiến trúc hiện tại (v3.5.1)

```mermaid
graph LR
    A["Excel Workbook<br/>11 sheet nguồn"] --> B["local_excel_runner.py<br/>Load + Header Map"]
    B --> C["local_fact_builder.py<br/>QA + Sync Raw → Fact"]
    B --> D["transform.py<br/>Analytics Engine"]
    C --> D
    D --> E["8 sheet báo cáo<br/>ghi vào Excel"]
```

| Component | File | Dòng code | Tình trạng |
|-----------|------|-----------|------------|
| Data loader | `local_excel_runner.py` | 774 | ✅ Ổn định |
| QA/Sync engine | `local_fact_builder.py` | 580 | ⚠️ Đã fix bug ~bool |
| Analytics engine | `transform.py` | 2,282 | ⚠️ Thiếu grade/level |
| Test suite (GAS) | `raw-sync-regression.test.cjs` | 668 | ✅ PASS |
| Test suite (Python) | ❌ Không có | 0 | 🔴 **Thiếu hoàn toàn** |

### 1.2. Bug đã tìm thấy & fix

| # | Bug | Nguyên nhân | Trạng thái |
|---|-----|-------------|------------|
| 1 | 579 dòng fail dù data sạch | Python 3.12+ `~True` trả `-2` thay `False` khi dtype=object | ✅ **Đã fix** |

### 1.3. Gap Analysis — Dữ liệu `grade` / `level`

> [!IMPORTANT]
> **Phát hiện quan trọng:** Dữ liệu Ngạch và Cấp bậc **ĐÃ được map đúng** trong `local_excel_runner.py` (dòng 67-68 cho Employee, 154-155 cho Training Records), nhưng `transform.py` ***hoàn toàn không đọc*** 2 cột này.

**Dòng chảy dữ liệu hiện tại:**

```
Excel "Ngạch" → runner map thành "grade" → ✅ có trong DataFrame
Excel "Cấp bậc" → runner map thành "level" → ✅ có trong DataFrame
                                                    ↓
transform._normalize_training_records() → ❌ KHÔNG đọc "grade", "level"
                                                    ↓
Analytics sheets → ❌ Không có thông tin ngạch/bậc
```

### 1.4. Các sheet báo cáo hiện tại (8 sheets)

| # | Sheet Key | Tên tiếng Việt | Kích thước |
|---|-----------|----------------|------------|
| 1 | `dashboard_exec` | Bảng điều khiển | KPI tổng hợp |
| 2 | `dashboard_operations` | Điều hành hệ thống | Trạng thái sync |
| 3 | `analytics_course` | Phân tích khóa học | Group by course |
| 4 | `analytics_department` | Phân tích phòng ban | Group by dept |
| 5 | `analytics_department_course` | Phòng ban theo khóa | Dept × Course |
| 6 | `analytics_trend` | Phân tích xu hướng | Group by month |
| 7 | `report_external_assignment` | Cử đi học bên ngoài | External training |
| 8 | `report_session_reconciliation` | Đối chiếu lớp đào tạo | Session QA |

> [!TIP]
> **Nhận xét:** Hệ thống đã có phân tích theo Phòng ban, Khóa học, Xu hướng — nhưng **THIẾU** góc nhìn theo cấp bậc nhân sự. Đây chính xác là gap mà anh muốn giải quyết.

---

## 2. DANH SÁCH NÂNG CẤP CẦN THIẾT (V2)

### Phase 1: Fix & Normalize (transform.py)

> [!CAUTION]
> Phải fix trước khi thêm tính năng mới.

| # | Công việc | File | Ưu tiên |
|---|-----------|------|---------|
| 1.1 | Thêm `"grade"`, `"level"` vào `_normalize_training_records()` | transform.py | 🔴 Cao |
| 1.2 | Thêm `"grade"`, `"level"` vào `_normalize_employees()` | transform.py | 🔴 Cao |

**Chi tiết thay đổi:**

```diff
# _normalize_training_records() - line 444-479
 _prepare_frame(frame, [
     ...
     "job_title",
+    "grade",
+    "level",
     "session_id",
     ...
 ])

# .assign() - line 487-555
 df.assign(
     ...
     job_title=_sanitize_series(df["job_title"]),
+    grade_name=_first_nonblank([_sanitize_series(df["grade"])], "Không xác định"),
+    level_name=_first_nonblank([_sanitize_series(df["level"])], "Không xác định"),
     session_id=...
 )

# _normalize_employees() - line 558-566
-_prepare_frame(frame, ["row_status", "department", "emp_id", "email", "full_name"])
+_prepare_frame(frame, ["row_status", "department", "emp_id", "email", "full_name", "grade", "level"])
 df.assign(
     ...
+    grade_name=_first_nonblank([_sanitize_series(df["grade"])], "Không xác định"),
+    level_name=_first_nonblank([_sanitize_series(df["level"])], "Không xác định"),
 )
```

---

### Phase 2: Sheet mới — "Phân tích ngạch cấp bậc"

| # | Công việc | File |
|---|-----------|------|
| 2.1 | Thêm `"analytics_grade_level"` vào `SHEET_NAMES` | transform.py |
| 2.2 | Thêm gọi `build_grade_level_canvas()` vào `transform_data()` | transform.py |
| 2.3 | Tạo hàm `_build_grade_level_report_data()` | transform.py |
| 2.4 | Tạo hàm `build_grade_level_canvas()` | transform.py |
| 2.5 | Cập nhật `__all__` | transform.py |

**Layout sheet mới (6 section):**

| Section | Nội dung | Group by |
|---------|----------|----------|
| KPI | Tổng giờ, Số ngạch, Giờ TB/người | — |
| Tổng hợp theo Ngạch | Số NV, Lượt ĐT, Tổng giờ, Giờ TB, Điểm, Hài lòng | `grade` |
| Tổng hợp theo Cấp bậc | Tương tự | `level` |
| Chi tiết PB × Ngạch | So sánh giữa phòng | `department` × `grade` |
| Chi tiết PB × Cấp bậc | Tương tự | `department` × `level` |
| Top 50 Học viên | Ai học nhiều nhất, ở ngạch/bậc nào | Individual |

---

### Phase 3: TDD Test Suite cho Python (MỚI)

> [!IMPORTANT]
> Hiện hệ thống **KHÔNG CÓ** test Python nào cho `transform.py` (2,282 dòng code không test!). Chỉ có test GAS bằng Node.js. Cần tạo test suite riêng.

| # | Test file | Phạm vi |
|---|-----------|---------|
| 3.1 | `tests/test_transform_normalize.py` | Các hàm normalize (records, employees, sessions) |
| 3.2 | `tests/test_transform_grade_level.py` | Sheet mới: grade/level analytics |
| 3.3 | `tests/test_transform_existing.py` | Smoke tests cho 8 sheet hiện tại |
| 3.4 | `tests/test_fact_builder_boolean.py` | Regression test cho bug ~bool Python 3.12+ |

**Test strategy:**

```python
# TDD: Viết test TRƯỚC, code SAU
def test_grade_level_canvas_groups_by_grade():
    """Sheet mới phải group đúng theo ngạch."""
    records = pd.DataFrame({
        "grade": ["Chuyên viên", "Chuyên viên", "Quản lý"],
        "department": ["IT", "IT", "HR"],
        "duration_hours": [8, 4, 6],
        ...
    })
    result = build_grade_level_canvas(inputs)
    # Assert: có section tổng hợp theo ngạch
    # Assert: "Chuyên viên" có tổng giờ = 12
    # Assert: "Quản lý" có tổng giờ = 6

def test_boolean_dtype_regression():
    """~True phải trả False, không phải -2."""
    lookup = pd.DataFrame({"_lookup_found": pd.array([True], dtype="boolean")})
    assert (~lookup["_lookup_found"].fillna(False)).iloc[0] == False
```

---

### Phase 4: Copy dự án sang V2

| # | Công việc |
|---|-----------|
| 4.1 | Copy file Python + `.cmd` + tests sang `D:\Automation\LD REPORT V2` |
| 4.2 | Cập nhật `DEFAULT_WORKBOOK_PATH` cho V2 |
| 4.3 | Copy file Excel workbook (nếu cần) |
| 4.4 | Chạy test đảm bảo V2 hoạt động độc lập |

---

### Phase 5: Cập nhật hướng dẫn sử dụng

| # | File | Nội dung cập nhật |
|---|------|-------------------|
| 5.1 | `docs/HUONG_DAN_SU_DUNG.md` | Hướng dẫn admin chạy + giải thích 9 sheet output |
| 5.2 | `docs/CHANGELOG_V2.md` | Danh sách thay đổi V1 → V2 |
| 5.3 | `README.md` trong LD REPORT V2 | Setup + cách chạy |

**Nội dung hướng dẫn:**

```
1. Chuẩn bị dữ liệu (cột Ngạch + Cấp bậc phải có)
2. Đóng Excel trước khi chạy
3. Bấm đúp Chay_Bao_Cao_Local.cmd
4. Bấm "Kiểm tra dữ liệu nguồn"
5. Bấm "Đồng bộ local + làm mới báo cáo"
6. Mở Excel xem 9 sheet báo cáo (mới: "Phân tích ngạch cấp bậc")
```

---

## 3. KẾ HOẠCH THỰC HIỆN

| Phase | Tên | Thời gian | Phụ thuộc |
|-------|-----|-----------|-----------|
| 1 | Fix & Normalize | 5 phút | — |
| 2 | Sheet mới | 15 phút | Phase 1 |
| 3 | TDD Tests | 10 phút | Phase 1+2 |
| 4 | Copy V2 | 3 phút | Phase 1+2+3 |
| 5 | Hướng dẫn sử dụng | 5 phút | Phase 4 |

**Tổng ước tính: ~40 phút**

---

## 4. VERIFICATION PLAN

### Automated Tests
1. `python -m pytest tests/` → Tất cả test mới PASS
2. `npm test` → 44/44 PASS (không regression)
3. `python tmp_simulate.py` → 577 pass, 2 fail

### Manual Verification
1. Chạy `Chay_Bao_Cao_Local.cmd` trên V2
2. Mở Excel → sheet "Phân tích ngạch cấp bậc" hiện đúng 6 section
3. Kiểm tra dữ liệu Ngạch/Cấp bậc nhất quán

---

## Open Questions

> [!IMPORTANT]
> **Anh xác nhận trước khi em bắt tay:**
> 1. Cấu trúc 6 section (KPI → Ngạch → Cấp bậc → PB×Ngạch → PB×Cấp bậc → Top 50 NV) có đúng ý anh không?
> 2. "Ngạch" và "Cấp bậc" trong dữ liệu của anh hiện có những giá trị gì? (VD: Chuyên viên, Quản lý, L1, L2...) — em cần biết để tạo test data đúng.
> 3. Folder V2 (`D:\Automation\LD REPORT V2`) em copy **toàn bộ dự án** hay chỉ file cần thiết?
