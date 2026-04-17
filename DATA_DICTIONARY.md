# DATA DICTIONARY

## MASTER_EMPLOYEES
- `emp_id`: khoa chinh nhan vien.
- `full_name`, `email`: thong tin nhan dang.
- `employment_status`, `company`, `division`, `department`, `job_title`, `grade`, `level`, `region`: snapshot HR hien hanh.
- `updated_at`, `updated_by`, `row_status`: system-managed metadata.

## MASTER_COURSES
- `course_id`: khoa chinh khoa hoc.
- `course_name`, `course_category`, `platform`: metadata khoa hoc.
- `duration_hours`, `cost_per_pax`: thong tin phan tich chi phi.
- `course_description`, `updated_at`, `updated_by`, `row_status`: metadata.

## TRAINING_RECORDS
- `record_id`: id noi bo duy nhat.
- `source_type`, `batch_id`: dau vet nguon vao.
- `emp_id` ... `level`, `region`: snapshot nhan vien tai thoi diem publish.
- `course_id` ... `cost_per_pax`: snapshot khoa hoc tai thoi diem publish.
- `training_date`, `training_format`, `attendance_status`, `score`: du lieu buoi hoc.
- `satisfaction`, `relevance`, `nps`, `applied_on_job`, `manager_comment`: du lieu danh gia.
- `created_at`, `created_by`, `updated_at`, `updated_by`, `archive_year`, `row_status`: lifecycle metadata.
- `source_row_hash`, `qa_status`, `metadata_json`: internal metadata phuc vu idempotency va audit.

## IMPORT_*_STAGING
- `staging_id`: id dong staging.
- `batch_id`, `source_file_name`, `source_row_number`, `source_row_hash`: metadata import.
- Truong business phu thuoc tung sheet staging.
- `row_status`: `STAGED`, `QA_PASSED`, `QA_WARN`, `QA_FAILED`, `PUBLISHED`.
- `notes`: tom tat QA/publish.

## QA_RESULTS
- 1 dong = 1 issue.
- `severity`: `FAIL` hoac `WARN`.
- `status`: hien tai dung `OPEN`.

## AUDIT_LOGS / ERROR_LOGS
- `AUDIT_LOGS`: ai lam gi, khi nao, thanh cong hay that bai.
- `ERROR_LOGS`: loi ky thuat co stack/context.
