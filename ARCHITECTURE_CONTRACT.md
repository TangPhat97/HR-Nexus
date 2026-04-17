# ARCHITECTURE CONTRACT

## Layering
- `Constants.gs`, `Utils.gs`: config, schema, constants, helper.
- `Repository.gs`: bootstrap workbook, CRUD, queue, config persistence.
- `SecurityService.gs`: RBAC va double-action prevention.
- `EmployeeService.gs`, `CourseCatalogService.gs`, `TrainingRecordService.gs`, `ImportService.gs`, `ValidationService.gs`, `AnalyticsService.gs`, `BackupService.gs`, `ArchiveService.gs`: business logic.
- `AuditLogService.gs`, `ErrorHandler.gs`: audit/error boundaries.
- `MenuService.gs`, `AppController.gs`: public entrypoints.
- `StartHere.html`, `ImportCenter.html`, `TrainingEntry.html`: sidebar UI.

## Public Entrypoints
- `setupSystem()`
- `openStartHere()`
- `openImportCenter()`
- `openTrainingEntry()`
- `runQaChecks(batchId?)`
- `publishStaging(batchId?)`
- `refreshAnalytics()`
- `backupSystem()`
- `archiveYear()`
- `restoreBackup()`
- `showSystemInfo()`
- `processQueue()`

## Workbook Topology
- Visible business sheets:
  - `START_HERE`
  - `DASHBOARD_EXEC`
  - `DASHBOARD_OPERATIONS`
  - `MASTER_EMPLOYEES`
  - `MASTER_COURSES`
  - `TRAINING_RECORDS`
  - `ANALYTICS_COURSE`
  - `ANALYTICS_DEPARTMENT`
  - `ANALYTICS_TREND`
  - `IMPORT_HR_STAGING`
  - `IMPORT_COURSES_STAGING`
  - `IMPORT_TRAINING_STAGING`
- Hidden system sheets:
  - `CONFIG_SYSTEM`
  - `CONFIG_USERS`
  - `QA_RESULTS`
  - `QUEUE_JOBS`
  - `AUDIT_LOGS`
  - `ERROR_LOGS`
  - `SNAPSHOT_REPORTS`

## Non-Negotiable Rules
- Tat ca write/read di qua `Repository.gs`.
- Master data duoc resolve thanh snapshot khi publish.
- `TRAINING_RECORDS` la source of truth cho analytics.
- Batch co `FAIL` khong duoc publish.
- Analytics chi duoc refresh bang service, khong sua tay.
- Backup va archive phai ghi audit log.
