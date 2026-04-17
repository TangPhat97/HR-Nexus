# SAVEBRAIN

## 1. Muc dich file
File nay la bo nho nhanh de dev/AI quay lai project co the bat mach ngay:
- he thong dang o trang thai nao
- da deploy gi len live
- luong van hanh dung la gi
- diem nao de vo
- ngay mai can lam tiep gi

## 2. Trang thai hien tai
- App version: `v3.4.2`
- Nen tang: `Google Sheets + Google Apps Script`
- Apps Script project id: `[YOUR_GAS_SCRIPT_ID]`
- Live deployment id: `AKfycbxyQ5qF3uOk-4I3SEeThFqyLAyAzvPOWNf-TOW22tpfSgu5uFcZ4C5DNoDal6-2CnXO`
- Live deployment version hien tai: `@13`
- Deployment note: `v3.4.2 - schema compatibility hotfix for setup and sync`
- Spreadsheet live:
  - [L&D - 2026 MASTER DATA - TEST](https://docs.google.com/spreadsheets/d/[YOUR_TEST_SPREADSHEET_ID]/edit?gid=1318223922#gid=1318223922)

## 3. He thong nay dang lam gi
Day la workbook van hanh du lieu L&D theo flow:
1. `Danh muc nhan vien`
2. `Danh muc khoa hoc`
3. `Lop dao tao`
4. `Data raw hoc vien`
5. `Du lieu dao tao`
6. `Bang dieu khien + analytics`

Y nghia nghiep vu:
- `Danh muc nhan vien` la employee master de map phong ban, khoi, cap bac, trang thai nhan su
- `Danh muc khoa hoc` la course master
- `Lop dao tao` la tang session/class
- `Data raw hoc vien` la input bulk quan trong nhat
- `Du lieu dao tao` la fact table chuan
- `Phan tich phong ban`, `Phong ban theo khoa`, `Phan tich khoa hoc`, `Phan tich xu huong` la output-only

## 4. Public actions quan trong
- `Lam moi data raw`
- `Dong bo du lieu dao tao`
- `Dong bo lai toan bo du lieu raw`
- `Lam moi Phan tich xu huong`  <-- moi them, chi rebuild trend sheet
- `Lam moi bao cao cu di hoc ben ngoai`  <-- moi them, chi rebuild report cu di hoc ben ngoai
- `Kiem tra auto sync`
- `Khoi tao he thong`
- `Sao luu`
- `Khoi phuc`

## 5. Cac thay doi da chot gan day

### 5.0 Performance + telemetry live
Da co 3 moc moi:
- `v3.4.0 @10`: toi uu hieu suat truoc live cho `Dong bo du lieu dao tao`
- `v3.4.1 @12`: them telemetry nhe vao `Nhat ky tac dong`
- `v3.4.2 @13`: hotfix schema compatibility cho `Khoi tao he thong` va `SYNC_TRAINING_DATA`

Toi uu hieu suat da lam:
- chia se data dung chung ngay trong `refreshAnalyticsCore_`, khong de moi report tu doc lai `training sessions`, `raw participants`, `training records`, `config`
- toi uu pipeline `Doi chieu lop dao tao` theo map / summary mot vong, tranh resolve session lap lai
- doi writer custom report sang reset theo vung da dung, khong dap ca sheet
- giam cac cho doc raw chi de lay `.length`

Telemetry live da lam:
- action chay qua `runAction_` se ghi them:
  - `duration_ms`
  - `started_at`
  - `finished_at`
- khong doi schema sheet `Nhat ky tac dong`
- thong tin thoi gian nam trong:
  - cot `message`
  - `details_json`

Baseline live dau tien da do duoc:
- `refreshTrendAnalyticsOnly`: `11607 ms`
- `refreshExternalAssignmentReport`: `33838 ms`
- `refreshSessionReconciliationReport`: `39581 ms`
- `processQueue`: `160538 ms`, `199891 ms`, `273137 ms`

Ket luan hien tai:
- `Phan tich xu huong` chay kha on
- `Cu di hoc ben ngoai` va `Doi chieu lop dao tao` van la nhom nang
- diem nghen lon nhat hien tai la `processQueue`

### 5.0.1 Hotfix schema compatibility
Da fix 2 tinh huong gay fail sai:
- sheet `Bat dau` da xoa noi dung nhung con format/merge nen van bi coi la "co du lieu"
- sheet `Danh muc khoa hoc` schema cu kieu append-only bi chan du "chi thieu cot moi o cuoi"

Root cause thuc te anh gap tren live:
- `Danh muc khoa hoc` bi mat header `A1 = Ma khoa hoc`
- du lieu cot A van con, nhung header bi trong nen `ensureSheet_()` xem la lech schema

Huong xu ly thuc te:
- dien lai `A1 = Ma khoa hoc` la chay duoc ngay

Fix code da chot:
- phan biet `bodyHasData` theo noi dung thuc, khong theo `lastRow > 1`
- cho phep migrate an toan voi schema cu append-only cho cac sheet input nhu `Danh muc khoa hoc`
- vi vay queue `SYNC_TRAINING_DATA` se khong bi fail day chuyen boi cung mot loi schema nua

### 5.1 Auto sync
Da them auto sync cau hinh qua sheet `Cau hinh he thong`:
- `AUTO_SYNC_ENABLED`
- `AUTO_SYNC_WEEKDAYS`
- `AUTO_SYNC_HOUR`
- `AUTO_SYNC_MINUTE`
- `AUTO_SYNC_TIMEZONE`
- `AUTO_SYNC_MODE`
- `AUTO_SYNC_NOTIFY_EMAIL`
- `AUTO_SYNC_LAST_RUN_AT`
- `AUTO_SYNC_LAST_STATUS`
- `AUTO_SYNC_LAST_MESSAGE`
- `AUTO_SYNC_LAST_SLOT_KEY`
- `AUTO_SYNC_SCHEDULER_ENABLED`
- `AUTO_SYNC_SCHEDULER_INTERVAL_MINUTES`

Da co:
- scheduler nen
- menu `Kiem tra auto sync`
- status lines cho admin

Luu y:
- mac dinh an toan la `AUTO_SYNC_ENABLED = false`
- scheduler nen co the dang bat nhung auto sync van chua chay neu config dang tat

### 5.2 Trend month display
`Phan tich xu huong` la bao cao theo thang, khong phai theo ngay.
Da sua logic de `Ky thang` hien thi dang `YYYY-MM` neu du lieu thang ton tai.
Vi du:
- `2026-01`
- `2026-03`

Neu thang nao khong co du lieu thi khong tu sinh dong trong trend.

### 5.3 Hotfix raw full rebuild
Da sua bug:
- `Loi dong bo lai toan bo du lieu raw`
- `finalRecords is not defined`

Nguyen nhan goc:
- bien `finalRecords` nam trong block scope nhung bi dung lai o ngoai block trong `RawSyncService.gs`

Ket qua sau fix:
- full rebuild khong con vang `ReferenceError`
- `finalRecords` tra ve dung output cuoi sau khi remove record raw cu

### 5.4 Trend-only refresh
Da them action moi:
- `Lam moi Phan tich xu huong`

Action nay:
- doc du lieu hien co trong `TRAINING_RECORDS`
- chi rewrite `ANALYTICS_TREND`
- khong dung vao:
  - `Du lieu dao tao`
  - `Phan tich phong ban`
  - `Phong ban theo khoa`

### 5.5 Bao cao Cu di hoc ben ngoai
Da them report output moi:
- sheet: `Cu di hoc ben ngoai`
- nguon du lieu: `TRAINING_RECORDS`
- dieu kien loc co dinh trong code:
  - `delivery_type = Dao tao ben ngoai`
  - `training_format = Cu nhan su di hoc`
  - mac dinh nam hien tai theo `CURRENT_FISCAL_YEAR`

Sheet nay da co:
- block KPI
- tong hop theo lop
- tong hop theo phong ban
- chi tiet tung hoc vien
- action rieng: `Lam moi bao cao cu di hoc ben ngoai`

Da fix trong cac ban `v3.3.2 -> v3.3.4`:
- them sheet report moi va manual action
- fallback chi phi tu `Lop dao tao` neu `TRAINING_RECORDS` chua co du `estimated_cost`
- ngay dao tao hien thi `dd/MM/yyyy`, khong con dang `Sat Mar 21 2026 ...`
- refresh lai report khong con loi merge phai xoa sheet
- UI duoc polish:
  - KPI dang card ngang
  - mau block ro rang
  - filter san cho block chi tiet hoc vien

Luu y quan trong:
- report nay la `custom multi-block managed output`
- KHONG dua vao `RESTORABLE_SHEETS`
- neu can dung lai sau backup/restore thi rebuild bang:
  - `Dong bo du lieu dao tao`
  - hoac `Lam moi bao cao cu di hoc ben ngoai`

## 6. File vua sua o nhip gan nhat
- [Repository.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/Repository.gs)
- [ErrorHandler.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/ErrorHandler.gs)
- [AnalyticsPresentationService.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/AnalyticsPresentationService.gs)
- [AnalyticsCustomReportService.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/AnalyticsCustomReportService.gs)
- [RawSyncService.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/RawSyncService.gs)
- [AnalyticsService.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/AnalyticsService.gs)
- [AppController.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/AppController.gs)
- [MenuService.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/MenuService.gs)
- [TrainingSessionService.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/TrainingSessionService.gs)
- [action-telemetry-tdd.test.cjs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/action-telemetry-tdd.test.cjs)
- [analytics-custom-report-tdd.test.cjs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/analytics-custom-report-tdd.test.cjs)
- [trend-only-refresh-tdd.test.cjs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/trend-only-refresh-tdd.test.cjs)
- [ExternalAssignmentReportService.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/ExternalAssignmentReportService.gs)
- [external-assignment-report-tdd.test.cjs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/external-assignment-report-tdd.test.cjs)

## 7. Test da verify
Da chay lai truoc khi deploy live:
- `node action-telemetry-tdd.test.cjs`
- `node analytics-custom-report-tdd.test.cjs`
- `node external-assignment-report-tdd.test.cjs`
- `node trend-only-refresh-tdd.test.cjs`
- `node trend-month-display-tdd.test.cjs`
- `node auto-sync-tdd.test.cjs`
- `node tests/raw-sync-regression.test.cjs`
- `node encoding-regression.test.cjs`

Test regression moi da them trong `raw-sync-regression`:
- `START_HERE` blank nhung con format/merge khong duoc fail schema
- `MASTER_COURSES` schema cu append-only phai duoc chap nhan migrate

## 8. Cac diem de vo nhat
- Nhap tay vao cac sheet he thong nhu `Bat dau`, `Huong dan nhap lieu`
- Xoa / mat header dong 1 o sheet input, dac biet `A1 = Ma khoa hoc` trong `Danh muc khoa hoc`
- Dung `Dong bo lai toan bo du lieu raw` khi chi muon refresh trend
- Sua header/cot ma khong update `Constants.gs`
- De `Danh muc nhan vien` thieu nhan su nghi viec can giu lich su dao tao
- Dung nham `Ma khoa hoc` va `Ma lop`
- Doi schema ma khong kiem tra lai `GuideService`, menu va test regression

## 9. Quy uoc du lieu quan trong
- `Ma khoa hoc` = ma cua noi dung dao tao goc
- `Ma lop` = ma cua tung lan to chuc thuc te
- `Ma lop dao tao` = id ky thuat he thong tu sinh

Khuyen nghi:
- nhan vien nghi viec nhung can giu lich su dao tao thi van phai co trong `Danh muc nhan vien`
- `employment_status = Nghi viec`
- `row_status = ACTIVE`

## 10. Neu quay lai project sau 1 thoi gian
Doc theo thu tu nay:
1. [README.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/README.md)
2. [CHANGELOG.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/CHANGELOG.md)
3. [SAVEBRAIN.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/SAVEBRAIN.md)
4. [docs/HANDOVER_2026-03-31.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/docs/HANDOVER_2026-03-31.md)
5. [Constants.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/Constants.gs)
6. [Repository.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/Repository.gs)
7. [RuntimeConfigService.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/RuntimeConfigService.gs)
8. [RawSyncService.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/RawSyncService.gs)
9. [AnalyticsService.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/AnalyticsService.gs)
10. [AppController.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/AppController.gs)

## 11. Viec con no / goi y ngay mai
- Doc tiep telemetry live sau khi admin chay them:
  - `Dong bo du lieu dao tao`
  - `Lam moi bao cao cu di hoc ben ngoai`
  - `Lam moi doi chieu lop dao tao`
- Tach telemetry chi tiet hon cho `processQueue` neu can biet buoc nao dang ton thoi gian nhat
- Toi uu tiep 3 diem nang nhat tren live:
  - `processQueue`
  - `refreshSessionReconciliationReport`
  - `refreshExternalAssignmentReport`
- Lam tiep 1 vong polish UI / phan tich sau cho:
  - `Phan tich phong ban`
  - `Phong ban theo khoa`
  - `Phan tich xu huong`
  - `Doi chieu lop dao tao`
- Nghien cuu tiep sheet `Doi chieu lop dao tao` de support doi soat nguoc giua:
  - so lieu nhap tay o `Lop dao tao`
  - `Data raw hoc vien`
  - `Du lieu dao tao`

## 12. Neu chi nho 5 dieu
- `Du lieu dao tao` la fact table chuan
- `Data raw hoc vien` la input bulk quan trong nhat
- `Phan tich xu huong` la bao cao theo thang
- `Lam moi Phan tich xu huong` chi rebuild trend, khong dung cac sheet analytics khac
- `Nhat ky tac dong` tu `v3.4.1` da co telemetry thoi gian chay action
- `Danh muc khoa hoc` mat header `A1 = Ma khoa hoc` se bi bao loi schema ngay
- `Cu di hoc ben ngoai` la report custom doc tu `TRAINING_RECORDS`, live dang chay `v3.4.2 @13`
