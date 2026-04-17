# HANDOVER GUIDE

## 1. Muc tieu ban giao
Tai lieu nay la ban tong hop chuan van hanh de ban giao cho:
- Admin nghiep vu
- Operator
- Dev/AI bao tri Apps Script

Muc tieu:
- hieu he thong dang lam gi
- biet luong van hanh dung
- biet sheet nao duoc nhap tay, sheet nao khong duoc sua
- biet cach xu ly cac loi thuong gap
- biet doc tai lieu nao truoc khi sua he thong

## 2. Trang thai he thong hien tai
- Version live: `v3.4.2`
- Apps Script project id: `[YOUR_GAS_SCRIPT_ID]`
- Live deployment:
  - `AKfycbxyQ5qF3uOk-4I3SEeThFqyLAyAzvPOWNf-TOW22tpfSgu5uFcZ4C5DNoDal6-2CnXO @13`
- Workbook live:
  - [L&D - 2026 MASTER DATA - TEST](https://docs.google.com/spreadsheets/d/[YOUR_TEST_SPREADSHEET_ID]/edit?gid=1318223922#gid=1318223922)

## 3. He thong nay dang lam gi
He thong nay van hanh du lieu L&D theo flow:
1. `Danh muc nhan vien`
2. `Danh muc khoa hoc`
3. `Lop dao tao`
4. `Data raw hoc vien`
5. `Du lieu dao tao`
6. `Bang dieu khien + cac sheet analytics`

Y nghia nghiep vu:
- `Danh muc nhan vien`: employee master
- `Danh muc khoa hoc`: course master
- `Lop dao tao`: session / class
- `Data raw hoc vien`: du lieu hoc vien paste bulk
- `Du lieu dao tao`: fact table chuan
- `Phan tich phong ban`, `Phong ban theo khoa`, `Phan tich khoa hoc`, `Phan tich xu huong`: output-only

## 4. Sheet nao duoc nhap tay
### Duoc nhap tay
- `Danh muc nhan vien`
- `Danh muc khoa hoc`
- `Lop dao tao`
- `Data raw hoc vien`
- `Cau hinh he thong`:
  - chi sua `config_value`
  - khong tu y them/sua `config_key` neu khong co huong dan ky thuat

### Khong nen nhap tay
- `Du lieu dao tao`
- `Bang dieu khien`
- `Dieu hanh he thong`
- `Phan tich phong ban`
- `Phong ban theo khoa`
- `Phan tich khoa hoc`
- `Phan tich xu huong`
- `Doi chieu lop dao tao`
- `Hang doi xu ly`
- `Nhat ky tac dong`
- `Nhat ky loi`
- `Bat dau`
- `Huong dan nhap lieu`

Luu y:
- `Bat dau` va `Huong dan nhap lieu` la sheet he thong
- neu nhap du lieu tay vao day, `Khoi tao he thong` co the bao loi schema
- `Danh muc khoa hoc` bat buoc giu dung header dong 1, dac biet `A1 = Ma khoa hoc`

## 5. Cac menu van hanh quan trong
- `Khoi tao he thong`
- `Kiem tra du lieu`
- `Publish du lieu`
- `Lam moi data raw`
- `Dong bo du lieu dao tao`
- `Dong bo lai toan bo du lieu raw`
- `Lam moi Phan tich xu huong`
- `Lam moi bao cao cu di hoc ben ngoai`
- `Lam moi doi chieu lop dao tao`
- `Kiem tra auto sync`
- `Sao luu`
- `Khoi phuc`

## 6. Y nghia tung action chinh
### Lam moi data raw
- xu ly sheet `Data raw hoc vien`
- cap nhat ma raw, hash dong, ghi chu, trang thai QA
- day la buoc lam sach raw truoc khi sync

### Dong bo du lieu dao tao
- dong bo tu raw da hop le sang `Du lieu dao tao`
- cap nhat analytics lien quan
- day la nut van hanh hang ngay
- neu loi schema o sheet input, queue co the fail day chuyen tai day

### Dong bo lai toan bo du lieu raw
- rebuild manh hon
- dung khi can doi soat toan bo hoac sau khi sua ma / mapping lon
- da fix loi `finalRecords is not defined` o live `@6`

### Lam moi Phan tich xu huong
- action moi
- chi doc du lieu hien co trong `TRAINING_RECORDS`
- chi rewrite sheet `Phan tich xu huong`
- khong dung vao:
  - `Du lieu dao tao`
  - `Phan tich phong ban`
  - `Phong ban theo khoa`

### Lam moi bao cao cu di hoc ben ngoai
- doc du lieu hien co trong `TRAINING_RECORDS`
- chi rebuild sheet `Cu di hoc ben ngoai`
- khong dung vao:
  - `Du lieu dao tao`
  - `Phan tich phong ban`
  - `Phong ban theo khoa`
  - `Phan tich xu huong`
- report duoc loc co dinh theo:
  - `delivery_type = Dao tao ben ngoai`
  - `training_format = Cu nhan su di hoc`
- nam hien tai

### Lam moi doi chieu lop dao tao
- doc du lieu hien co trong:
  - `Lop dao tao`
  - `Data raw hoc vien`
  - `Du lieu dao tao`
- chi rebuild sheet `Doi chieu lop dao tao`
- dung de doi soat so lieu nhap tay va so lieu da sync

### Kiem tra auto sync
- hien trang thai cau hinh auto sync
- cho biet auto sync dang bat/tat, scheduler nen, ngay chay, gio chay, mode, backlog, lan chay gan nhat

## 7. Luong van hanh chuan
### Luong hang ngay
1. cap nhat `Danh muc nhan vien` neu can
2. cap nhat `Danh muc khoa hoc` neu can
3. tao / cap nhat `Lop dao tao`
4. paste `Data raw hoc vien`
5. bam `Lam moi data raw`
6. bam `Dong bo du lieu dao tao`
7. kiem tra:
   - `Bang dieu khien`
   - `Phan tich phong ban`
   - `Phong ban theo khoa`
   - `Phan tich xu huong`
   - `Cu di hoc ben ngoai`
   - `Doi chieu lop dao tao`

### Luong chi can lam moi trend
1. giu nguyen `Du lieu dao tao`
2. bam `Lam moi Phan tich xu huong`
3. kiem tra sheet `Phan tich xu huong`

### Luong chi can lam moi bao cao cu di hoc ben ngoai
1. giu nguyen `Du lieu dao tao`
2. bam `Lam moi bao cao cu di hoc ben ngoai`
3. kiem tra sheet `Cu di hoc ben ngoai`

### Luong chi can lam moi doi chieu lop dao tao
1. giu nguyen `Du lieu dao tao`
2. bam `Lam moi doi chieu lop dao tao`
3. kiem tra sheet `Doi chieu lop dao tao`

### Luong doi soat manh
1. backup truoc
2. bam `Dong bo lai toan bo du lieu raw`
3. doi queue xu ly xong neu co backlog
4. kiem tra lai dashboard va analytics

## 8. Quy uoc du lieu quan trong
### Ma khoa hoc / Ma lop / Ma lop dao tao
- `Ma khoa hoc` = ma cua noi dung dao tao goc
- `Ma lop` = ma cua tung lan to chuc thuc te
- `Ma lop dao tao` = id ky thuat do he thong tu sinh

Khuyen nghi:
- nguoi dung nghiep vu chi dung `Ma lop`
- khong sua tay `Ma lop dao tao`
- khong xoa header `Ma khoa hoc` o `Danh muc khoa hoc`

### Loi schema thuong gap
- `Sheet Danh muc khoa hoc khong khop schema v3`:
  - thuong do mat / sua header dong 1
  - truong hop live da gap: `A1` bi trong trong khi du lieu ma khoa hoc van nam o cot A
  - cach xu ly nhanh: dien lai `A1 = Ma khoa hoc`

### Nhan vien nghi viec
Neu van can giu lich su dao tao:
- van phai co trong `Danh muc nhan vien`
- `employment_status = Nghi viec`
- `row_status = ACTIVE`

Neu de `INACTIVE`, he thong co the khong map duoc raw sang du lieu dao tao.

## 9. Phan tich xu huong
- `Phan tich xu huong` la bao cao theo thang
- cot `Ky thang` hien thi dang `YYYY-MM`
- vi du:
  - `2026-01`
  - `2026-03`
- neu thang nao khong co du lieu thi he thong khong tu sinh dong thang do

## 10. Auto sync
He thong da co co che auto sync cau hinh qua `Cau hinh he thong`.

Các key quan trong:
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

Nguyen tac:
- scheduler nen co the dang bat
- nhung neu `AUTO_SYNC_ENABLED = false` thi he thong van chua tu chay
- admin chi sua `config_value`

## 10.1 Bao cao Cu di hoc ben ngoai
Report nay dung de theo doi:
- ai duoc cu di hoc ben ngoai
- thuoc phong ban nao
- tham gia lop nao
- tong chi phi lop va chi phi phan bo tren tung hoc vien

Nguon du lieu:
- `TRAINING_RECORDS`

Dieu kien loc mac dinh trong code:
- `delivery_type = Dao tao ben ngoai`
- `training_format = Cu nhan su di hoc`
- nam hien tai theo `CURRENT_FISCAL_YEAR`

Luu y chi phi:
- uu tien `estimated_cost`
- neu thieu thi fallback theo session hien co trong `Lop dao tao`
- sau do moi fallback `avg_cost_per_pax`
- cuoi cung moi fallback `cost_per_pax`

Luu y ky thuat:
- day la `custom multi-block managed output`
- khong restore theo kieu row-table
- neu can dung lai sau backup/restore thi rebuild bang refresh report hoac sync analytics

## 10.2 Telemetry trong Nhat ky tac dong
Tu `v3.4.1`, cac action chay qua wrapper he thong se ghi them thoi gian chay vao `Nhat ky tac dong`.

Cach doc:
- cot `message`: nhin nhanh, vi du `Da lam moi Phan tich xu huong (11607 ms)`
- `details_json.duration_ms`: dung de loc / pivot / tinh trung binh
- `details_json.started_at`
- `details_json.finished_at`

Y nghia:
- day la nguon do live dung nhat de xem action nao chay cham
- log cu truoc `v3.4.1` se khong co `duration_ms`

## 11. Ban giao cho Admin
Admin can biet:
- cach mo `Thong tin ban giao`
- cach mo `Huong dan nhap lieu`
- cach doc `Cau hinh he thong`
- cach xem `Kiem tra auto sync`
- cach backup truoc khi doi lon
- cach xem `Nhat ky loi` va `Hang doi xu ly`

Checklist:
- [ ] biet 4 sheet nhap lieu chinh
- [ ] biet khong nhap tay vao analytics
- [ ] biet `Lam moi data raw` khac `Dong bo du lieu dao tao`
- [ ] biet khi nao chi dung `Lam moi Phan tich xu huong`
- [ ] biet backup truoc khi doi soat manh

## 12. Ban giao cho Operator
Operator chi can nam luong:
1. cap nhat master neu can
2. tao `Lop dao tao`
3. paste `Data raw hoc vien`
4. chay `Lam moi data raw`
5. chay `Dong bo du lieu dao tao`
6. kiem tra cac sheet bao cao

Operator khong nen:
- sua `Du lieu dao tao`
- sua dashboard / analytics
- sua `Cau hinh he thong` neu khong co huong dan
- tu doi cot / header cac sheet chuan

## 13. Ban giao cho Dev/AI
Dev/AI nen doc theo thu tu:
1. [README.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/README.md)
2. [SYSTEM_SUMMARY.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/SYSTEM_SUMMARY.md)
3. [CHANGELOG.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/CHANGELOG.md)
4. [SAVEBRAIN.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/SAVEBRAIN.md)
5. [HANDOVER_GUIDE.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/HANDOVER_GUIDE.md)
6. [docs/HANDOVER_2026-03-30.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/docs/HANDOVER_2026-03-30.md)
7. [docs/HANDOVER_2026-03-31.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/docs/HANDOVER_2026-03-31.md)
8. [Constants.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/Constants.gs)
9. [Repository.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/Repository.gs)
10. [RuntimeConfigService.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/RuntimeConfigService.gs)
11. [RawSyncService.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/RawSyncService.gs)
12. [AnalyticsService.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/AnalyticsService.gs)
13. [AppController.gs](D:/clean-architecture-canonical/file%20excel%20tu%20dong/gas/AppController.gs)

## 14. Xu ly loi thuong gap
### Khong thay so lieu len report
1. kiem tra `Ma lop`, `Ma khoa hoc`, `Ma nhan vien`
2. chay `Lam moi data raw`
3. chay `Dong bo du lieu dao tao`
4. kiem tra `Nhat ky loi`

### Bao loi schema o sheet he thong
Nguyen nhan thuong gap:
- da nhap du lieu tay vao `Bat dau` hoac `Huong dan nhap lieu`

Cach xu ly:
1. chuyen du lieu do sang sheet khac
2. de lai sheet he thong dung vai tro huong dan
3. chay lai `Khoi tao he thong`

### QA fail: khong tim thay nhan vien
Thuong la:
- nhan vien chua co trong `Danh muc nhan vien`
- hoac da de `row_status = INACTIVE`

Neu la nhan vien nghi viec nhung can giu lich su:
- them/giu ban ghi trong master
- `employment_status = Nghi viec`
- `row_status = ACTIVE`

### Chi can cap nhat bao cao xu huong
- dung `Lam moi Phan tich xu huong`
- khong can chay lai toan bo raw sync

## 15. Backlog giao dien ngay mai
- Tiep tuc polish UI cho:
  - `Phan tich phong ban`
  - `Phong ban theo khoa`
  - `Phan tich xu huong`
- Huong thuc hien:
  - co filter view san o bang chinh
  - block title / tong quan / header dep hon
  - mau sac / border / number format / column width nhat quan hon
  - giu logic du lieu cu, uu tien nang presentation layer

## 15. Tai lieu lich su handover
Day la log chi tiet theo ngay, khong thay the cho guide tong hop:
- [HANDOVER_2026-03-30.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/docs/HANDOVER_2026-03-30.md)
- [HANDOVER_2026-03-31.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/docs/HANDOVER_2026-03-31.md)

## 16. Neu can ban giao gap
Toi thieu phai dua 5 tai lieu nay:
- [README.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/README.md)
- [SYSTEM_SUMMARY.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/SYSTEM_SUMMARY.md)
- [HANDOVER_GUIDE.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/HANDOVER_GUIDE.md)
- [SAVEBRAIN.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/SAVEBRAIN.md)
- [CHANGELOG.md](D:/clean-architecture-canonical/file%20excel%20tu%20dong/CHANGELOG.md)

## 17. Ket luan ngan
Neu chi nho 6 dieu:
- `Du lieu dao tao` la fact table chuan
- `Data raw hoc vien` la input bulk quan trong nhat
- `Lam moi data raw` va `Dong bo du lieu dao tao` la 2 buoc khac nhau
- `Phan tich xu huong` la bao cao theo thang
- `Lam moi Phan tich xu huong` chi rebuild trend sheet
- live hien dang chay `v3.3.1 @6`
