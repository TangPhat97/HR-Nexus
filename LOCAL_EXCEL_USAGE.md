# Local Excel Runner

## Muc tieu

Tai lieu nay dung cho admin test Python analytics engine tren file Excel local truoc khi dua len Cloud Run.

## File can mo

- Double-click `Chay_Bao_Cao_Local.cmd`
- Hoac chay tay: `py -3 local_excel_app.py`

App se mo man hinh `HR-NEXUS Local Excel Runner`.

## Admin nhap gi tren workbook

### 1. Sheet `Danh muc nhan vien`

Can cap nhat it nhat:

- `Ma nhan vien`
- `Ho ten`
- `Email`
- `Phong ban`
- `Trang thai dong`

Neu nhan vien da nghi viec nhung van can giu lich su dao tao, van de `Trang thai dong = ACTIVE`.

### 2. Sheet `Lop dao tao`

Can cap nhat it nhat:

- `Ma lop dao tao`
- `Ma lop`
- `Ten khoa hoc`
- `Ngay dao tao`
- `Trang thai dong`

Neu can bao cao chi phi ngoai bo, nen co them:

- `Chi phi dao tao (du kien)`
- `Chi phi binh quan / nguoi`
- `Hoc phi / Chi phi giang vien`

### 3. Sheet `Data raw hoc vien`

Dung khi can doi soat hoc vien theo lop:

- `Ma lop dao tao`
- `Ma nhan vien`
- `Trang thai diem danh`
- `Trang thai dong`

### 4. Sheet `Du lieu dao tao`

Local runner se tu rebuild lai sheet nay tu:

- `Danh muc nhan vien`
- `Danh muc khoa hoc`
- `Lop dao tao`
- `Data raw hoc vien`

Admin khong can chuan bi fact table bang tay nua.

## Bam gi de chay

1. Dong workbook Excel truoc khi chay.
2. Mo `Chay_Bao_Cao_Local.cmd`.
3. Bam `Chon file` neu muon doi workbook.
4. Bam `Kiem tra du lieu nguon`.
5. Neu chi canh bao nhe, bam `Dong bo local + lam moi bao cao`.
6. Mo lai workbook va xem 8 sheet output:

- `Bang dieu khien`
- `Dieu hanh he thong`
- `Phan tich khoa hoc`
- `Phan tich phong ban`
- `Phong ban theo khoa`
- `Phan tich xu huong`
- `Cu di hoc ben ngoai`
- `Doi chieu lop dao tao`

## App lam gi khi chay

- Doc workbook local
- Map header tieng Viet sang schema noi bo cua `transform.py`
- Tu dong dong bo `Data raw hoc vien -> Du lieu dao tao`
- Chay `pandas` engine
- Tao backup timestamp trong thu muc `_local_backups`
- Ghi de 8 sheet bao cao

## Loi thuong gap

### Khong ghi duoc workbook

Nguyen nhan thuong la file Excel dang mo.

Cach xu ly:

- Dong workbook trong Excel
- Chay lai local runner

### Co `Data raw hoc vien` nhung bao cao van trong

Nguyen nhan:

- Raw dang FAIL vi khong match duoc `Danh muc nhan vien` hoac `Lop dao tao`

Cach xu ly:

- Kiem tra log trong app
- Sua lai ma nhan vien, ma lop, ngay dao tao, trang thai diem danh
- Bam chay lai local runner

### Sheet thieu cot

App se hien canh bao ngay trong phan log.

Cach xu ly:

- Doi chieu lai ten cot trong workbook
- Khong doi ten cac cot nguon quan trong neu chua cap nhat mapper trong `local_excel_runner.py`
