@echo off
setlocal
cd /d "%~dp0"
py -3 local_excel_app.py
if errorlevel 1 (
  echo.
  echo Local runner gap loi. Nhan phim bat ky de dong cua so nay.
  pause >nul
)
