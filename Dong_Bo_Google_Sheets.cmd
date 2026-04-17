@echo off
setlocal
cd /d "%~dp0"
echo ============================================
echo   HR-NEXUS Google Sheets Sync
echo   Dang dong bo du lieu...
echo ============================================
echo.
py -3 gsheet_sync.py sync
if errorlevel 1 (
  echo.
  echo Dong bo gap loi. Nhan phim bat ky de dong cua so nay.
  pause >nul
) else (
  echo.
  echo Hoan tat! Nhan phim bat ky de dong cua so nay.
  pause >nul
)
