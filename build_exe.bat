@echo off
title 建立 EXE 執行檔
echo =====================================
echo   建立 Outlook 群發郵件工具 EXE 檔案
echo =====================================
echo.

REM 檢查 PyInstaller 是否安裝
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在安裝 PyInstaller...
    pip install pyinstaller
)

echo.
echo 正在打包程式為 EXE 檔案...
echo 這可能需要幾分鐘時間，請耐心等待...
echo.

REM 建立 EXE（單一檔案、無控制台視窗）
pyinstaller --onefile ^
            --windowed ^
            --name "OutlookBulkMailer" ^
            --add-data "README.md;." ^
            --hidden-import=win32com.client ^
            --hidden-import=pandas ^
            --hidden-import=openpyxl ^
            --hidden-import=xlrd ^
            outlook_bulk_mailer.py

if %errorlevel% equ 0 (
    echo.
    echo =====================================
    echo    ✓ EXE 檔案建立成功！
    echo =====================================
    echo.
    echo 檔案位置：dist\OutlookBulkMailer.exe
    echo.
    echo 您可以將此檔案複製到任何 Windows 電腦上執行
    echo （無需安裝 Python）
    echo.
    explorer dist
) else (
    echo.
    echo [錯誤] EXE 檔案建立失敗！
    echo 請檢查錯誤訊息
)

pause