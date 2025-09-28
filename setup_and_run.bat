@echo off
echo ================================
echo Outlook 群發郵件工具 安裝程式
echo ================================
echo.

echo 檢查 Python 安裝...
python --version >nul 2>&1
if errorlevel 1 (
    echo 錯誤：未找到 Python！
    echo 請先安裝 Python 3.8 或更高版本
    echo 下載地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Python 已安裝
echo.

echo 安裝相依套件...
pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo 錯誤：套件安裝失敗！
    echo 請檢查網路連接或手動安裝套件：
    echo pip install pandas openpyxl pywin32 pyperclip
    pause
    exit /b 1
)

echo.
echo ================================
echo 安裝完成！正在啟動應用程式...
echo ================================
echo.

python outlook_mail_sender.py

if errorlevel 1 (
    echo.
    echo 程式執行時發生錯誤！
    echo 請確認：
    echo 1. Microsoft Outlook 已安裝並設定完成
    echo 2. 所有相依套件已正確安裝
    echo 3. 防毒軟體沒有阻擋程式執行
    pause
)

echo.
echo 程式已結束
pause