@echo off
title Outlook 群發郵件工具
echo =====================================
echo     Outlook 群發郵件工具
echo =====================================
echo.
echo 正在啟動程式...
echo.

REM 檢查 Python 是否安裝
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [錯誤] 未偵測到 Python！
    echo 請先安裝 Python 3.8 或更高版本
    echo 下載網址：https://www.python.org/downloads/
    pause
    exit /b 1
)

REM 檢查並安裝依賴套件
echo 正在檢查依賴套件...
pip show pandas >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在安裝必要套件，請稍候...
    pip install -r requirements.txt
)

REM 執行主程式
echo.
echo 啟動 Outlook 群發郵件工具...
python outlook_bulk_mailer.py

if %errorlevel% neq 0 (
    echo.
    echo [錯誤] 程式執行失敗！
    echo 請檢查錯誤訊息並確認：
    echo 1. 已安裝所有必要套件
    echo 2. Microsoft Outlook 已安裝
    pause
)