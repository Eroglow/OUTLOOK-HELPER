#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook 群發郵件工具 - 打包為可執行文件
使用 PyInstaller 將程式打包為 .exe 文件
"""

import os
import sys
import subprocess

def install_pyinstaller():
    """安裝 PyInstaller"""
    try:
        import PyInstaller
        print("PyInstaller 已安裝")
    except ImportError:
        print("正在安裝 PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

def build_exe():
    """打包為可執行文件"""
    print("開始打包程式...")
    
    # PyInstaller 命令參數
    cmd = [
        "pyinstaller",
        "--onefile",  # 打包為單一文件
        "--windowed",  # 不顯示控制台窗口
        "--name=Outlook群發郵件工具",  # 可執行文件名稱
        "--icon=icon.ico",  # 圖標文件（如果存在）
        "--add-data=README.md;.",  # 包含 README 文件
        "outlook_bulk_email_tool.py"
    ]
    
    # 如果沒有圖標文件，移除圖標參數
    if not os.path.exists("icon.ico"):
        cmd = [arg for arg in cmd if not arg.startswith("--icon")]
    
    try:
        subprocess.check_call(cmd)
        print("打包完成！可執行文件位於 dist/ 目錄中")
    except subprocess.CalledProcessError as e:
        print(f"打包失敗：{e}")
        return False
    
    return True

def main():
    """主函數"""
    print("Outlook 群發郵件工具 - 打包腳本")
    print("=" * 50)
    
    # 檢查必要文件
    if not os.path.exists("outlook_bulk_email_tool.py"):
        print("錯誤：找不到主程式文件 outlook_bulk_email_tool.py")
        return
    
    # 安裝 PyInstaller
    install_pyinstaller()
    
    # 打包程式
    if build_exe():
        print("\n打包成功！")
        print("可執行文件位置：dist/Outlook群發郵件工具.exe")
        print("\n注意事項：")
        print("1. 確保目標電腦已安裝 Microsoft Outlook")
        print("2. 首次運行時可能需要管理員權限")
        print("3. 防毒軟體可能會誤報，請添加信任")
    else:
        print("\n打包失敗，請檢查錯誤訊息")

if __name__ == "__main__":
    main()