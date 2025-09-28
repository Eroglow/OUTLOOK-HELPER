#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
打包腳本：將應用程式打包為 .exe 檔案
使用 PyInstaller 創建獨立的可執行檔
"""

import os
import sys
import subprocess
from pathlib import Path

def check_pyinstaller():
    """檢查是否已安裝 PyInstaller"""
    try:
        import PyInstaller
        return True
    except ImportError:
        return False

def install_pyinstaller():
    """安裝 PyInstaller"""
    print("正在安裝 PyInstaller...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("PyInstaller 安裝成功!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"PyInstaller 安裝失敗: {e}")
        return False

def create_spec_file():
    """創建 PyInstaller 規格文件"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['outlook_mail_sender.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'win32com.client',
        'win32com.gen_py',
        'pandas',
        'openpyxl',
        'pyperclip',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Outlook群發郵件工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 不顯示控制台視窗
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None,
)
'''
    
    with open('outlook_mail_sender.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print("已創建 PyInstaller 規格文件")

def build_executable():
    """編譯可執行檔"""
    print("開始編譯可執行檔...")
    print("這可能需要幾分鐘時間，請耐心等待...")
    
    try:
        # 使用規格文件編譯
        subprocess.check_call([
            'pyinstaller', 
            '--clean',
            '--noconfirm',
            'outlook_mail_sender.spec'
        ])
        
        print("\n" + "="*50)
        print("編譯完成！")
        print("可執行檔位置: dist/Outlook群發郵件工具.exe")
        print("您可以將此檔案分發給其他使用者")
        print("="*50)
        
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"編譯失敗: {e}")
        return False

def cleanup():
    """清理臨時檔案"""
    import shutil
    
    cleanup_dirs = ['build', '__pycache__']
    cleanup_files = ['outlook_mail_sender.spec']
    
    for dir_name in cleanup_dirs:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"已清理: {dir_name}")
    
    for file_name in cleanup_files:
        if os.path.exists(file_name):
            os.remove(file_name)
            print(f"已清理: {file_name}")

def main():
    """主函數"""
    print("Outlook 群發郵件工具 - 打包程式")
    print("="*40)
    
    # 檢查當前目錄
    if not os.path.exists('outlook_mail_sender.py'):
        print("錯誤：找不到 outlook_mail_sender.py")
        print("請在正確的目錄中執行此腳本")
        return False
    
    # 檢查並安裝 PyInstaller
    if not check_pyinstaller():
        print("PyInstaller 未安裝，正在安裝...")
        if not install_pyinstaller():
            print("無法安裝 PyInstaller，請手動安裝:")
            print("pip install pyinstaller")
            return False
    
    # 創建規格文件
    create_spec_file()
    
    # 編譯可執行檔
    if build_executable():
        print("\n是否清理臨時檔案？(y/n): ", end='')
        choice = input().lower()
        if choice in ['y', 'yes', '是']:
            cleanup()
            print("清理完成")
        
        print("\n編譯成功！可執行檔已準備就緒。")
        return True
    else:
        print("\n編譯失敗，請檢查錯誤訊息。")
        return False

if __name__ == "__main__":
    try:
        success = main()
        input("\n按任意鍵結束...")
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n用戶中斷操作")
        sys.exit(1)
    except Exception as e:
        print(f"\n發生未預期的錯誤: {e}")
        input("按任意鍵結束...")
        sys.exit(1)