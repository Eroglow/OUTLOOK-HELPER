## Excel/剪貼簿群發郵件自動化工具（Windows）

一個以 Python 與 PyQt5 開發的 Windows 桌面應用程式，支援透過 Excel/CSV 或剪貼簿資料批量發送個人化郵件到 Outlook（每位收件人單獨寄送）。

### 功能亮點
- 視覺化 GUI 操作（PyQt5）
- 支援上傳 Excel/CSV 或從剪貼簿貼上資料
- 欄位對應（電子信箱、姓名等）
- 郵件主旨/內文模板，支援變數 `{欄位名}`
- 預覽單封郵件內容
- 以 Outlook COM 自動化 `Display` 預覽或 `Send` 直接寄送
- 進度、成功/失敗數與日誌顯示
- 可選：HTML 格式、附件、匯出日誌

### 系統需求
- Windows 10/11
- 已安裝 Microsoft Outlook（並已設定預設帳戶）
- Python 3.8+

### 安裝與執行
1. 安裝 Python 3.8+（建議於 Windows）
2. 在專案根目錄安裝依賴：
   ```bash
   pip install -r requirements.txt
   ```
3. 執行應用：
   ```bash
   python -m app.main
   ```

> 第一次使用可能遇到 Outlook 安全性提示，請選擇允許。

### 使用說明
1. 點選「選擇檔案」載入 `.xlsx`/`.csv`，或點「從剪貼簿貼上」貼上自 Excel 的表格。
2. 在資料預覽下方設定欄位對應（至少需選擇「電子信箱」欄位）。
3. 編輯主旨與內文（支援 `{欄位名}` 變數）。
4. 選擇是否以 `Display` 預覽，或直接 `Send` 寄送；可勾選 HTML 模式與指定附件。
5. 點「預覽郵件」查看樣本；確認後點「開始發送」。可用「停止」中斷。

### 注意事項
- 僅支援 Windows 系統。
- Outlook 必須已安裝並設定預設帳戶。
- 若 Outlook 未啟動，程式會透過 COM 自動啟動。

### 疑難排解
- 匯入 pywin32 失敗：請在 Windows 上安裝並確保 Python 與位元數一致（x64 ↔ x86）。
- Excel 讀取錯誤：確認安裝 `openpyxl`，檔案未被其他程式鎖定。
- 剪貼簿：部分系統需安裝 `pyperclip` 或允許剪貼簿權限。

### 授權
MIT
# OUTLOOK-HELPER
OUTLOOK HELPER
