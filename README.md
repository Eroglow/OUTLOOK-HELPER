## Outlook 群發郵件工具（Python + tkinter）

**功能**
- 從 Excel（.xlsx/.xls）或剪貼簿解析收件人（姓名/郵件）。
- Outlook COM 發送，每位收件人獨立郵件，支援個人化變數如 `{姓名}`。
- 支援附件、多附件、預覽第一封郵件、進度顯示與中斷、錯誤日誌。

**需求**
- Windows（已安裝 Microsoft Outlook 桌面版）
- Python 3.8+

**安裝**
```bash
pip install -r requirements.txt
```

**執行**
```bash
python src/app_tk.py
```

**打包成 exe（可選）**
```bash
pyinstaller --noconfirm --windowed --onefile --name OutlookBulkMailer src/app_tk.py
```

首次開發時建議以 `mail.Display()` 測試郵件內容後，再改為 `mail.Send()`。

# OUTLOOK-HELPER
OUTLOOK HELPER
