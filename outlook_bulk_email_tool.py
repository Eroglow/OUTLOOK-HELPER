#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook 群發郵件工具
功能：從 Excel 或剪貼簿讀取收件人，透過 Outlook 發送個性化郵件
作者：AI Assistant
版本：1.0
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import win32com.client
import re
import os
import sys
from typing import List, Dict, Tuple, Optional
import threading
import time


class OutlookBulkEmailTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Outlook 群發郵件工具")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        # 數據存儲
        self.recipients = []  # 收件人列表 [(name, email), ...]
        self.attachment_paths = []  # 附件路徑列表
        
        # 創建 GUI
        self.create_widgets()
        
        # 初始化 Outlook
        self.outlook = None
        self.init_outlook()
    
    def create_widgets(self):
        """創建 GUI 組件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置網格權重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 標題
        title_label = ttk.Label(main_frame, text="Outlook 群發郵件工具", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 收件人來源選擇
        source_frame = ttk.LabelFrame(main_frame, text="收件人來源", padding="10")
        source_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        source_frame.columnconfigure(1, weight=1)
        
        ttk.Button(source_frame, text="選擇 Excel 文件", 
                  command=self.load_excel_file).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(source_frame, text="從剪貼簿貼上", 
                  command=self.load_from_clipboard).grid(row=0, column=1, padx=(0, 10))
        ttk.Button(source_frame, text="清空收件人", 
                  command=self.clear_recipients).grid(row=0, column=2)
        
        # 收件人列表顯示
        recipients_frame = ttk.LabelFrame(main_frame, text="收件人列表", padding="10")
        recipients_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        recipients_frame.columnconfigure(0, weight=1)
        recipients_frame.rowconfigure(0, weight=1)
        
        # 收件人列表樹形視圖
        columns = ("姓名", "郵件地址")
        self.recipients_tree = ttk.Treeview(recipients_frame, columns=columns, show="headings", height=8)
        
        for col in columns:
            self.recipients_tree.heading(col, text=col)
            self.recipients_tree.column(col, width=200)
        
        # 滾動條
        recipients_scrollbar = ttk.Scrollbar(recipients_frame, orient=tk.VERTICAL, command=self.recipients_tree.yview)
        self.recipients_tree.configure(yscrollcommand=recipients_scrollbar.set)
        
        self.recipients_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        recipients_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 郵件內容設置
        email_frame = ttk.LabelFrame(main_frame, text="郵件內容", padding="10")
        email_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        email_frame.columnconfigure(1, weight=1)
        
        # 主題
        ttk.Label(email_frame, text="主題:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.subject_var = tk.StringVar()
        self.subject_entry = ttk.Entry(email_frame, textvariable=self.subject_var, width=50)
        self.subject_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # 郵件內容
        ttk.Label(email_frame, text="內容:").grid(row=1, column=0, sticky=(tk.W, tk.N), pady=(0, 5))
        self.content_text = scrolledtext.ScrolledText(email_frame, height=8, width=50)
        self.content_text.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # 預設內容
        default_content = """親愛的 {姓名}，

這是您的專屬通知信。

祝好！
"""
        self.content_text.insert(tk.END, default_content)
        
        # 附件設置
        attachment_frame = ttk.LabelFrame(main_frame, text="附件設置", padding="10")
        attachment_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        attachment_frame.columnconfigure(1, weight=1)
        
        self.include_attachment_var = tk.BooleanVar()
        ttk.Checkbutton(attachment_frame, text="包含附件", 
                       variable=self.include_attachment_var,
                       command=self.toggle_attachment).grid(row=0, column=0, sticky=tk.W)
        
        self.attachment_label = ttk.Label(attachment_frame, text="未選擇附件", foreground="gray")
        self.attachment_label.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0))
        
        ttk.Button(attachment_frame, text="選擇附件", 
                  command=self.select_attachments).grid(row=0, column=2, padx=(10, 0))
        
        # 操作按鈕
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=(10, 0))
        
        ttk.Button(button_frame, text="預覽郵件", 
                  command=self.preview_email).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(button_frame, text="發送郵件", 
                  command=self.send_emails).grid(row=0, column=1, padx=(0, 10))
        ttk.Button(button_frame, text="退出", 
                  command=self.root.quit).grid(row=0, column=2)
        
        # 狀態欄
        self.status_var = tk.StringVar()
        self.status_var.set("就緒")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 進度條
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
    
    def init_outlook(self):
        """初始化 Outlook 連接"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.status_var.set("Outlook 連接成功")
        except Exception as e:
            messagebox.showerror("錯誤", f"無法連接到 Outlook：{str(e)}\n請確保已安裝 Microsoft Outlook")
            self.status_var.set("Outlook 連接失敗")
    
    def load_excel_file(self):
        """從 Excel 文件載入收件人"""
        file_path = filedialog.askopenfilename(
            title="選擇 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # 讀取 Excel 文件
            df = pd.read_excel(file_path)
            self.parse_recipients_from_dataframe(df)
            self.status_var.set(f"已從 Excel 文件載入 {len(self.recipients)} 個收件人")
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取 Excel 文件失敗：{str(e)}")
    
    def load_from_clipboard(self):
        """從剪貼簿載入收件人"""
        try:
            clipboard_text = self.root.clipboard_get()
            self.parse_recipients_from_text(clipboard_text)
            self.status_var.set(f"已從剪貼簿載入 {len(self.recipients)} 個收件人")
        except tk.TclError:
            messagebox.showwarning("警告", "剪貼簿為空或無法讀取")
        except Exception as e:
            messagebox.showerror("錯誤", f"解析剪貼簿內容失敗：{str(e)}")
    
    def parse_recipients_from_dataframe(self, df: pd.DataFrame):
        """從 DataFrame 解析收件人"""
        recipients = []
        
        # 尋找姓名和郵件欄位
        name_col = None
        email_col = None
        
        # 可能的姓名欄位名稱
        name_patterns = ['姓名', 'name', 'Name', '名字', '收件人', 'recipient']
        # 可能的郵件欄位名稱
        email_patterns = ['郵件', 'email', 'Email', '電子郵件', 'e-mail', 'E-Mail', 'mail', 'Mail']
        
        # 尋找姓名欄位
        for col in df.columns:
            if any(pattern in str(col) for pattern in name_patterns):
                name_col = col
                break
        
        # 尋找郵件欄位
        for col in df.columns:
            if any(pattern in str(col) for pattern in email_patterns):
                email_col = col
                break
        
        # 如果沒找到明確的郵件欄位，尋找包含 @ 的欄位
        if email_col is None:
            for col in df.columns:
                if df[col].astype(str).str.contains('@', na=False).any():
                    email_col = col
                    break
        
        if email_col is None:
            raise ValueError("無法找到包含郵件地址的欄位")
        
        # 解析收件人
        for index, row in df.iterrows():
            email = str(row[email_col]).strip()
            if self.is_valid_email(email):
                name = ""
                if name_col and pd.notna(row[name_col]):
                    name = str(row[name_col]).strip()
                elif email_col != name_col:
                    # 如果沒有姓名，嘗試從郵件地址提取
                    name = email.split('@')[0]
                
                recipients.append((name, email))
        
        self.recipients = recipients
        self.update_recipients_display()
    
    def parse_recipients_from_text(self, text: str):
        """從文字解析收件人"""
        recipients = []
        lines = text.strip().split('\n')
        
        # 檢查第一行是否為標題
        header_line = lines[0] if lines else ""
        has_header = any(keyword in header_line.lower() for keyword in ['姓名', 'name', '郵件', 'email'])
        
        start_line = 1 if has_header else 0
        
        for line in lines[start_line:]:
            line = line.strip()
            if not line:
                continue
            
            # 分割行（支援 tab、逗號、空格分隔）
            parts = re.split(r'[\t,;]', line)
            parts = [part.strip() for part in parts if part.strip()]
            
            if len(parts) >= 2:
                # 假設前兩列是姓名和郵件
                name, email = parts[0], parts[1]
                if self.is_valid_email(email):
                    recipients.append((name, email))
            elif len(parts) == 1:
                # 只有一列，檢查是否為郵件
                email = parts[0]
                if self.is_valid_email(email):
                    name = email.split('@')[0]
                    recipients.append((name, email))
        
        self.recipients = recipients
        self.update_recipients_display()
    
    def is_valid_email(self, email: str) -> bool:
        """驗證郵件地址格式"""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    def update_recipients_display(self):
        """更新收件人列表顯示"""
        # 清空現有項目
        for item in self.recipients_tree.get_children():
            self.recipients_tree.delete(item)
        
        # 添加新項目
        for name, email in self.recipients:
            self.recipients_tree.insert("", tk.END, values=(name, email))
    
    def clear_recipients(self):
        """清空收件人列表"""
        self.recipients = []
        self.update_recipients_display()
        self.status_var.set("已清空收件人列表")
    
    def toggle_attachment(self):
        """切換附件選項"""
        if self.include_attachment_var.get():
            self.select_attachments()
        else:
            self.attachment_paths = []
            self.attachment_label.config(text="未選擇附件", foreground="gray")
    
    def select_attachments(self):
        """選擇附件文件"""
        file_paths = filedialog.askopenfilenames(
            title="選擇附件文件",
            filetypes=[("所有文件", "*.*")]
        )
        
        if file_paths:
            self.attachment_paths = list(file_paths)
            if len(file_paths) == 1:
                self.attachment_label.config(text=f"已選擇：{os.path.basename(file_paths[0])}", foreground="black")
            else:
                self.attachment_label.config(text=f"已選擇 {len(file_paths)} 個文件", foreground="black")
        else:
            self.attachment_paths = []
            self.attachment_label.config(text="未選擇附件", foreground="gray")
    
    def preview_email(self):
        """預覽郵件內容"""
        if not self.recipients:
            messagebox.showwarning("警告", "請先載入收件人")
            return
        
        if not self.subject_var.get().strip():
            messagebox.showwarning("警告", "請輸入郵件主題")
            return
        
        # 使用第一個收件人進行預覽
        name, email = self.recipients[0]
        subject = self.subject_var.get().strip()
        content = self.content_text.get("1.0", tk.END).strip()
        
        # 替換變數
        preview_content = content.replace("{姓名}", name if name else "您好")
        
        # 顯示預覽窗口
        preview_window = tk.Toplevel(self.root)
        preview_window.title("郵件預覽")
        preview_window.geometry("600x500")
        
        ttk.Label(preview_window, text="收件人:", font=("Arial", 10, "bold")).pack(anchor=tk.W, padx=10, pady=(10, 5))
        ttk.Label(preview_window, text=f"{name} <{email}>").pack(anchor=tk.W, padx=10, pady=(0, 10))
        
        ttk.Label(preview_window, text="主題:", font=("Arial", 10, "bold")).pack(anchor=tk.W, padx=10, pady=(0, 5))
        ttk.Label(preview_window, text=subject).pack(anchor=tk.W, padx=10, pady=(0, 10))
        
        ttk.Label(preview_window, text="內容:", font=("Arial", 10, "bold")).pack(anchor=tk.W, padx=10, pady=(0, 5))
        
        content_frame = ttk.Frame(preview_window)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        content_text = scrolledtext.ScrolledText(content_frame, wrap=tk.WORD)
        content_text.pack(fill=tk.BOTH, expand=True)
        content_text.insert(tk.END, preview_content)
        content_text.config(state=tk.DISABLED)
        
        ttk.Button(preview_window, text="關閉", command=preview_window.destroy).pack(pady=10)
    
    def send_emails(self):
        """發送郵件"""
        if not self.recipients:
            messagebox.showwarning("警告", "請先載入收件人")
            return
        
        if not self.subject_var.get().strip():
            messagebox.showwarning("警告", "請輸入郵件主題")
            return
        
        if not self.outlook:
            messagebox.showerror("錯誤", "Outlook 未連接")
            return
        
        # 確認發送
        result = messagebox.askyesno("確認發送", 
                                   f"即將發送 {len(self.recipients)} 封郵件，確定嗎？")
        if not result:
            return
        
        # 在新線程中發送郵件
        threading.Thread(target=self._send_emails_thread, daemon=True).start()
    
    def _send_emails_thread(self):
        """在後台線程中發送郵件"""
        try:
            total = len(self.recipients)
            success_count = 0
            failed_recipients = []
            
            for i, (name, email) in enumerate(self.recipients):
                try:
                    # 更新狀態
                    self.root.after(0, lambda: self.status_var.set(f"正在發送第 {i+1}/{total} 封郵件..."))
                    self.root.after(0, lambda: self.progress_var.set((i / total) * 100))
                    
                    # 創建郵件
                    mail = self.outlook.CreateItem(0)  # 0 = olMailItem
                    mail.To = email
                    mail.Subject = self.subject_var.get().strip()
                    
                    # 處理郵件內容
                    content = self.content_text.get("1.0", tk.END).strip()
                    personalized_content = content.replace("{姓名}", name if name else "您好")
                    mail.Body = personalized_content
                    
                    # 添加附件
                    for attachment_path in self.attachment_paths:
                        if os.path.exists(attachment_path):
                            mail.Attachments.Add(attachment_path)
                    
                    # 發送郵件
                    mail.Send()
                    success_count += 1
                    
                    # 短暫延遲，避免發送過快
                    time.sleep(0.5)
                    
                except Exception as e:
                    failed_recipients.append((name, email, str(e)))
                    print(f"發送失敗 {email}: {str(e)}")
            
            # 更新最終狀態
            self.root.after(0, lambda: self.progress_var.set(100))
            
            if failed_recipients:
                error_msg = f"發送完成！\n成功：{success_count} 封\n失敗：{len(failed_recipients)} 封\n\n失敗詳情：\n"
                for name, email, error in failed_recipients:
                    error_msg += f"{name} <{email}>: {error}\n"
                self.root.after(0, lambda: messagebox.showwarning("發送完成", error_msg))
            else:
                self.root.after(0, lambda: messagebox.showinfo("發送完成", f"全部 {success_count} 封郵件已成功發送！"))
            
            self.root.after(0, lambda: self.status_var.set(f"發送完成 - 成功：{success_count}，失敗：{len(failed_recipients)}"))
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("錯誤", f"發送過程中發生錯誤：{str(e)}"))
            self.root.after(0, lambda: self.status_var.set("發送失敗"))


def main():
    """主函數"""
    root = tk.Tk()
    app = OutlookBulkEmailTool(root)
    
    # 設置窗口圖標（如果有的話）
    try:
        root.iconbitmap("icon.ico")
    except:
        pass
    
    # 啟動應用程式
    root.mainloop()


if __name__ == "__main__":
    main()