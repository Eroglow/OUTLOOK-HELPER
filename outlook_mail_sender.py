#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook 群發郵件工具
功能：從 Excel 文件或剪貼簿讀取收件人資訊，透過 Outlook 發送個性化郵件
作者：AI Assistant
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import re
import os
import threading
import time
from typing import List, Dict, Tuple, Optional
import logging

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

try:
    import win32com.client
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False
    logger.warning("win32com.client not available. Outlook integration disabled.")

try:
    import pyperclip
    CLIPBOARD_AVAILABLE = True
except ImportError:
    CLIPBOARD_AVAILABLE = False
    logger.warning("pyperclip not available. Clipboard functionality may be limited.")


class EmailRecipient:
    """郵件收件人類別"""
    def __init__(self, name: str = "", email: str = ""):
        self.name = name.strip() if name else ""
        self.email = email.strip() if email else ""
    
    def __str__(self):
        if self.name:
            return f"{self.name} <{self.email}>"
        return self.email
    
    def is_valid(self) -> bool:
        """檢查郵件地址是否有效"""
        email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(email_pattern, self.email))


class DataProcessor:
    """資料處理類別，負責從 Excel 和剪貼簿讀取收件人資訊"""
    
    @staticmethod
    def read_excel_file(file_path: str) -> List[EmailRecipient]:
        """從 Excel 文件讀取收件人資訊"""
        try:
            # 讀取 Excel 文件
            df = pd.read_excel(file_path)
            
            # 自動識別姓名和郵件列
            name_col, email_col = DataProcessor._identify_columns(df)
            
            recipients = []
            for _, row in df.iterrows():
                name = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else ""
                email = str(row[email_col]).strip() if email_col and pd.notna(row[email_col]) else ""
                
                if email and email != 'nan':
                    # 如果沒有姓名，從郵件地址提取
                    if not name or name == 'nan':
                        name = email.split('@')[0]
                    
                    recipient = EmailRecipient(name, email)
                    if recipient.is_valid():
                        recipients.append(recipient)
            
            return recipients
            
        except Exception as e:
            logger.error(f"讀取 Excel 文件時發生錯誤: {e}")
            raise
    
    @staticmethod
    def read_clipboard_data() -> List[EmailRecipient]:
        """從剪貼簿讀取收件人資訊"""
        try:
            if CLIPBOARD_AVAILABLE:
                clipboard_text = pyperclip.paste()
            else:
                # 使用 tkinter 的剪貼簿功能作為備選
                root = tk.Tk()
                root.withdraw()
                try:
                    clipboard_text = root.clipboard_get()
                except tk.TclError:
                    clipboard_text = ""
                root.destroy()
            
            if not clipboard_text.strip():
                return []
            
            return DataProcessor._parse_text_data(clipboard_text)
            
        except Exception as e:
            logger.error(f"讀取剪貼簿資料時發生錯誤: {e}")
            raise
    
    @staticmethod
    def _identify_columns(df: pd.DataFrame) -> Tuple[Optional[str], Optional[str]]:
        """自動識別姓名和郵件列"""
        name_col = None
        email_col = None
        
        # 常見的列名對應
        name_patterns = ['姓名', 'name', '名字', '名稱', '联系人', '聯絡人']
        email_patterns = ['郵件', 'email', 'e-mail', 'mail', '電子郵件', '电子邮件', '邮箱']
        
        # 檢查列名
        for col in df.columns:
            col_lower = str(col).lower().strip()
            
            # 尋找姓名列
            if not name_col:
                for pattern in name_patterns:
                    if pattern in col_lower:
                        name_col = col
                        break
            
            # 尋找郵件列
            if not email_col:
                for pattern in email_patterns:
                    if pattern in col_lower:
                        email_col = col
                        break
        
        # 如果沒有找到明確的列名，嘗試從內容判斷
        if not email_col:
            for col in df.columns:
                # 檢查該列是否包含郵件地址
                sample_values = df[col].dropna().astype(str).head(10)
                email_count = sum(1 for val in sample_values if '@' in val)
                if email_count > len(sample_values) * 0.5:  # 超過50%包含@符號
                    email_col = col
                    break
        
        # 如果仍然沒有找到姓名列，使用第一個非郵件列
        if not name_col and email_col:
            for col in df.columns:
                if col != email_col:
                    name_col = col
                    break
        
        return name_col, email_col
    
    @staticmethod
    def _parse_text_data(text: str) -> List[EmailRecipient]:
        """解析文字資料"""
        recipients = []
        lines = text.strip().split('\n')
        
        # 檢查第一行是否為標題
        first_line = lines[0].strip()
        has_header = bool(re.search(r'(姓名|name|郵件|email|e-mail)', first_line, re.IGNORECASE))
        
        start_index = 1 if has_header else 0
        
        for line in lines[start_index:]:
            line = line.strip()
            if not line:
                continue
            
            # 嘗試不同的分隔符
            parts = []
            for separator in ['\t', ',', ';', '|']:
                parts = [part.strip() for part in line.split(separator) if part.strip()]
                if len(parts) >= 2:
                    break
            
            if not parts:
                # 如果沒有分隔符，檢查是否只是一個郵件地址
                if '@' in line:
                    email = line.strip()
                    name = email.split('@')[0]
                    recipient = EmailRecipient(name, email)
                    if recipient.is_valid():
                        recipients.append(recipient)
                continue
            
            # 識別姓名和郵件
            name = ""
            email = ""
            
            for part in parts:
                if '@' in part and not email:
                    email = part
                elif not name:
                    name = part
            
            if email:
                if not name:
                    name = email.split('@')[0]
                
                recipient = EmailRecipient(name, email)
                if recipient.is_valid():
                    recipients.append(recipient)
        
        return recipients


class OutlookMailSender:
    """Outlook 郵件發送器"""
    
    def __init__(self):
        self.outlook = None
        self._initialize_outlook()
    
    def _initialize_outlook(self):
        """初始化 Outlook COM 物件"""
        if not OUTLOOK_AVAILABLE:
            raise Exception("win32com.client 未安裝，無法連接 Outlook")
        
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            logger.info("Outlook COM 物件初始化成功")
        except Exception as e:
            logger.error(f"無法連接到 Outlook: {e}")
            raise
    
    def send_emails(self, recipients: List[EmailRecipient], subject: str, 
                   body: str, attachments: List[str] = None, 
                   progress_callback=None, stop_event=None) -> Tuple[int, int, List[str]]:
        """
        發送郵件給所有收件人
        返回：(成功數量, 失敗數量, 錯誤列表)
        """
        if not self.outlook:
            raise Exception("Outlook 未正確初始化")
        
        success_count = 0
        failed_count = 0
        errors = []
        
        total = len(recipients)
        
        for i, recipient in enumerate(recipients):
            if stop_event and stop_event.is_set():
                break
            
            try:
                # 創建郵件
                mail = self.outlook.CreateItem(0)  # 0 = olMailItem
                
                # 設定收件人（單獨發送，確保隱私）
                mail.To = recipient.email
                
                # 設定主題
                mail.Subject = subject
                
                # 設定內容（支援變數替換）
                personalized_body = self._personalize_content(body, recipient)
                mail.Body = personalized_body
                
                # 添加附件
                if attachments:
                    for attachment_path in attachments:
                        if os.path.exists(attachment_path):
                            mail.Attachments.Add(attachment_path)
                
                # 發送郵件
                mail.Send()
                success_count += 1
                
                logger.info(f"郵件已發送給: {recipient}")
                
            except Exception as e:
                failed_count += 1
                error_msg = f"發送給 {recipient} 時失敗: {str(e)}"
                errors.append(error_msg)
                logger.error(error_msg)
            
            # 更新進度
            if progress_callback:
                progress_callback(i + 1, total, recipient.name or recipient.email)
            
            # 短暫延遲，避免過快發送
            time.sleep(0.1)
        
        return success_count, failed_count, errors
    
    def preview_email(self, recipient: EmailRecipient, subject: str, body: str) -> Tuple[str, str, str]:
        """預覽郵件內容"""
        personalized_body = self._personalize_content(body, recipient)
        return recipient.email, subject, personalized_body
    
    def _personalize_content(self, content: str, recipient: EmailRecipient) -> str:
        """個性化郵件內容"""
        name = recipient.name if recipient.name else "您好"
        
        # 替換變數
        personalized = content.replace("{姓名}", name)
        personalized = personalized.replace("{name}", name)
        personalized = personalized.replace("{Name}", name)
        personalized = personalized.replace("{NAME}", name.upper())
        
        return personalized


class OutlookMailSenderGUI:
    """GUI 主類別"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Outlook 群發郵件工具")
        self.root.geometry("800x700")
        
        # 資料
        self.recipients = []
        self.attachments = []
        self.mail_sender = None
        self.send_thread = None
        self.stop_event = threading.Event()
        
        # 初始化 GUI
        self._setup_gui()
        
        # 初始化郵件發送器
        self._initialize_mail_sender()
    
    def _setup_gui(self):
        """設置GUI界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 標題
        title_label = ttk.Label(main_frame, text="Outlook 群發郵件工具", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # 收件人區域
        self._setup_recipients_section(main_frame)
        
        # 郵件內容區域
        self._setup_email_content_section(main_frame)
        
        # 附件區域
        self._setup_attachments_section(main_frame)
        
        # 操作按鈕區域
        self._setup_actions_section(main_frame)
        
        # 狀態欄
        self._setup_status_section(main_frame)
        
        # 配置網格權重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
    
    def _setup_recipients_section(self, parent):
        """設置收件人區域"""
        # 收件人標籤
        recipients_label = ttk.Label(parent, text="收件人管理", font=('Arial', 12, 'bold'))
        recipients_label.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # 按鈕框架
        buttons_frame = ttk.Frame(parent)
        buttons_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 選擇Excel按鈕
        self.excel_btn = ttk.Button(buttons_frame, text="選擇 Excel 文件", 
                                   command=self._load_excel_file)
        self.excel_btn.grid(row=0, column=0, padx=(0, 10))
        
        # 從剪貼簿貼上按鈕
        self.clipboard_btn = ttk.Button(buttons_frame, text="從剪貼簿貼上", 
                                       command=self._load_clipboard_data)
        self.clipboard_btn.grid(row=0, column=1, padx=(0, 10))
        
        # 清空收件人按鈕
        self.clear_btn = ttk.Button(buttons_frame, text="清空收件人", 
                                   command=self._clear_recipients)
        self.clear_btn.grid(row=0, column=2)
        
        # 收件人列表
        list_frame = ttk.Frame(parent)
        list_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # 創建 Treeview 來顯示收件人
        self.recipients_tree = ttk.Treeview(list_frame, columns=('name', 'email'), 
                                           show='headings', height=6)
        self.recipients_tree.heading('name', text='姓名')
        self.recipients_tree.heading('email', text='郵件地址')
        self.recipients_tree.column('name', width=150)
        self.recipients_tree.column('email', width=250)
        
        # 滾動條
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.recipients_tree.yview)
        self.recipients_tree.configure(yscrollcommand=scrollbar.set)
        
        self.recipients_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
    
    def _setup_email_content_section(self, parent):
        """設置郵件內容區域"""
        content_label = ttk.Label(parent, text="郵件內容", font=('Arial', 12, 'bold'))
        content_label.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # 主題
        subject_frame = ttk.Frame(parent)
        subject_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        subject_frame.columnconfigure(1, weight=1)
        
        ttk.Label(subject_frame, text="主題:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.subject_entry = ttk.Entry(subject_frame)
        self.subject_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # 內容
        ttk.Label(parent, text="內容 (支援變數 {姓名}):").grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        self.content_text = scrolledtext.ScrolledText(parent, height=8, wrap=tk.WORD)
        self.content_text.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        
        # 設定預設內容
        default_content = """親愛的 {姓名}，

您好！這是一封測試郵件。

祝好！"""
        self.content_text.insert(tk.END, default_content)
    
    def _setup_attachments_section(self, parent):
        """設置附件區域"""
        attachments_frame = ttk.Frame(parent)
        attachments_frame.grid(row=8, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        attachments_frame.columnconfigure(1, weight=1)
        
        # 附件勾選框和按鈕
        self.attachments_var = tk.BooleanVar()
        attachments_check = ttk.Checkbutton(attachments_frame, text="包含附件", 
                                          variable=self.attachments_var,
                                          command=self._toggle_attachments)
        attachments_check.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.attachments_btn = ttk.Button(attachments_frame, text="選擇附件", 
                                         command=self._select_attachments, state='disabled')
        self.attachments_btn.grid(row=0, column=1, sticky=tk.W, padx=(0, 10))
        
        # 附件列表
        self.attachments_listbox = tk.Listbox(attachments_frame, height=3, state='disabled')
        self.attachments_listbox.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
    
    def _setup_actions_section(self, parent):
        """設置操作按鈕區域"""
        actions_frame = ttk.Frame(parent)
        actions_frame.grid(row=9, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # 預覽按鈕
        self.preview_btn = ttk.Button(actions_frame, text="預覽郵件", 
                                     command=self._preview_email)
        self.preview_btn.grid(row=0, column=0, padx=(0, 10))
        
        # 發送按鈕
        self.send_btn = ttk.Button(actions_frame, text="發送郵件", 
                                  command=self._send_emails)
        self.send_btn.grid(row=0, column=1, padx=(0, 10))
        
        # 停止按鈕
        self.stop_btn = ttk.Button(actions_frame, text="停止發送", 
                                  command=self._stop_sending, state='disabled')
        self.stop_btn.grid(row=0, column=2)
    
    def _setup_status_section(self, parent):
        """設置狀態區域"""
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=10, column=0, columnspan=2, sticky=(tk.W, tk.E))
        status_frame.columnconfigure(0, weight=1)
        
        # 狀態標籤
        self.status_var = tk.StringVar(value="就緒")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, 
                               relief=tk.SUNKEN, anchor=tk.W)
        status_label.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # 進度條
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, 
                                          maximum=100)
        self.progress_bar.grid(row=0, column=1, sticky=(tk.W, tk.E))
    
    def _initialize_mail_sender(self):
        """初始化郵件發送器"""
        try:
            self.mail_sender = OutlookMailSender()
            self.status_var.set("Outlook 連接成功，就緒")
        except Exception as e:
            self.status_var.set(f"Outlook 連接失敗: {str(e)}")
            messagebox.showerror("錯誤", f"無法連接到 Outlook:\n{str(e)}")
    
    def _load_excel_file(self):
        """載入 Excel 文件"""
        file_path = filedialog.askopenfilename(
            title="選擇 Excel 文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.status_var.set("正在讀取 Excel 文件...")
                self.root.update()
                
                recipients = DataProcessor.read_excel_file(file_path)
                self._update_recipients_list(recipients)
                
                self.status_var.set(f"已從 Excel 載入 {len(recipients)} 位收件人")
                
            except Exception as e:
                self.status_var.set("Excel 讀取失敗")
                messagebox.showerror("錯誤", f"讀取 Excel 文件時發生錯誤:\n{str(e)}")
    
    def _load_clipboard_data(self):
        """載入剪貼簿資料"""
        try:
            self.status_var.set("正在讀取剪貼簿資料...")
            self.root.update()
            
            recipients = DataProcessor.read_clipboard_data()
            
            if not recipients:
                messagebox.showwarning("警告", "剪貼簿中未找到有效的收件人資料")
                self.status_var.set("剪貼簿資料為空")
                return
            
            self._update_recipients_list(recipients)
            self.status_var.set(f"已從剪貼簿載入 {len(recipients)} 位收件人")
            
        except Exception as e:
            self.status_var.set("剪貼簿讀取失敗")
            messagebox.showerror("錯誤", f"讀取剪貼簿資料時發生錯誤:\n{str(e)}")
    
    def _update_recipients_list(self, recipients: List[EmailRecipient]):
        """更新收件人列表"""
        # 清空現有列表
        for item in self.recipients_tree.get_children():
            self.recipients_tree.delete(item)
        
        # 添加新收件人
        self.recipients = recipients
        for recipient in recipients:
            self.recipients_tree.insert('', 'end', 
                                      values=(recipient.name, recipient.email))
    
    def _clear_recipients(self):
        """清空收件人列表"""
        if messagebox.askyesno("確認", "確定要清空所有收件人嗎？"):
            self.recipients = []
            for item in self.recipients_tree.get_children():
                self.recipients_tree.delete(item)
            self.status_var.set("收件人列表已清空")
    
    def _toggle_attachments(self):
        """切換附件功能"""
        if self.attachments_var.get():
            self.attachments_btn.config(state='normal')
            self.attachments_listbox.config(state='normal')
        else:
            self.attachments_btn.config(state='disabled')
            self.attachments_listbox.config(state='disabled')
            self.attachments = []
            self.attachments_listbox.delete(0, tk.END)
    
    def _select_attachments(self):
        """選擇附件"""
        files = filedialog.askopenfilenames(
            title="選擇附件",
            filetypes=[("All files", "*.*")]
        )
        
        if files:
            self.attachments = list(files)
            self.attachments_listbox.delete(0, tk.END)
            for file_path in files:
                filename = os.path.basename(file_path)
                self.attachments_listbox.insert(tk.END, filename)
    
    def _preview_email(self):
        """預覽郵件"""
        if not self.recipients:
            messagebox.showwarning("警告", "請先載入收件人列表")
            return
        
        if not self.mail_sender:
            messagebox.showerror("錯誤", "Outlook 未正確初始化")
            return
        
        subject = self.subject_entry.get().strip()
        body = self.content_text.get(1.0, tk.END).strip()
        
        if not subject:
            messagebox.showwarning("警告", "請輸入郵件主題")
            return
        
        if not body:
            messagebox.showwarning("警告", "請輸入郵件內容")
            return
        
        # 預覽第一位收件人的郵件
        recipient = self.recipients[0]
        to_email, preview_subject, preview_body = self.mail_sender.preview_email(
            recipient, subject, body)
        
        # 顯示預覽窗口
        self._show_preview_window(to_email, preview_subject, preview_body)
    
    def _show_preview_window(self, to_email: str, subject: str, body: str):
        """顯示郵件預覽窗口"""
        preview_window = tk.Toplevel(self.root)
        preview_window.title("郵件預覽")
        preview_window.geometry("600x500")
        
        # 收件人
        ttk.Label(preview_window, text="收件人:", font=('Arial', 10, 'bold')).pack(anchor=tk.W, padx=10, pady=(10, 5))
        ttk.Label(preview_window, text=to_email).pack(anchor=tk.W, padx=20, pady=(0, 10))
        
        # 主題
        ttk.Label(preview_window, text="主題:", font=('Arial', 10, 'bold')).pack(anchor=tk.W, padx=10, pady=(0, 5))
        ttk.Label(preview_window, text=subject).pack(anchor=tk.W, padx=20, pady=(0, 10))
        
        # 內容
        ttk.Label(preview_window, text="內容:", font=('Arial', 10, 'bold')).pack(anchor=tk.W, padx=10, pady=(0, 5))
        
        content_frame = ttk.Frame(preview_window)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        content_text = scrolledtext.ScrolledText(content_frame, wrap=tk.WORD, state='disabled')
        content_text.pack(fill=tk.BOTH, expand=True)
        
        content_text.config(state='normal')
        content_text.insert(tk.END, body)
        content_text.config(state='disabled')
        
        # 關閉按鈕
        ttk.Button(preview_window, text="關閉", 
                  command=preview_window.destroy).pack(pady=10)
    
    def _send_emails(self):
        """發送郵件"""
        if not self.recipients:
            messagebox.showwarning("警告", "請先載入收件人列表")
            return
        
        if not self.mail_sender:
            messagebox.showerror("錯誤", "Outlook 未正確初始化")
            return
        
        subject = self.subject_entry.get().strip()
        body = self.content_text.get(1.0, tk.END).strip()
        
        if not subject:
            messagebox.showwarning("警告", "請輸入郵件主題")
            return
        
        if not body:
            messagebox.showwarning("警告", "請輸入郵件內容")
            return
        
        # 確認發送
        count = len(self.recipients)
        if not messagebox.askyesno("確認發送", f"即將發送 {count} 封郵件，確定嗎？"):
            return
        
        # 準備附件
        attachments = self.attachments if self.attachments_var.get() else []
        
        # 禁用相關按鈕
        self.send_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.progress_var.set(0)
        
        # 重置停止事件
        self.stop_event.clear()
        
        # 在新線程中發送郵件
        self.send_thread = threading.Thread(
            target=self._send_emails_thread,
            args=(self.recipients, subject, body, attachments)
        )
        self.send_thread.daemon = True
        self.send_thread.start()
    
    def _send_emails_thread(self, recipients: List[EmailRecipient], subject: str, 
                           body: str, attachments: List[str]):
        """郵件發送線程"""
        def progress_callback(current, total, current_recipient):
            progress = (current / total) * 100
            self.root.after(0, lambda: self._update_progress(progress, current, total, current_recipient))
        
        try:
            success_count, failed_count, errors = self.mail_sender.send_emails(
                recipients, subject, body, attachments, 
                progress_callback, self.stop_event
            )
            
            # 發送完成
            self.root.after(0, lambda: self._on_send_complete(success_count, failed_count, errors))
            
        except Exception as e:
            self.root.after(0, lambda: self._on_send_error(str(e)))
    
    def _update_progress(self, progress: float, current: int, total: int, current_recipient: str):
        """更新進度"""
        self.progress_var.set(progress)
        self.status_var.set(f"正在發送第 {current}/{total} 封郵件給 {current_recipient}")
    
    def _on_send_complete(self, success_count: int, failed_count: int, errors: List[str]):
        """發送完成回調"""
        self.send_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.progress_var.set(100)
        
        if failed_count == 0:
            self.status_var.set(f"全部 {success_count} 封郵件發送成功！")
            messagebox.showinfo("發送完成", f"全部 {success_count} 封郵件已成功發送！")
        else:
            self.status_var.set(f"發送完成：成功 {success_count} 封，失敗 {failed_count} 封")
            
            # 顯示錯誤詳情
            if errors:
                error_msg = "以下郵件發送失敗：\n\n" + "\n".join(errors[:10])
                if len(errors) > 10:
                    error_msg += f"\n... 還有 {len(errors) - 10} 個錯誤"
                
                messagebox.showerror("發送結果", f"成功：{success_count} 封\n失敗：{failed_count} 封\n\n{error_msg}")
    
    def _on_send_error(self, error_msg: str):
        """發送錯誤回調"""
        self.send_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.progress_var.set(0)
        self.status_var.set("發送失敗")
        messagebox.showerror("發送錯誤", f"發送過程中發生錯誤:\n{error_msg}")
    
    def _stop_sending(self):
        """停止發送"""
        if messagebox.askyesno("確認停止", "確定要停止發送郵件嗎？"):
            self.stop_event.set()
            self.status_var.set("正在停止...")
    
    def run(self):
        """運行應用程式"""
        self.root.mainloop()


def main():
    """主函數"""
    app = OutlookMailSenderGUI()
    app.run()


if __name__ == "__main__":
    main()