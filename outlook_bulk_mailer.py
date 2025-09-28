#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook 群發郵件工具
支援從 Excel 文件或剪貼簿讀取收件人資訊，並透過 Outlook 發送個性化郵件
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import win32com.client
import os
import re
import threading
from datetime import datetime
import traceback


class OutlookBulkMailer:
    def __init__(self, root):
        self.root = root
        self.root.title("Outlook 群發郵件工具")
        self.root.geometry("900x750")
        
        # 設定視窗圖示（如果有的話）
        try:
            self.root.iconbitmap(default='icon.ico')
        except:
            pass
        
        # 收件人資料
        self.recipients_data = []
        self.attachments = []
        
        # 建立 GUI
        self.create_widgets()
        
        # 設定視窗置中
        self.center_window()
    
    def center_window(self):
        """將視窗置中於螢幕"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_widgets(self):
        """建立 GUI 元件"""
        # 主標題
        title_frame = tk.Frame(self.root, bg='#2c3e50', height=60)
        title_frame.pack(fill=tk.X, padx=0, pady=0)
        title_label = tk.Label(title_frame, text="📧 Outlook 群發郵件工具", 
                               font=('Microsoft JhengHei', 18, 'bold'),
                               fg='white', bg='#2c3e50')
        title_label.pack(pady=15)
        
        # 主容器
        main_container = tk.Frame(self.root, padx=20, pady=10)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # 1. 資料來源區塊
        source_frame = tk.LabelFrame(main_container, text="📋 資料來源", 
                                    font=('Microsoft JhengHei', 11, 'bold'),
                                    padx=10, pady=10)
        source_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Excel 和剪貼簿按鈕
        btn_frame = tk.Frame(source_frame)
        btn_frame.pack(fill=tk.X)
        
        self.btn_excel = tk.Button(btn_frame, text="📁 選擇 Excel 文件",
                                   font=('Microsoft JhengHei', 10),
                                   command=self.load_excel,
                                   bg='#3498db', fg='white',
                                   padx=20, pady=8)
        self.btn_excel.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_clipboard = tk.Button(btn_frame, text="📋 從剪貼簿貼上",
                                       font=('Microsoft JhengHei', 10),
                                       command=self.load_clipboard,
                                       bg='#27ae60', fg='white',
                                       padx=20, pady=8)
        self.btn_clipboard.pack(side=tk.LEFT)
        
        # 清除按鈕
        self.btn_clear = tk.Button(btn_frame, text="🗑️ 清除資料",
                                   font=('Microsoft JhengHei', 10),
                                   command=self.clear_recipients,
                                   bg='#e74c3c', fg='white',
                                   padx=20, pady=8)
        self.btn_clear.pack(side=tk.LEFT, padx=(10, 0))
        
        # 2. 收件人列表區塊
        recipients_frame = tk.LabelFrame(main_container, text="👥 收件人列表", 
                                        font=('Microsoft JhengHei', 11, 'bold'),
                                        padx=10, pady=10)
        recipients_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 建立 Treeview 來顯示收件人
        columns = ('姓名', '郵件地址')
        self.tree = ttk.Treeview(recipients_frame, columns=columns, show='tree headings', height=8)
        self.tree.heading('#0', text='#')
        self.tree.heading('姓名', text='姓名')
        self.tree.heading('郵件地址', text='郵件地址')
        
        self.tree.column('#0', width=50)
        self.tree.column('姓名', width=200)
        self.tree.column('郵件地址', width=350)
        
        # 滾動條
        scrollbar = ttk.Scrollbar(recipients_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 收件人統計標籤
        self.recipients_count_label = tk.Label(recipients_frame, 
                                               text="共 0 位收件人",
                                               font=('Microsoft JhengHei', 10),
                                               fg='#7f8c8d')
        self.recipients_count_label.pack(pady=(5, 0))
        
        # 3. 郵件內容區塊
        email_frame = tk.LabelFrame(main_container, text="✉️ 郵件內容", 
                                   font=('Microsoft JhengHei', 11, 'bold'),
                                   padx=10, pady=10)
        email_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 主題
        subject_frame = tk.Frame(email_frame)
        subject_frame.pack(fill=tk.X, pady=(0, 10))
        tk.Label(subject_frame, text="主題：", font=('Microsoft JhengHei', 10)).pack(side=tk.LEFT)
        self.subject_entry = tk.Entry(subject_frame, font=('Microsoft JhengHei', 10))
        self.subject_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        
        # 內容
        tk.Label(email_frame, text="內容（支援變數 {姓名}）：", 
                font=('Microsoft JhengHei', 10)).pack(anchor=tk.W)
        
        self.content_text = scrolledtext.ScrolledText(email_frame, 
                                                      font=('Microsoft JhengHei', 10),
                                                      height=8,
                                                      wrap=tk.WORD)
        self.content_text.pack(fill=tk.BOTH, expand=True, pady=(5, 10))
        
        # 預設內容範例
        default_content = """親愛的 {姓名}，

這是您的專屬通知信。

祝好
發件人"""
        self.content_text.insert('1.0', default_content)
        
        # 附件
        attachment_frame = tk.Frame(email_frame)
        attachment_frame.pack(fill=tk.X)
        
        self.attach_var = tk.BooleanVar()
        self.attach_check = tk.Checkbutton(attachment_frame, 
                                           text="包含附件",
                                           font=('Microsoft JhengHei', 10),
                                           variable=self.attach_var,
                                           command=self.toggle_attachment)
        self.attach_check.pack(side=tk.LEFT)
        
        self.attach_button = tk.Button(attachment_frame, 
                                       text="選擇附件",
                                       font=('Microsoft JhengHei', 10),
                                       command=self.select_attachments,
                                       state=tk.DISABLED)
        self.attach_button.pack(side=tk.LEFT, padx=(10, 0))
        
        self.attach_label = tk.Label(attachment_frame, 
                                     text="",
                                     font=('Microsoft JhengHei', 9),
                                     fg='#7f8c8d')
        self.attach_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # 4. 操作按鈕區塊
        action_frame = tk.Frame(main_container)
        action_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.btn_preview = tk.Button(action_frame, 
                                     text="👁️ 預覽郵件",
                                     font=('Microsoft JhengHei', 11, 'bold'),
                                     command=self.preview_email,
                                     bg='#f39c12', fg='white',
                                     padx=30, pady=10)
        self.btn_preview.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_send = tk.Button(action_frame, 
                                 text="📤 發送郵件",
                                 font=('Microsoft JhengHei', 11, 'bold'),
                                 command=self.send_emails,
                                 bg='#27ae60', fg='white',
                                 padx=30, pady=10)
        self.btn_send.pack(side=tk.LEFT)
        
        # 5. 狀態欄
        status_frame = tk.Frame(self.root, bg='#ecf0f1', height=30)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.status_label = tk.Label(status_frame, 
                                     text="就緒",
                                     font=('Microsoft JhengHei', 9),
                                     bg='#ecf0f1',
                                     anchor=tk.W)
        self.status_label.pack(side=tk.LEFT, padx=10, pady=5)
        
        # 進度條
        self.progress = ttk.Progressbar(status_frame, 
                                        length=200,
                                        mode='determinate')
        self.progress.pack(side=tk.RIGHT, padx=10, pady=5)
    
    def load_excel(self):
        """載入 Excel 文件"""
        file_path = filedialog.askopenfilename(
            title="選擇 Excel 文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # 讀取 Excel
            df = pd.read_excel(file_path)
            
            # 識別姓名和郵件欄位
            name_col = None
            email_col = None
            
            # 常見的欄位名稱
            name_patterns = ['姓名', 'name', '名稱', '名字', 'Name', '收件人']
            email_patterns = ['郵件', 'email', 'mail', 'e-mail', 'Email', '電子郵件', '信箱']
            
            for col in df.columns:
                col_lower = str(col).lower()
                
                # 檢查是否為姓名欄位
                if not name_col:
                    for pattern in name_patterns:
                        if pattern.lower() in col_lower:
                            name_col = col
                            break
                
                # 檢查是否為郵件欄位
                if not email_col:
                    for pattern in email_patterns:
                        if pattern.lower() in col_lower:
                            email_col = col
                            break
                    
                    # 如果還沒找到，檢查是否包含 @
                    if not email_col and df[col].astype(str).str.contains('@').any():
                        email_col = col
            
            if not email_col:
                messagebox.showerror("錯誤", "無法識別郵件地址欄位！")
                return
            
            # 提取資料
            self.recipients_data = []
            for _, row in df.iterrows():
                email = str(row[email_col]).strip()
                if '@' in email:  # 確保是有效的郵件地址
                    name = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else email.split('@')[0]
                    self.recipients_data.append({'name': name, 'email': email})
            
            self.update_recipients_display()
            self.status_label.config(text=f"已載入 {len(self.recipients_data)} 位收件人")
            
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取 Excel 失敗：\n{str(e)}")
    
    def load_clipboard(self):
        """從剪貼簿載入資料"""
        try:
            clipboard_text = self.root.clipboard_get()
            
            if not clipboard_text:
                messagebox.showwarning("警告", "剪貼簿是空的！")
                return
            
            # 嘗試解析為表格資料
            lines = clipboard_text.strip().split('\n')
            self.recipients_data = []
            
            # 判斷分隔符（Tab 或逗號）
            separator = '\t' if '\t' in lines[0] else ','
            
            # 檢查第一行是否為標題
            first_line = lines[0].split(separator)
            has_header = not any('@' in cell for cell in first_line)
            
            start_index = 1 if has_header else 0
            
            for line in lines[start_index:]:
                if not line.strip():
                    continue
                
                parts = line.split(separator)
                
                # 尋找郵件地址
                email = None
                name = None
                
                for part in parts:
                    part = part.strip()
                    if '@' in part:
                        email = part
                    elif not name and part:  # 第一個非郵件的欄位作為姓名
                        name = part
                
                if email:
                    if not name:
                        name = email.split('@')[0]
                    self.recipients_data.append({'name': name, 'email': email})
            
            if self.recipients_data:
                self.update_recipients_display()
                self.status_label.config(text=f"已從剪貼簿載入 {len(self.recipients_data)} 位收件人")
            else:
                messagebox.showwarning("警告", "剪貼簿中沒有找到有效的郵件地址！")
            
        except tk.TclError:
            messagebox.showwarning("警告", "剪貼簿是空的！")
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取剪貼簿失敗：\n{str(e)}")
    
    def clear_recipients(self):
        """清除收件人資料"""
        self.recipients_data = []
        self.update_recipients_display()
        self.status_label.config(text="已清除所有收件人")
    
    def update_recipients_display(self):
        """更新收件人顯示列表"""
        # 清除現有項目
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 添加新項目
        for i, recipient in enumerate(self.recipients_data, 1):
            self.tree.insert('', 'end', text=str(i), 
                           values=(recipient['name'], recipient['email']))
        
        # 更新計數
        self.recipients_count_label.config(text=f"共 {len(self.recipients_data)} 位收件人")
    
    def toggle_attachment(self):
        """切換附件選項"""
        if self.attach_var.get():
            self.attach_button.config(state=tk.NORMAL)
        else:
            self.attach_button.config(state=tk.DISABLED)
            self.attachments = []
            self.attach_label.config(text="")
    
    def select_attachments(self):
        """選擇附件"""
        files = filedialog.askopenfilenames(
            title="選擇附件",
            filetypes=[("All files", "*.*")]
        )
        
        if files:
            self.attachments = list(files)
            # 顯示附件數量
            if len(self.attachments) == 1:
                filename = os.path.basename(self.attachments[0])
                self.attach_label.config(text=f"已選擇: {filename}")
            else:
                self.attach_label.config(text=f"已選擇 {len(self.attachments)} 個附件")
    
    def personalize_content(self, content, name):
        """個性化郵件內容"""
        if not name or name == "":
            name = "您好"
        
        # 替換 {姓名} 變數
        personalized = content.replace('{姓名}', name)
        return personalized
    
    def preview_email(self):
        """預覽郵件"""
        if not self.recipients_data:
            messagebox.showwarning("警告", "請先載入收件人資料！")
            return
        
        subject = self.subject_entry.get().strip()
        content = self.content_text.get('1.0', tk.END).strip()
        
        if not subject:
            messagebox.showwarning("警告", "請輸入郵件主題！")
            return
        
        if not content:
            messagebox.showwarning("警告", "請輸入郵件內容！")
            return
        
        # 使用第一個收件人作為預覽
        first_recipient = self.recipients_data[0]
        personalized_content = self.personalize_content(content, first_recipient['name'])
        
        # 建立預覽視窗
        preview_window = tk.Toplevel(self.root)
        preview_window.title("郵件預覽")
        preview_window.geometry("600x500")
        
        # 預覽內容
        preview_frame = tk.Frame(preview_window, padx=20, pady=20)
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # 收件人
        tk.Label(preview_frame, text=f"收件人：{first_recipient['email']}", 
                font=('Microsoft JhengHei', 10, 'bold')).pack(anchor=tk.W, pady=(0, 5))
        
        # 主題
        tk.Label(preview_frame, text=f"主題：{subject}", 
                font=('Microsoft JhengHei', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # 內容
        tk.Label(preview_frame, text="內容：", 
                font=('Microsoft JhengHei', 10, 'bold')).pack(anchor=tk.W, pady=(0, 5))
        
        content_display = scrolledtext.ScrolledText(preview_frame, 
                                                    font=('Microsoft JhengHei', 10),
                                                    height=15,
                                                    wrap=tk.WORD)
        content_display.pack(fill=tk.BOTH, expand=True)
        content_display.insert('1.0', personalized_content)
        content_display.config(state=tk.DISABLED)
        
        # 附件資訊
        if self.attachments:
            attachments_text = "附件：\n" + "\n".join([f"  • {os.path.basename(f)}" for f in self.attachments])
            tk.Label(preview_frame, text=attachments_text, 
                    font=('Microsoft JhengHei', 9),
                    fg='#7f8c8d',
                    justify=tk.LEFT).pack(anchor=tk.W, pady=(10, 0))
        
        # 關閉按鈕
        tk.Button(preview_frame, text="關閉", 
                 command=preview_window.destroy,
                 font=('Microsoft JhengHei', 10),
                 bg='#95a5a6', fg='white',
                 padx=20, pady=5).pack(pady=(10, 0))
    
    def send_emails(self):
        """發送郵件"""
        if not self.recipients_data:
            messagebox.showwarning("警告", "請先載入收件人資料！")
            return
        
        subject = self.subject_entry.get().strip()
        content = self.content_text.get('1.0', tk.END).strip()
        
        if not subject:
            messagebox.showwarning("警告", "請輸入郵件主題！")
            return
        
        if not content:
            messagebox.showwarning("警告", "請輸入郵件內容！")
            return
        
        # 確認對話框
        total_count = len(self.recipients_data)
        result = messagebox.askyesno("確認發送", 
                                     f"即將發送 {total_count} 封郵件，確定嗎？")
        
        if not result:
            return
        
        # 在新執行緒中發送郵件，避免凍結 GUI
        threading.Thread(target=self._send_emails_thread, daemon=True).start()
    
    def _send_emails_thread(self):
        """在背景執行緒中發送郵件"""
        try:
            # 初始化 Outlook
            outlook = win32com.client.Dispatch('Outlook.Application')
            
            subject = self.subject_entry.get().strip()
            content = self.content_text.get('1.0', tk.END).strip()
            total_count = len(self.recipients_data)
            
            # 設定進度條
            self.progress['maximum'] = total_count
            self.progress['value'] = 0
            
            success_count = 0
            failed_list = []
            
            for i, recipient in enumerate(self.recipients_data, 1):
                try:
                    # 更新狀態
                    self.status_label.config(text=f"正在發送第 {i}/{total_count} 封郵件...")
                    
                    # 建立郵件
                    mail = outlook.CreateItem(0)  # 0 = olMailItem
                    
                    # 設定收件人（每封郵件只有一個收件人）
                    mail.To = recipient['email']
                    
                    # 設定主題
                    mail.Subject = subject
                    
                    # 設定個性化內容
                    personalized_content = self.personalize_content(content, recipient['name'])
                    mail.Body = personalized_content
                    
                    # 添加附件
                    if self.attachments:
                        for attachment in self.attachments:
                            if os.path.exists(attachment):
                                mail.Attachments.Add(attachment)
                    
                    # 發送郵件
                    # 為了安全起見，先使用 Display() 顯示，實際使用時改為 Send()
                    mail.Send()  # 如果要測試，可以改為 mail.Display()
                    
                    success_count += 1
                    
                except Exception as e:
                    failed_list.append(f"{recipient['email']}: {str(e)}")
                
                # 更新進度條
                self.progress['value'] = i
                self.root.update_idletasks()
            
            # 完成後的訊息
            if success_count == total_count:
                self.status_label.config(text=f"全部 {total_count} 封郵件已成功發送！")
                messagebox.showinfo("成功", f"全部 {total_count} 封郵件已成功建立並發送！")
            else:
                self.status_label.config(text=f"發送完成：成功 {success_count}/{total_count}")
                
                # 顯示失敗清單
                if failed_list:
                    error_msg = f"成功發送 {success_count} 封，失敗 {len(failed_list)} 封。\n\n失敗清單：\n"
                    error_msg += "\n".join(failed_list[:10])  # 只顯示前10個錯誤
                    if len(failed_list) > 10:
                        error_msg += f"\n... 還有 {len(failed_list) - 10} 個錯誤"
                    
                    messagebox.showwarning("部分失敗", error_msg)
                    
                    # 將錯誤日誌寫入檔案
                    with open(f"email_error_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt", 'w', encoding='utf-8') as f:
                        f.write("郵件發送錯誤日誌\n")
                        f.write(f"時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                        f.write("="*50 + "\n")
                        for error in failed_list:
                            f.write(f"{error}\n")
            
            # 重設進度條
            self.progress['value'] = 0
            
        except Exception as e:
            self.status_label.config(text="發送失敗")
            messagebox.showerror("錯誤", f"無法連接 Outlook：\n{str(e)}\n\n請確認：\n1. 已安裝 Microsoft Outlook\n2. Outlook 已登入帳號\n3. Outlook 正在運行")
            self.progress['value'] = 0


def main():
    """主程式"""
    root = tk.Tk()
    app = OutlookBulkMailer(root)
    
    # 設定關閉事件
    def on_closing():
        if messagebox.askokcancel("退出", "確定要退出程式嗎？"):
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # 啟動主迴圈
    root.mainloop()


if __name__ == "__main__":
    main()