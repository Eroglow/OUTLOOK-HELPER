#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel/剪贴板群发邮件自动化工具
一个专业的Windows桌面应用程序，用于通过Excel文件或剪贴板数据批量发送个性化邮件
"""

import customtkinter as ctk
import pandas as pd
import win32com.client
import pyperclip
from tkinter import filedialog, messagebox, StringVar, BooleanVar
import threading
import time
import re
import os
from typing import Dict, List, Optional, Tuple
import logging
from datetime import datetime

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_automation.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

class EmailAutomationApp:
    """Excel/剪贴板群发邮件自动化工具主类"""
    
    def __init__(self):
        # 初始化主窗口
        self.root = ctk.CTk()
        self.root.title("Excel/剪贴板群发邮件自动化工具")
        self.root.geometry("900x800")
        
        # 设置主题
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        
        # 数据存储
        self.data = None
        self.columns = []
        self.email_column = StringVar()
        self.name_column = StringVar()
        self.is_sending = BooleanVar(value=False)
        self.should_stop = BooleanVar(value=False)
        
        # 发送统计
        self.total_emails = 0
        self.sent_count = 0
        self.failed_count = 0
        
        # 附件路径
        self.attachment_path = None
        
        self.setup_ui()
        
    def setup_ui(self):
        """设置用户界面"""
        # 主容器
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 标题
        title_label = ctk.CTkLabel(
            main_frame, 
            text="Excel/剪贴板群发邮件自动化工具",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=(0, 20))
        
        # 数据输入区域
        self.setup_data_input_section(main_frame)
        
        # 数据预览区域
        self.setup_data_preview_section(main_frame)
        
        # 字段映射区域
        self.setup_field_mapping_section(main_frame)
        
        # 邮件设置区域
        self.setup_email_settings_section(main_frame)
        
        # 控制按钮区域
        self.setup_control_buttons_section(main_frame)
        
        # 进度显示区域
        self.setup_progress_section(main_frame)
        
        # 日志输出区域
        self.setup_log_section(main_frame)
        
    def setup_data_input_section(self, parent):
        """设置数据输入区域"""
        input_frame = ctk.CTkFrame(parent)
        input_frame.pack(fill="x", pady=(0, 10))
        
        input_label = ctk.CTkLabel(input_frame, text="数据输入", font=ctk.CTkFont(size=16, weight="bold"))
        input_label.pack(pady=(10, 5))
        
        button_frame = ctk.CTkFrame(input_frame)
        button_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        # 选择文件按钮
        self.file_button = ctk.CTkButton(
            button_frame,
            text="选择Excel/CSV文件",
            command=self.select_file,
            width=200
        )
        self.file_button.pack(side="left", padx=(0, 10))
        
        # 从剪贴板导入按钮
        self.clipboard_button = ctk.CTkButton(
            button_frame,
            text="从剪贴板导入",
            command=self.import_from_clipboard,
            width=200
        )
        self.clipboard_button.pack(side="left")
        
        # 文件路径显示
        self.file_path_var = StringVar(value="未选择文件")
        self.file_path_label = ctk.CTkLabel(input_frame, textvariable=self.file_path_var)
        self.file_path_label.pack(pady=(0, 10))
        
    def setup_data_preview_section(self, parent):
        """设置数据预览区域"""
        preview_frame = ctk.CTkFrame(parent)
        preview_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        preview_label = ctk.CTkLabel(preview_frame, text="数据预览", font=ctk.CTkFont(size=16, weight="bold"))
        preview_label.pack(pady=(10, 5))
        
        # 数据预览文本框
        self.preview_text = ctk.CTkTextbox(preview_frame, height=120)
        self.preview_text.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        
    def setup_field_mapping_section(self, parent):
        """设置字段映射区域"""
        mapping_frame = ctk.CTkFrame(parent)
        mapping_frame.pack(fill="x", pady=(0, 10))
        
        mapping_label = ctk.CTkLabel(mapping_frame, text="字段映射", font=ctk.CTkFont(size=16, weight="bold"))
        mapping_label.pack(pady=(10, 5))
        
        mapping_content = ctk.CTkFrame(mapping_frame)
        mapping_content.pack(fill="x", padx=20, pady=(0, 10))
        
        # 电子邮箱字段映射
        email_frame = ctk.CTkFrame(mapping_content)
        email_frame.pack(fill="x", pady=(0, 5))
        
        ctk.CTkLabel(email_frame, text="电子邮箱字段:", width=120).pack(side="left", padx=(10, 5))
        self.email_combobox = ctk.CTkComboBox(email_frame, variable=self.email_column, width=200)
        self.email_combobox.pack(side="left", padx=(0, 10))
        
        # 姓名字段映射
        name_frame = ctk.CTkFrame(mapping_content)
        name_frame.pack(fill="x", pady=(0, 5))
        
        ctk.CTkLabel(name_frame, text="姓名字段:", width=120).pack(side="left", padx=(10, 5))
        self.name_combobox = ctk.CTkComboBox(name_frame, variable=self.name_column, width=200)
        self.name_combobox.pack(side="left", padx=(0, 10))
        
    def setup_email_settings_section(self, parent):
        """设置邮件设置区域"""
        email_frame = ctk.CTkFrame(parent)
        email_frame.pack(fill="x", pady=(0, 10))
        
        email_label = ctk.CTkLabel(email_frame, text="邮件设置", font=ctk.CTkFont(size=16, weight="bold"))
        email_label.pack(pady=(10, 5))
        
        settings_content = ctk.CTkFrame(email_frame)
        settings_content.pack(fill="x", padx=20, pady=(0, 10))
        
        # 邮件主题
        subject_frame = ctk.CTkFrame(settings_content)
        subject_frame.pack(fill="x", pady=(0, 10))
        
        ctk.CTkLabel(subject_frame, text="邮件主题:", width=120).pack(side="left", padx=(10, 5))
        self.subject_entry = ctk.CTkEntry(subject_frame, placeholder_text="请输入邮件主题")
        self.subject_entry.pack(fill="x", padx=(0, 10), pady=10)
        
        # 邮件内容
        content_label = ctk.CTkLabel(settings_content, text="邮件内容 (支持变量: {姓名}, {电子邮箱}, {公司}, {职位} 等):")
        content_label.pack(anchor="w", padx=10)
        
        # 模板选择按钮
        template_frame = ctk.CTkFrame(settings_content)
        template_frame.pack(fill="x", padx=10, pady=(0, 5))
        
        ctk.CTkLabel(template_frame, text="选择模板:", width=100).pack(side="left", padx=(10, 5))
        
        template_button = ctk.CTkButton(
            template_frame,
            text="商务邮件",
            command=lambda: self.load_template("business"),
            width=80
        )
        template_button.pack(side="left", padx=(0, 5))
        
        template_button2 = ctk.CTkButton(
            template_frame,
            text="感谢邮件", 
            command=lambda: self.load_template("thanks"),
            width=80
        )
        template_button2.pack(side="left", padx=(0, 5))
        
        template_button3 = ctk.CTkButton(
            template_frame,
            text="邀请邮件",
            command=lambda: self.load_template("invitation"),
            width=80
        )
        template_button3.pack(side="left", padx=(0, 5))
        
        self.content_text = ctk.CTkTextbox(settings_content, height=100)
        self.content_text.pack(fill="x", padx=10, pady=(0, 5))
        self.content_text.insert("1.0", "亲爱的 {姓名}，\n\n这是一封测试邮件。\n\n谢谢！")
        
        # 附件选择
        attachment_frame = ctk.CTkFrame(settings_content)
        attachment_frame.pack(fill="x", padx=10, pady=(5, 10))
        
        ctk.CTkLabel(attachment_frame, text="附件:", width=100).pack(side="left", padx=(10, 5))
        
        self.attachment_path_var = StringVar(value="未选择附件")
        self.attachment_label = ctk.CTkLabel(attachment_frame, textvariable=self.attachment_path_var)
        self.attachment_label.pack(side="left", padx=(0, 10))
        
        attachment_button = ctk.CTkButton(
            attachment_frame,
            text="选择附件",
            command=self.select_attachment,
            width=80
        )
        attachment_button.pack(side="right", padx=(0, 10))
        
        clear_attachment_button = ctk.CTkButton(
            attachment_frame,
            text="清除",
            command=self.clear_attachment,
            width=60
        )
        clear_attachment_button.pack(side="right", padx=(0, 5))
        
    def setup_control_buttons_section(self, parent):
        """设置控制按钮区域"""
        control_frame = ctk.CTkFrame(parent)
        control_frame.pack(fill="x", pady=(0, 10))
        
        button_container = ctk.CTkFrame(control_frame)
        button_container.pack(pady=10)
        
        # 预览邮件按钮
        self.preview_button = ctk.CTkButton(
            button_container,
            text="预览邮件",
            command=self.preview_email,
            width=120
        )
        self.preview_button.pack(side="left", padx=(0, 10))
        
        # 开始发送按钮
        self.send_button = ctk.CTkButton(
            button_container,
            text="开始发送",
            command=self.start_sending,
            width=120,
            fg_color="green",
            hover_color="darkgreen"
        )
        self.send_button.pack(side="left", padx=(0, 10))
        
        # 停止发送按钮
        self.stop_button = ctk.CTkButton(
            button_container,
            text="停止发送",
            command=self.stop_sending,
            width=120,
            fg_color="red",
            hover_color="darkred",
            state="disabled"
        )
        self.stop_button.pack(side="left")
        
    def setup_progress_section(self, parent):
        """设置进度显示区域"""
        progress_frame = ctk.CTkFrame(parent)
        progress_frame.pack(fill="x", pady=(0, 10))
        
        progress_label = ctk.CTkLabel(progress_frame, text="发送进度", font=ctk.CTkFont(size=16, weight="bold"))
        progress_label.pack(pady=(10, 5))
        
        # 进度条
        self.progress_bar = ctk.CTkProgressBar(progress_frame)
        self.progress_bar.pack(fill="x", padx=20, pady=(0, 10))
        self.progress_bar.set(0)
        
        # 进度信息
        self.progress_info = StringVar(value="准备就绪")
        progress_info_label = ctk.CTkLabel(progress_frame, textvariable=self.progress_info)
        progress_info_label.pack(pady=(0, 10))
        
    def setup_log_section(self, parent):
        """设置日志输出区域"""
        log_frame = ctk.CTkFrame(parent)
        log_frame.pack(fill="both", expand=True)
        
        log_label = ctk.CTkLabel(log_frame, text="执行日志", font=ctk.CTkFont(size=16, weight="bold"))
        log_label.pack(pady=(10, 5))
        
        # 日志文本框
        self.log_text = ctk.CTkTextbox(log_frame, height=150)
        self.log_text.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        
    def select_file(self):
        """选择Excel或CSV文件"""
        file_types = [
            ("Excel files", "*.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="选择Excel或CSV文件",
            filetypes=file_types
        )
        
        if filename:
            try:
                self.load_file_data(filename)
                self.file_path_var.set(f"已选择: {os.path.basename(filename)}")
                self.log_message(f"成功加载文件: {filename}")
            except Exception as e:
                messagebox.showerror("错误", f"加载文件失败: {str(e)}")
                self.log_message(f"加载文件失败: {str(e)}")
                
    def load_file_data(self, filename: str):
        """加载文件数据"""
        file_ext = os.path.splitext(filename)[1].lower()
        
        if file_ext in ['.xlsx', '.xls']:
            self.data = pd.read_excel(filename)
        elif file_ext == '.csv':
            # 尝试不同编码
            encodings = ['utf-8', 'gbk', 'gb2312', 'utf-8-sig']
            for encoding in encodings:
                try:
                    self.data = pd.read_csv(filename, encoding=encoding)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                raise ValueError("无法使用常见编码读取CSV文件")
        else:
            raise ValueError("不支持的文件格式")
            
        self.update_data_preview()
        self.update_column_mapping()
        
    def import_from_clipboard(self):
        """从剪贴板导入数据"""
        try:
            # 获取剪贴板内容
            clipboard_content = pyperclip.paste()
            
            if not clipboard_content.strip():
                messagebox.showwarning("警告", "剪贴板为空")
                return
                
            # 尝试解析为DataFrame
            from io import StringIO
            self.data = pd.read_csv(StringIO(clipboard_content), sep='\t')
            
            # 如果只有一列，尝试其他分隔符
            if len(self.data.columns) == 1:
                self.data = pd.read_csv(StringIO(clipboard_content), sep=',')
                
            self.update_data_preview()
            self.update_column_mapping()
            self.file_path_var.set("已从剪贴板导入数据")
            self.log_message("成功从剪贴板导入数据")
            
        except Exception as e:
            messagebox.showerror("错误", f"从剪贴板导入失败: {str(e)}")
            self.log_message(f"从剪贴板导入失败: {str(e)}")
            
    def update_data_preview(self):
        """更新数据预览"""
        if self.data is not None:
            # 显示前5行数据
            preview_data = self.data.head(5)
            preview_text = preview_data.to_string(index=False, max_cols=6)
            
            self.preview_text.delete("1.0", "end")
            self.preview_text.insert("1.0", preview_text)
            
    def update_column_mapping(self):
        """更新字段映射下拉框"""
        if self.data is not None:
            self.columns = list(self.data.columns)
            
            # 更新下拉框选项
            self.email_combobox.configure(values=self.columns)
            self.name_combobox.configure(values=self.columns)
            
            # 自动检测可能的邮箱和姓名字段
            self.auto_detect_columns()
            
    def auto_detect_columns(self):
        """自动检测邮箱和姓名字段"""
        email_keywords = ['email', 'mail', '邮箱', '邮件', 'e-mail']
        name_keywords = ['name', 'user', '姓名', '用户', '客户', '联系人']
        
        # 自动检测邮箱字段
        for col in self.columns:
            col_lower = col.lower()
            if any(keyword in col_lower for keyword in email_keywords):
                self.email_column.set(col)
                break
                
        # 自动检测姓名字段
        for col in self.columns:
            col_lower = col.lower()
            if any(keyword in col_lower for keyword in name_keywords):
                self.name_column.set(col)
                break
                
    def validate_email(self, email: str) -> bool:
        """验证邮箱格式"""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, str(email)) is not None
        
    def replace_variables(self, template: str, row: pd.Series) -> str:
        """替换模板中的变量"""
        result = template
        
        # 替换姓名变量
        if self.name_column.get() and self.name_column.get() in row.index:
            name_value = str(row[self.name_column.get()])
            result = result.replace('{姓名}', name_value)
            
        # 替换邮箱变量
        if self.email_column.get() and self.email_column.get() in row.index:
            email_value = str(row[self.email_column.get()])
            result = result.replace('{电子邮箱}', email_value)
            
        # 替换其他字段变量
        for col in self.columns:
            placeholder = f'{{{col}}}'
            if placeholder in result:
                value = str(row[col]) if pd.notna(row[col]) else ''
                result = result.replace(placeholder, value)
                
        return result
        
    def preview_email(self):
        """预览邮件"""
        if not self.validate_settings():
            return
            
        try:
            # 获取第一行数据进行预览
            first_row = self.data.iloc[0]
            
            subject = self.replace_variables(self.subject_entry.get(), first_row)
            content = self.replace_variables(self.content_text.get("1.0", "end-1c"), first_row)
            email_address = str(first_row[self.email_column.get()])
            
            preview_text = f"""收件人: {email_address}
主题: {subject}

内容:
{content}"""
            
            # 创建预览窗口
            preview_window = ctk.CTkToplevel(self.root)
            preview_window.title("邮件预览")
            preview_window.geometry("600x400")
            
            preview_textbox = ctk.CTkTextbox(preview_window)
            preview_textbox.pack(fill="both", expand=True, padx=20, pady=20)
            preview_textbox.insert("1.0", preview_text)
            
        except Exception as e:
            messagebox.showerror("错误", f"预览邮件失败: {str(e)}")
            
    def validate_settings(self) -> bool:
        """验证设置"""
        if self.data is None:
            messagebox.showerror("错误", "请先导入数据")
            return False
            
        if not self.email_column.get():
            messagebox.showerror("错误", "请选择电子邮箱字段")
            return False
            
        if not self.subject_entry.get().strip():
            messagebox.showerror("错误", "请输入邮件主题")
            return False
            
        if not self.content_text.get("1.0", "end-1c").strip():
            messagebox.showerror("错误", "请输入邮件内容")
            return False
            
        return True
        
    def start_sending(self):
        """开始发送邮件"""
        if not self.validate_settings():
            return
            
        if self.is_sending.get():
            messagebox.showwarning("警告", "正在发送中，请勿重复点击")
            return
            
        # 重置统计
        self.total_emails = len(self.data)
        self.sent_count = 0
        self.failed_count = 0
        self.should_stop.set(False)
        
        # 更新UI状态
        self.is_sending.set(True)
        self.send_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        
        # 在新线程中发送邮件
        threading.Thread(target=self.send_emails_thread, daemon=True).start()
        
    def stop_sending(self):
        """停止发送邮件"""
        self.should_stop.set(True)
        self.log_message("用户请求停止发送...")
        
    def send_emails_thread(self):
        """发送邮件的线程函数"""
        outlook = None
        try:
            # 初始化Outlook应用程序
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                # 测试连接
                namespace = outlook.GetNamespace("MAPI")
                accounts = namespace.Accounts
                if accounts.Count == 0:
                    raise Exception("未找到已配置的邮箱账户")
                self.log_message(f"成功连接到Outlook，找到 {accounts.Count} 个邮箱账户")
            except Exception as e:
                raise Exception(f"无法连接到Outlook: {str(e)}")
            
            # 预处理数据，检查邮箱格式
            valid_data = []
            for index, row in self.data.iterrows():
                email_address = str(row[self.email_column.get()]).strip()
                if pd.isna(row[self.email_column.get()]) or email_address == 'nan':
                    self.failed_count += 1
                    self.log_message(f"第 {index+1} 行: 邮箱地址为空")
                    continue
                    
                if not self.validate_email(email_address):
                    self.failed_count += 1
                    self.log_message(f"第 {index+1} 行: 无效邮箱格式 {email_address}")
                    continue
                    
                valid_data.append((index, row, email_address))
            
            if not valid_data:
                self.log_message("没有有效的邮箱地址可供发送")
                return
                
            self.log_message(f"开始发送，共 {len(valid_data)} 封有效邮件")
            
            for index, (original_index, row, email_address) in enumerate(valid_data):
                if self.should_stop.get():
                    self.log_message("用户请求停止发送")
                    break
                    
                try:
                    self.log_message(f"正在准备邮件 ({index+1}/{len(valid_data)}): {email_address}")
                    
                    # 创建邮件
                    mail = outlook.CreateItem(0)  # 0 = olMailItem
                    
                    # 设置收件人
                    mail.To = email_address
                    
                    # 设置主题
                    subject = self.replace_variables(self.subject_entry.get(), row)
                    mail.Subject = subject
                    
                    # 设置内容
                    content = self.replace_variables(self.content_text.get("1.0", "end-1c"), row)
                    mail.Body = content
                    
                    # 设置优先级为普通
                    mail.Importance = 1  # 1 = Normal importance
                    
                    # 添加附件（如果有）
                    if self.attachment_path and os.path.exists(self.attachment_path):
                        try:
                            mail.Attachments.Add(self.attachment_path)
                            self.log_message(f"  └─ 已添加附件: {os.path.basename(self.attachment_path)}")
                        except Exception as attach_error:
                            self.log_message(f"  └─ 附件添加失败: {str(attach_error)}")
                    
                    # 发送邮件
                    mail.Send()
                    
                    self.sent_count += 1
                    self.log_message(f"✓ 成功发送到: {email_address}")
                    
                    # 更新进度 - 使用线程安全的方式
                    processed = self.sent_count + self.failed_count
                    progress = processed / self.total_emails
                    self.root.after(0, lambda p=progress: self.progress_bar.set(p))
                    self.root.after(0, lambda: self.progress_info.set(
                        f"已处理 {processed}/{self.total_emails}，"
                        f"成功 {self.sent_count}，失败 {self.failed_count}"
                    ))
                    
                    # 防止发送过快，给Outlook时间处理
                    if not self.should_stop.get():
                        time.sleep(2)  # 增加间隔时间
                    
                except Exception as e:
                    self.failed_count += 1
                    error_msg = str(e)
                    if "被拒绝" in error_msg or "rejected" in error_msg.lower():
                        self.log_message(f"✗ 邮件被拒绝 {email_address}: 可能是邮箱服务器限制")
                    elif "网络" in error_msg or "network" in error_msg.lower():
                        self.log_message(f"✗ 网络错误 {email_address}: {error_msg}")
                    else:
                        self.log_message(f"✗ 发送失败 {email_address}: {error_msg}")
                    
                    # 更新进度
                    processed = self.sent_count + self.failed_count
                    progress = processed / self.total_emails
                    self.root.after(0, lambda p=progress: self.progress_bar.set(p))
                    self.root.after(0, lambda: self.progress_info.set(
                        f"已处理 {processed}/{self.total_emails}，"
                        f"成功 {self.sent_count}，失败 {self.failed_count}"
                    ))
                    
        except Exception as e:
            error_message = f"发送过程出现严重错误: {str(e)}"
            self.log_message(error_message)
            self.root.after(0, lambda: messagebox.showerror("严重错误", error_message))
        finally:
            # 清理资源
            if outlook:
                try:
                    # 不需要显式关闭Outlook，让用户自己管理
                    pass
                except:
                    pass
            # 重置UI状态
            self.root.after(0, self.reset_sending_state)
            
    def reset_sending_state(self):
        """重置发送状态"""
        self.is_sending.set(False)
        self.send_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        
        if self.sent_count + self.failed_count == self.total_emails:
            self.log_message("所有邮件处理完成")
        else:
            self.log_message("发送过程已停止")
            
    def log_message(self, message: str):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        # 在主线程中更新UI
        self.root.after(0, lambda: self._update_log_ui(log_entry))
        
        # 同时写入日志文件
        logging.info(message)
        
    def _update_log_ui(self, log_entry: str):
        """更新日志UI"""
        self.log_text.insert("end", log_entry)
        self.log_text.see("end")
        
    def load_template(self, template_type: str):
        """加载邮件模板"""
        templates = {
            "business": {
                "subject": "重要业务通知",
                "content": """尊敬的 {姓名}，

您好！

我是来自 [您的公司名称] 的 [您的姓名]。

感谢您对我们产品/服务的关注。根据您的需求，我们为您准备了专业的解决方案。

如有任何疑问，请随时联系我们：
- 邮箱：{电子邮箱}
- 电话：[您的电话]

期待与您的进一步合作！

此致
敬礼

[您的姓名]
[您的公司名称]
[您的职位]"""
            },
            "thanks": {
                "subject": "感谢您的参与！",
                "content": """亲爱的 {姓名}，

非常感谢您参与我们的活动！

您的参与对我们非常重要，我们真诚地希望这次体验对您有所帮助。

如果您有任何建议或意见，请随时通过 {电子邮箱} 与我们联系。

再次感谢您的支持！

最诚挚的问候

[您的团队]"""
            },
            "invitation": {
                "subject": "诚挚邀请您参加我们的活动",
                "content": """尊敬的 {姓名}，

我们诚挚地邀请您参加即将举办的活动：

📅 活动时间：[请填写具体时间]
📍 活动地点：[请填写具体地点]
🎯 活动主题：[请填写活动主题]

活动详情：
[请添加活动详细信息]

请确认您的参与意向，我们将为您预留座位。

期待您的光临！

活动联系人：[您的姓名]
联系邮箱：{电子邮箱}
联系电话：[您的电话]

此致
敬礼"""
            }
        }
        
        if template_type in templates:
            template = templates[template_type]
            self.subject_entry.delete(0, "end")
            self.subject_entry.insert(0, template["subject"])
            
            self.content_text.delete("1.0", "end")
            self.content_text.insert("1.0", template["content"])
            
            self.log_message(f"已加载 {template_type} 模板")
            
    def select_attachment(self):
        """选择附件文件"""
        file_types = [
            ("所有文件", "*.*"),
            ("文档文件", "*.pdf *.doc *.docx *.txt"),
            ("图片文件", "*.jpg *.jpeg *.png *.gif *.bmp"),
            ("Excel文件", "*.xlsx *.xls"),
            ("压缩文件", "*.zip *.rar *.7z")
        ]
        
        filename = filedialog.askopenfilename(
            title="选择附件文件",
            filetypes=file_types
        )
        
        if filename:
            self.attachment_path = filename
            self.attachment_path_var.set(f"已选择: {os.path.basename(filename)}")
            self.log_message(f"已选择附件: {filename}")
        else:
            self.attachment_path = None
            
    def clear_attachment(self):
        """清除附件"""
        self.attachment_path = None
        self.attachment_path_var.set("未选择附件")
        self.log_message("已清除附件")
        
    def run(self):
        """运行应用程序"""
        self.log_message("应用程序已启动")
        self.root.mainloop()

def main():
    """主函数"""
    try:
        app = EmailAutomationApp()
        app.run()
    except Exception as e:
        logging.error(f"应用程序启动失败: {str(e)}")
        messagebox.showerror("错误", f"应用程序启动失败: {str(e)}")

if __name__ == "__main__":
    main()