#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
創建範例收件人 Excel 文件
用於測試和示範目的
"""

import pandas as pd
import os

def create_example_excel():
    """創建範例 Excel 文件"""
    
    # 範例資料
    recipients_data = {
        '姓名': [
            '張三',
            '李四', 
            '王五',
            '趙六',
            '陳七',
            'John Smith',
            'Mary Johnson',
            'David Wilson'
        ],
        '郵件': [
            'zhangsan@company.com',
            'lisi@company.com',
            'wangwu@company.com',
            'zhaoliu@company.com',
            'chenqi@company.com',
            'john.smith@example.com',
            'mary.johnson@example.com',
            'david.wilson@example.com'
        ],
        '部門': [
            '資訊部',
            '人事部',
            '財務部',
            '行銷部',
            '研發部',
            'IT Department',
            'HR Department',
            'Sales Department'
        ]
    }
    
    # 創建 DataFrame
    df = pd.DataFrame(recipients_data)
    
    # 儲存為 Excel 文件
    filename = 'example_recipients.xlsx'
    df.to_excel(filename, index=False, engine='openpyxl')
    
    print(f"已創建範例文件: {filename}")
    print("文件內容:")
    print(df.to_string(index=False))
    
    return filename

if __name__ == "__main__":
    # 確保安裝了必要的套件
    try:
        import pandas as pd
        import openpyxl
    except ImportError as e:
        print(f"缺少必要套件: {e}")
        print("請先安裝: pip install pandas openpyxl")
        exit(1)
    
    create_example_excel()