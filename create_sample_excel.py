#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
建立範例 Excel 檔案供測試使用
"""

import pandas as pd
import os

def create_sample_excel():
    """建立範例 Excel 檔案"""
    
    # 範例資料
    data = {
        '姓名': [
            '張三',
            '李四', 
            '王五',
            '趙六',
            '陳七',
            '林八',
            '黃九',
            '周十'
        ],
        '郵件': [
            'zhangsan@example.com',
            'lisi@example.com',
            'wangwu@example.com',
            'zhaoliu@example.com',
            'chenqi@example.com',
            'linba@example.com',
            'huangjiu@example.com',
            'zhoushi@example.com'
        ],
        '部門': [
            '業務部',
            '技術部',
            '人事部',
            '財務部',
            '業務部',
            '技術部',
            '市場部',
            '行政部'
        ],
        '職位': [
            '經理',
            '工程師',
            '主管',
            '會計',
            '業務員',
            '資深工程師',
            '專員',
            '助理'
        ]
    }
    
    # 建立 DataFrame
    df = pd.DataFrame(data)
    
    # 儲存為 Excel 檔案
    filename = 'sample_recipients.xlsx'
    df.to_excel(filename, index=False, engine='openpyxl')
    
    print(f"✓ 已建立範例檔案：{filename}")
    print(f"  檔案位置：{os.path.abspath(filename)}")
    print(f"  收件人數量：{len(df)} 位")
    print("\n範例內容預覽：")
    print(df.head())
    print("\n您可以使用此檔案測試郵件群發功能。")
    print("注意：請將範例郵件地址改為實際的測試郵件地址！")

if __name__ == "__main__":
    create_sample_excel()
    input("\n按 Enter 鍵結束...")