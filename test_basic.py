#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
邮件自动化工具基础功能测试
"""

import pandas as pd
import os
import sys

def test_data_processing():
    """测试数据处理功能"""
    print("🧪 测试数据处理功能...")
    
    # 测试CSV读取
    try:
        df = pd.read_csv('sample_data.csv')
        print(f"✅ CSV文件读取成功，共 {len(df)} 行数据")
        print(f"   列名: {list(df.columns)}")
    except Exception as e:
        print(f"❌ CSV文件读取失败: {e}")
        return False
    
    # 测试数据预览
    try:
        preview = df.head(3)
        print("✅ 数据预览功能正常")
        print(preview.to_string(index=False))
    except Exception as e:
        print(f"❌ 数据预览失败: {e}")
        return False
    
    return True

def test_email_validation():
    """测试邮箱验证功能"""
    print("\n🧪 测试邮箱验证功能...")
    
    import re
    
    def validate_email(email: str) -> bool:
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, str(email)) is not None
    
    test_emails = [
        ("test@example.com", True),
        ("user.name@domain.co.uk", True),
        ("invalid-email", False),
        ("@domain.com", False),
        ("user@", False),
        ("正常邮箱@domain.com", True),
    ]
    
    all_passed = True
    for email, expected in test_emails:
        result = validate_email(email)
        if result == expected:
            print(f"✅ {email} -> {result}")
        else:
            print(f"❌ {email} -> {result} (期望: {expected})")
            all_passed = False
    
    return all_passed

def test_template_replacement():
    """测试模板变量替换"""
    print("\n🧪 测试模板变量替换功能...")
    
    # 模拟数据行
    row_data = pd.Series({
        '姓名': '张小明',
        '电子邮箱': 'xiaoming@example.com',
        '公司': 'ABC科技',
        '职位': '软件工程师'
    })
    
    def replace_variables(template: str, row: pd.Series) -> str:
        result = template
        for col in row.index:
            placeholder = f'{{{col}}}'
            if placeholder in result:
                value = str(row[col]) if pd.notna(row[col]) else ''
                result = result.replace(placeholder, value)
        return result
    
    # 测试模板
    template = "尊敬的 {姓名}，您在 {公司} 担任 {职位}，联系邮箱：{电子邮箱}"
    expected = "尊敬的 张小明，您在 ABC科技 担任 软件工程师，联系邮箱：xiaoming@example.com"
    
    result = replace_variables(template, row_data)
    
    if result == expected:
        print("✅ 变量替换功能正常")
        print(f"   原始: {template}")
        print(f"   结果: {result}")
        return True
    else:
        print(f"❌ 变量替换失败")
        print(f"   期望: {expected}")
        print(f"   实际: {result}")
        return False

def test_requirements():
    """测试依赖包是否正确安装"""
    print("\n🧪 测试依赖包安装情况...")
    
    required_packages = [
        'pandas',
        'customtkinter', 
        'pywin32',
        'openpyxl',
        'pyperclip'
    ]
    
    all_installed = True
    for package in required_packages:
        try:
            __import__(package)
            print(f"✅ {package} 已安装")
        except ImportError:
            print(f"❌ {package} 未安装")
            all_installed = False
    
    return all_installed

def test_file_structure():
    """测试文件结构"""
    print("\n🧪 测试文件结构...")
    
    required_files = [
        'email_automation.py',
        'requirements.txt',
        'start.bat',
        'README.md',
        'sample_data.csv'
    ]
    
    all_exist = True
    for file in required_files:
        if os.path.exists(file):
            print(f"✅ {file} 存在")
        else:
            print(f"❌ {file} 不存在")
            all_exist = False
    
    return all_exist

def main():
    """主测试函数"""
    print("🚀 开始基础功能测试...\n")
    
    tests = [
        ("文件结构", test_file_structure),
        ("依赖包", test_requirements),
        ("数据处理", test_data_processing),
        ("邮箱验证", test_email_validation),
        ("模板替换", test_template_replacement),
    ]
    
    passed = 0
    total = len(tests)
    
    for name, test_func in tests:
        print(f"\n{'='*50}")
        print(f"测试项目: {name}")
        print('='*50)
        
        try:
            if test_func():
                passed += 1
                print(f"✅ {name} 测试通过")
            else:
                print(f"❌ {name} 测试失败")
        except Exception as e:
            print(f"❌ {name} 测试异常: {e}")
    
    print(f"\n{'='*50}")
    print(f"📊 测试总结")
    print('='*50)
    print(f"总测试项: {total}")
    print(f"通过测试: {passed}")
    print(f"失败测试: {total - passed}")
    print(f"通过率: {passed/total*100:.1f}%")
    
    if passed == total:
        print("🎉 所有测试通过！程序可以正常使用。")
        return 0
    else:
        print("⚠️  部分测试失败，请检查相关功能。")
        return 1

if __name__ == "__main__":
    sys.exit(main())