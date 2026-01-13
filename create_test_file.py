#!/usr/bin/env python3
# 创建一个简单的测试Excel文件
# 由于我们没有pandas，我们将使用xlsxwriter库创建Excel文件
import sys

try:
    import xlsxwriter
    workbook = xlsxwriter.Workbook('test_data.xlsx')
    
    # 工作表1：表格数据
    worksheet1 = workbook.add_worksheet('Data Sheet')
    worksheet1.write(0, 0, 'Name')
    worksheet1.write(0, 1, 'Age')
    worksheet1.write(0, 2, 'City')
    worksheet1.write(1, 0, 'Alice')
    worksheet1.write(1, 1, 25)
    worksheet1.write(1, 2, 'New York')
    worksheet1.write(2, 0, 'Bob')
    worksheet1.write(2, 1, 30)
    worksheet1.write(2, 2, 'London')
    worksheet1.write(3, 0, 'Charlie')
    worksheet1.write(3, 1, 35)
    worksheet1.write(3, 2, 'Tokyo')
    
    # 工作表2：表单数据（大部分是空白的）
    worksheet2 = workbook.add_worksheet('Form Sheet')
    worksheet2.write(0, 0, 'Field')
    worksheet2.write(0, 1, 'Value')
    worksheet2.write(1, 0, 'Name')
    worksheet2.write(2, 0, 'Email')
    worksheet2.write(3, 0, 'Age')
    worksheet2.write(4, 0, 'City')
    
    # 工作表3：稀疏数据
    worksheet3 = workbook.add_worksheet('Sparse Sheet')
    worksheet3.write(0, 0, 'A')
    worksheet3.write(0, 1, 'B')
    worksheet3.write(0, 2, 'C')
    worksheet3.write(1, 0, 1)
    worksheet3.write(1, 1, '')
    worksheet3.write(1, 2, 9)
    worksheet3.write(2, 0, 2)
    worksheet3.write(2, 1, 0)
    worksheet3.write(2, 2, 10)
    worksheet3.write(3, 0, 0)
    worksheet3.write(3, 1, 0)
    worksheet3.write(3, 2, 11)
    worksheet3.write(4, 0, 4)
    worksheet3.write(4, 1, 8)
    worksheet3.write(4, 2, 12)
    
    workbook.close()
    print('Created test_data.xlsx with 3 sheets')
except ImportError:
    print('Install xlsxwriter with: pip install xlsxwriter')
    print('Or create a test Excel file manually for testing the Rust program')
    sys.exit(1)