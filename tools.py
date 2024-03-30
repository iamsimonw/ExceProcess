import xlrd

def read_excel_sheets(filename):
    # 打开 Excel 文件
    workbook = xlrd.open_workbook(filename)
    
    # 获取所有的 sheet 名字
    sheet_names = workbook.sheet_names()
    
    # 输出每个 sheet 的名字
    for sheet_name in sheet_names:
        print("Sheet Name:", sheet_name)

# 要读取的 Excel 文件名
filename = r"D:\JB\表(1)\浙江区业绩预测240330-浙南.xls"

# 调用函数读取并输出每个 sheet 的名字
read_excel_sheets(filename)
