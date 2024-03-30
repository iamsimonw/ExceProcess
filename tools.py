import os
import openpyxl
import csv

def copyCSVtoXlsx(csv_directory,csv_filename,excel_file):
    csv_file = os.path.join(csv_directory, csv_filename)
    excel_file = os.path.join(csv_directory, excel_file)
    sheet_name = os.path.splitext(csv_filename)[0]

    if not os.path.exists(csv_directory):
        os.makedirs(csv_directory)

    # 创建或加载Excel工作簿
    if os.path.exists(excel_file):
        workbook = openpyxl.load_workbook(excel_file)
    else:
        workbook = openpyxl.Workbook()
        
    # 读取CSV文件
    with open(csv_file, 'r', newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        next(reader) 
        data = list(reader)
        
    # 创建或加载工作表
    if sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
    else:
        worksheet = workbook.create_sheet(title=sheet_name)


    # 将数据写入工作表
    for row_idx, row_data in enumerate(data, start=2):
        for col_idx, cell_value in enumerate(row_data, start=1):
            worksheet.cell(row=row_idx, column=col_idx, value=cell_value)

    # 保存工作簿
    workbook.save(excel_file)
    print("CSV文件已成功复制到Excel文件的对应工作表中。")
