import os
import openpyxl
import csv
import datetime

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
        # 检查第一列是否为“合计”，如果是，则跳过当前行
        if row_data[0] == "合计":
            continue
        for col_idx, cell_value in enumerate(row_data, start=1):
            worksheet.cell(row=row_idx, column=col_idx, value=cell_value)

    # 保存工作簿
    workbook.save(excel_file)
    print("CSV文件已成功复制到Excel文件的对应工作表中。")


def processXlsx_ShuiHouHuiKuan(workbook,unit):
    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取收款明细表中M列的求和
        sum_M_column = 0
        if "收款明细表" in workbook.sheetnames:
            worksheet_receipt = workbook["收款明细表"]
            for row in worksheet_receipt.iter_rows(min_row=2, min_col=11, max_col=13, values_only=True):
                receipt_date_str = row[0]  # K列，收款时间字符串
                receipt_amount = row[2]  # M列，收款金额
                try:
                    receipt_date = datetime.datetime.strptime(receipt_date_str, "%Y-%m-%d")
                    if receipt_date >= datetime.datetime(2024, 1, 1):
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_M_column -= float(cleaned_value)
                                else:
                                    sum_M_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_M_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}")
        # 在D4单元格填充求和值
        sum_M_column = sum_M_column / 10000.0
        worksheet_24[unit] = sum_M_column
    return workbook



def processXlsx_ZhiXiaoYingShouQueBao(workbook,unit):
    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取收款明细表中M列的求和
        sum_M_column = 0
        if "应收及分销预测汇总" in workbook.sheetnames:
            worksheet_receipt = workbook["应收及分销预测汇总"]
            for row in worksheet_receipt.iter_rows(min_row=2, min_col=11, max_col=13, values_only=True):
                receipt_date_str = row[0]  # K列，收款时间字符串
                receipt_amount = row[2]  # M列，收款金额
                try:
                    receipt_date = datetime.datetime.strptime(receipt_date_str, "%Y-%m-%d")
                    if receipt_date >= datetime.datetime(2024, 1, 1):
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_M_column -= float(cleaned_value)
                                else:
                                    sum_M_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_M_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}")
        # 在D4单元格填充求和值
        sum_M_column = sum_M_column / 10000.0
        worksheet_24[unit] = sum_M_column
    return workbook