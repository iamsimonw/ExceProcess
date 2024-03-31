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

def parse_datetime(input_datetime, row_idx=None):
    if isinstance(input_datetime, datetime.datetime):
        return input_datetime
    elif isinstance(input_datetime, str):
        try:
            return datetime.datetime.strptime(input_datetime, "%Y-%m-%d")
        except ValueError as e:
            if row_idx is not None:
                print(f"日期解析错误：{e}，位于第 {row_idx} 行")
            else:
                print(f"日期解析错误：{e}")
                if row_idx is not None:
                    print(f"类型错误：输入必须是字符串或 datetime.datetime 对象，位于第 {row_idx} 行")
                else:
                    print("类型错误：输入必须是字符串或 datetime.datetime 对象")
            raise TypeError("Input must be a string or datetime.datetime object")

#整体 
def processXlsx_ShuiHouHuiKuan(workbook, unit, dateStart, dateEnd,ratio,product_name):
    # 解析整数形式的日期为日期对象
    receipt_date_start = datetime.datetime.strptime(str(dateStart), "%Y%m%d")
    receipt_date_end = datetime.datetime.strptime(str(dateEnd), "%Y%m%d")
    
    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取收款明细表中M列的求和
        sum_column = 0
        if "收款明细表" in workbook.sheetnames:
            worksheet_receipt = workbook["收款明细表"]
            for row_idx, row in enumerate(worksheet_receipt.iter_rows(min_row=2, min_col=11, max_col=13, values_only=True), start=2):
                # 检查是否为空行
                if not any(row):
                    continue
                receipt_date_str = row[0]  # K列，收款时间字符串
                receipt_amount = row[2]  # M列，收款金额
                try:
                    receipt_date = parse_datetime(receipt_date_str, row_idx=row_idx)
                    if receipt_date_start <= receipt_date <= receipt_date_end:
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_column -= float(cleaned_value)
                                else:
                                    sum_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}，位于工作表 '收款明细表' 的第 {row_idx} 行")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}，位于工作表 '收款明细表' 的第 {row_idx} 行")
        # 在D4单元格填充求和值
        sum_column = sum_column / ratio/10000.0
        worksheet_24[unit] = sum_column
    return workbook

def processXlsx_ZhiXiaoYingShouQueBao(workbook, unit, dateStart, dateEnd,ratio,product_name):
    # 解析整数形式的日期为日期对象
    receipt_date_start = datetime.datetime.strptime(str(dateStart), "%Y%m%d")
    receipt_date_end = datetime.datetime.strptime(str(dateEnd), "%Y%m%d")

    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取应收及分销预测汇总中Q列的求和
        sum_column = 0
        if "应收及分销预测汇总" in workbook.sheetnames:
            worksheet_receipt = workbook["应收及分销预测汇总"]
            for row_idx, row in enumerate(worksheet_receipt.iter_rows(min_row=2, min_col=10, max_col=22, values_only=True), start=2):
                # 检查是否为空行
                if not any(row):
                    continue
                receipt_sales = row[0]  # J列，直销分销
                receipt_date_str = row[6]  # P列，日期
                receipt_amount = row[7]  # Q列，收款金额
                receipt_is_amount = row[12]  # V列，是否未回款
                try:
                    receipt_date = parse_datetime(receipt_date_str, row_idx=row_idx)
                    if receipt_date_start <= receipt_date <= receipt_date_end and receipt_sales == '直销' and receipt_is_amount == '未回款':
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_column -= float(cleaned_value)
                                else:
                                    sum_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
        # 在D4单元格填充求和值
        sum_column = sum_column / ratio / 10000.0
        worksheet_24[unit] = sum_column
    return workbook

def processXlsx_FenXiaoYingShouQueBao(workbook, unit, dateStart, dateEnd,ratio,product_name):
    # 解析整数形式的日期为日期对象
    receipt_date_start = datetime.datetime.strptime(str(dateStart), "%Y%m%d")
    receipt_date_end = datetime.datetime.strptime(str(dateEnd), "%Y%m%d")

    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取应收及分销预测汇总中Q列的求和
        sum_column = 0
        if "应收及分销预测汇总" in workbook.sheetnames:
            worksheet_receipt = workbook["应收及分销预测汇总"]
            for row_idx, row in enumerate(worksheet_receipt.iter_rows(min_row=2, min_col=10, max_col=22, values_only=True), start=2):
                # 检查是否为空行
                if not any(row):
                    continue
                receipt_sales = row[0]  # J列，直销分销
                receipt_date_str = row[6]  # P列，日期
                receipt_amount = row[7]  # Q列，收款金额
                receipt_is_amount = row[12]  # V列，是否未回款
                try:
                    receipt_date = parse_datetime(receipt_date_str, row_idx=row_idx)
                    if receipt_date_start <= receipt_date <= receipt_date_end and receipt_sales == '分销' and receipt_is_amount == '未回款':
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_column -= float(cleaned_value)
                                else:
                                    sum_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
        # 在D4单元格填充求和值
        sum_column = sum_column / ratio / 10000.0
        worksheet_24[unit] = sum_column
    return workbook

def processXlsx_XinQianQueBao(workbook, unit, dateStart, dateEnd,ratio,product_name):
    # 解析整数形式的日期为日期对象
    receipt_date_start = datetime.datetime.strptime(str(dateStart), "%Y%m%d")
    receipt_date_end = datetime.datetime.strptime(str(dateEnd), "%Y%m%d")

    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取项目漏斗汇总-签约金额替重中Y列的求和
        sum_column = 0
        if "项目漏斗汇总-签约金额替重" in workbook.sheetnames:
            worksheet_receipt = workbook["项目漏斗汇总-签约金额替重"]
            for row_idx, row in enumerate(worksheet_receipt.iter_rows(min_row=2, min_col=24, max_col=25, values_only=True), start=2):
                # 检查是否为空行
                if not any(row):
                    continue
                receipt_date_str = row[0]  # X列，日期
                receipt_amount = row[1]  # Y列，回款

                try:
                    receipt_date = parse_datetime(receipt_date_str, row_idx=row_idx)
                    if receipt_date_start <= receipt_date <= receipt_date_end:
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_column -= float(cleaned_value)
                                else:
                                    sum_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
        # 在D4单元格填充求和值
        sum_column = sum_column / ratio / 10000.0
        worksheet_24[unit] = sum_column
    return workbook

def processXlsx_ZhiXiaoYingShouChongCi(workbook, unit, dateStart, dateEnd,ratio,product_name):
    # 解析整数形式的日期为日期对象
    receipt_date_start = datetime.datetime.strptime(str(dateStart), "%Y%m%d")
    receipt_date_end = datetime.datetime.strptime(str(dateEnd), "%Y%m%d")

    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取应收及分销预测汇总中R列的求和
        sum_column = 0
        if "应收及分销预测汇总" in workbook.sheetnames:
            worksheet_receipt = workbook["应收及分销预测汇总"]
            for row_idx, row in enumerate(worksheet_receipt.iter_rows(min_row=2, min_col=10, max_col=22, values_only=True), start=2):
                # 检查是否为空行
                if not any(row):
                    continue
                receipt_sales = row[0]  # J列，直销分销
                receipt_date_str = row[6]  # P列，日期
                receipt_amount = row[8]  # R列，收款金额
                receipt_is_amount = row[12]  # V列，是否未回款
                try:
                    receipt_date = parse_datetime(receipt_date_str, row_idx=row_idx)
                    if receipt_date_start <= receipt_date <= receipt_date_end and receipt_sales == '直销' and receipt_is_amount == '未回款':
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_column -= float(cleaned_value)
                                else:
                                    sum_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
        # 在D4单元格填充求和值
        sum_column = sum_column / ratio / 10000.0
        worksheet_24[unit] = sum_column
    return workbook

def processXlsx_FenXiaoYingShouChongCi(workbook, unit, dateStart, dateEnd,ratio,product_name):
    # 解析整数形式的日期为日期对象
    receipt_date_start = datetime.datetime.strptime(str(dateStart), "%Y%m%d")
    receipt_date_end = datetime.datetime.strptime(str(dateEnd), "%Y%m%d")

    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取应收及分销预测汇总中R列的求和
        sum_column = 0
        if "应收及分销预测汇总" in workbook.sheetnames:
            worksheet_receipt = workbook["应收及分销预测汇总"]
            for row_idx, row in enumerate(worksheet_receipt.iter_rows(min_row=2, min_col=10, max_col=22, values_only=True), start=2):
                # 检查是否为空行
                if not any(row):
                    continue
                receipt_sales = row[0]  # J列，直销分销
                receipt_date_str = row[6]  # P列，日期
                receipt_amount = row[8]  # R列，收款金额
                receipt_is_amount = row[12]  # V列，是否未回款
                try:
                    receipt_date = parse_datetime(receipt_date_str, row_idx=row_idx)
                    if receipt_date_start <= receipt_date <= receipt_date_end and receipt_sales == '分销' and receipt_is_amount == '未回款':
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_column -= float(cleaned_value)
                                else:
                                    sum_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
        # 在D4单元格填充求和值
        sum_column = sum_column / ratio / 10000.0
        worksheet_24[unit] = sum_column
    return workbook

def processXlsx_XinQianChongCi(workbook, unit, dateStart, dateEnd,ratio,product_name):
    # 解析整数形式的日期为日期对象
    receipt_date_start = datetime.datetime.strptime(str(dateStart), "%Y%m%d")
    receipt_date_end = datetime.datetime.strptime(str(dateEnd), "%Y%m%d")

    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取项目漏斗汇总-签约金额替重中Z列的求和
        sum_column = 0
        if "项目漏斗汇总-签约金额替重" in workbook.sheetnames:
            worksheet_receipt = workbook["项目漏斗汇总-签约金额替重"]
            for row_idx, row in enumerate(worksheet_receipt.iter_rows(min_row=2, min_col=24, max_col=26, values_only=True), start=2):
                # 检查是否为空行
                if not any(row):
                    continue
                receipt_date_str = row[0]  # X列，日期
                receipt_amount = row[2]  # Z列，回款

                try:
                    receipt_date = parse_datetime(receipt_date_str, row_idx=row_idx)
                    if receipt_date_start <= receipt_date <= receipt_date_end:
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_column -= float(cleaned_value)
                                else:
                                    sum_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
        # 在D4单元格填充求和值
        sum_column = sum_column / ratio / 10000.0
        worksheet_24[unit] = sum_column
    return workbook

#直销
def processXlsx_ShuiHouHuiKuan_ZhiXiao(workbook, unit, dateStart, dateEnd,ratio,product_name):
    # 解析整数形式的日期为日期对象
    receipt_date_start = datetime.datetime.strptime(str(dateStart), "%Y%m%d")
    receipt_date_end = datetime.datetime.strptime(str(dateEnd), "%Y%m%d")
    
    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取收款明细表中M列的求和
        sum_column = 0
        if "收款明细表" in workbook.sheetnames:
            worksheet_receipt = workbook["收款明细表"]
            for row_idx, row in enumerate(worksheet_receipt.iter_rows(min_row=2, min_col=9, max_col=13, values_only=True), start=2):
                # 检查是否为空行
                if not any(row):
                    continue
                receipt_sales = row[0]  # I列，直销分销
                receipt_date_str = row[2]  # K列，收款时间字符串
                receipt_amount = row[4]  # M列，收款金额
                try:
                    receipt_date = parse_datetime(receipt_date_str, row_idx=row_idx)
                    if receipt_date_start <= receipt_date <= receipt_date_end and receipt_sales=='直销':
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_column -= float(cleaned_value)
                                else:
                                    sum_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}，位于工作表 '收款明细表' 的第 {row_idx} 行")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}，位于工作表 '收款明细表' 的第 {row_idx} 行")
        # 在D4单元格填充求和值
        sum_column = sum_column / ratio/10000.0
        worksheet_24[unit] = sum_column
    return workbook
#分销
def processXlsx_ShuiHouHuiKuan_FenXiao(workbook, unit, dateStart, dateEnd,ratio,product_name):
    # 解析整数形式的日期为日期对象
    receipt_date_start = datetime.datetime.strptime(str(dateStart), "%Y%m%d")
    receipt_date_end = datetime.datetime.strptime(str(dateEnd), "%Y%m%d")
    
    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取收款明细表中M列的求和
        sum_column = 0
        if "收款明细表" in workbook.sheetnames:
            worksheet_receipt = workbook["收款明细表"]
            for row_idx, row in enumerate(worksheet_receipt.iter_rows(min_row=2, min_col=9, max_col=13, values_only=True), start=2):
                # 检查是否为空行
                if not any(row):
                    continue
                receipt_sales = row[0]  # I列，直销分销
                receipt_date_str = row[2]  # K列，收款时间字符串
                receipt_amount = row[4]  # M列，收款金额
                try:
                    receipt_date = parse_datetime(receipt_date_str, row_idx=row_idx)
                    if receipt_date_start <= receipt_date <= receipt_date_end and receipt_sales=='分销':
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_column -= float(cleaned_value)
                                else:
                                    sum_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}，位于工作表 '收款明细表' 的第 {row_idx} 行")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}，位于工作表 '收款明细表' 的第 {row_idx} 行")
        # 在D4单元格填充求和值
        sum_column = sum_column / ratio/10000.0
        worksheet_24[unit] = sum_column
    return workbook

#订阅
def processXlsx_ShuiHouHuiKuan_DingYue(workbook, unit, dateStart, dateEnd,ratio,product_name):
    # 解析整数形式的日期为日期对象
    receipt_date_start = datetime.datetime.strptime(str(dateStart), "%Y%m%d")
    receipt_date_end = datetime.datetime.strptime(str(dateEnd), "%Y%m%d")
    
    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        sum_column = 0
        if "收款明细表" in workbook.sheetnames:
            worksheet_receipt = workbook["收款明细表"]
            for row_idx, row in enumerate(worksheet_receipt.iter_rows(min_row=2, min_col=9, max_col=13, values_only=True), start=2):
                # 检查是否为空行
                if not any(row):
                    continue
                receipt_booking = row[1]  # J列，订阅与否
                receipt_date_str = row[2]  # K列，收款时间字符串
                receipt_amount = row[4]  # M列，收款金额
                try:
                    receipt_date = parse_datetime(receipt_date_str, row_idx=row_idx)
                    if receipt_date_start <= receipt_date <= receipt_date_end and '非订阅' not in receipt_booking:
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_column -= float(cleaned_value)
                                else:
                                    sum_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}，位于工作表 '收款明细表' 的第 {row_idx} 行")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}，位于工作表 '收款明细表' 的第 {row_idx} 行")
        # 在D4单元格填充求和值
        sum_column = sum_column / ratio/10000.0
        worksheet_24[unit] = sum_column
    return workbook

def processXlsx_ZhiXiaoYingShouQueBao_DingYue(workbook, unit, dateStart, dateEnd,ratio,product_name):
    # 解析整数形式的日期为日期对象
    receipt_date_start = datetime.datetime.strptime(str(dateStart), "%Y%m%d")
    receipt_date_end = datetime.datetime.strptime(str(dateEnd), "%Y%m%d")

    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取应收及分销预测汇总中Q列的求和
        sum_column = 0
        if "应收及分销预测汇总" in workbook.sheetnames:
            worksheet_receipt = workbook["应收及分销预测汇总"]
            for row_idx, row in enumerate(worksheet_receipt.iter_rows(min_row=2, min_col=10, max_col=22, values_only=True), start=2):
                # 检查是否为空行
                if not any(row):
                    continue
                receipt_sales = row[0]  # J列，直销分销
                receipt_booking = row[1]  # K列，订阅与否
                receipt_date_str = row[6]  # P列，日期
                receipt_amount = row[7]  # Q列，收款金额
                receipt_is_amount = row[12]  # V列，是否未回款
                try:
                    receipt_date = parse_datetime(receipt_date_str, row_idx=row_idx)
                    if receipt_date_start <= receipt_date <= receipt_date_end and receipt_sales == '直销' and receipt_is_amount == '未回款'and '非订阅' not in receipt_booking:
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_column -= float(cleaned_value)
                                else:
                                    sum_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
        # 在D4单元格填充求和值
        sum_column = sum_column / ratio / 10000.0
        worksheet_24[unit] = sum_column
    return workbook

def processXlsx_FenXiaoYingShouQueBao_DingYue(workbook, unit, dateStart, dateEnd,ratio,product_name):
    # 解析整数形式的日期为日期对象
    receipt_date_start = datetime.datetime.strptime(str(dateStart), "%Y%m%d")
    receipt_date_end = datetime.datetime.strptime(str(dateEnd), "%Y%m%d")

    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取应收及分销预测汇总中Q列的求和
        sum_column = 0
        if "应收及分销预测汇总" in workbook.sheetnames:
            worksheet_receipt = workbook["应收及分销预测汇总"]
            for row_idx, row in enumerate(worksheet_receipt.iter_rows(min_row=2, min_col=10, max_col=22, values_only=True), start=2):
                # 检查是否为空行
                if not any(row):
                    continue
                receipt_sales = row[0]  # J列，直销分销
                receipt_booking = row[1]  # K列，订阅与否
                receipt_date_str = row[6]  # P列，日期
                receipt_amount = row[7]  # Q列，收款金额
                receipt_is_amount = row[12]  # V列，是否未回款
                try:
                    receipt_date = parse_datetime(receipt_date_str, row_idx=row_idx)
                    if receipt_date_start <= receipt_date <= receipt_date_end and receipt_sales == '分销' and receipt_is_amount == '未回款'and '非订阅' not in receipt_booking:
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_column -= float(cleaned_value)
                                else:
                                    sum_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
        # 在D4单元格填充求和值
        sum_column = sum_column / ratio / 10000.0
        worksheet_24[unit] = sum_column
    return workbook

def processXlsx_XinQianQueBao_DingYue(workbook, unit, dateStart, dateEnd,ratio,product_name):
    # 解析整数形式的日期为日期对象
    receipt_date_start = datetime.datetime.strptime(str(dateStart), "%Y%m%d")
    receipt_date_end = datetime.datetime.strptime(str(dateEnd), "%Y%m%d")

    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取项目漏斗汇总-签约金额替重中Y列的求和
        sum_column = 0
        if "项目漏斗汇总-签约金额替重" in workbook.sheetnames:
            worksheet_receipt = workbook["项目漏斗汇总-签约金额替重"]
            for row_idx, row in enumerate(worksheet_receipt.iter_rows(min_row=2, min_col=23, max_col=25, values_only=True), start=2):
                # 检查是否为空行
                if not any(row):
                    continue
                receipt_booking = row[0]  # W列，订阅与否
                receipt_date_str = row[1]  # X列，日期
                receipt_amount = row[2]  # Y列，回款

                try:
                    receipt_date = parse_datetime(receipt_date_str, row_idx=row_idx)
                    if receipt_date_start <= receipt_date <= receipt_date_end and '订阅' in receipt_booking:
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_column -= float(cleaned_value)
                                else:
                                    sum_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
        # 在D4单元格填充求和值
        sum_column = sum_column / ratio / 10000.0
        worksheet_24[unit] = sum_column
    return workbook

def processXlsx_ZhiXiaoYingShouChongCi_DingYue(workbook, unit, dateStart, dateEnd,ratio,product_name):
    # 解析整数形式的日期为日期对象
    receipt_date_start = datetime.datetime.strptime(str(dateStart), "%Y%m%d")
    receipt_date_end = datetime.datetime.strptime(str(dateEnd), "%Y%m%d")

    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取应收及分销预测汇总中R列的求和
        sum_column = 0
        if "应收及分销预测汇总" in workbook.sheetnames:
            worksheet_receipt = workbook["应收及分销预测汇总"]
            for row_idx, row in enumerate(worksheet_receipt.iter_rows(min_row=2, min_col=10, max_col=22, values_only=True), start=2):
                # 检查是否为空行
                if not any(row):
                    continue
                receipt_sales = row[0]  # J列，直销分销
                receipt_booking = row[1]  # K列，订阅与否
                receipt_date_str = row[6]  # P列，日期
                receipt_amount = row[8]  # R列，收款金额
                receipt_is_amount = row[12]  # V列，是否未回款
                try:
                    receipt_date = parse_datetime(receipt_date_str, row_idx=row_idx)
                    if receipt_date_start <= receipt_date <= receipt_date_end and receipt_sales == '直销' and receipt_is_amount == '未回款'and '非订阅' not in receipt_booking:
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_column -= float(cleaned_value)
                                else:
                                    sum_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
        # 在D4单元格填充求和值
        sum_column = sum_column / ratio / 10000.0
        worksheet_24[unit] = sum_column
    return workbook

def processXlsx_FenXiaoYingShouChongCi_DingYue(workbook, unit, dateStart, dateEnd,ratio,product_name):
    # 解析整数形式的日期为日期对象
    receipt_date_start = datetime.datetime.strptime(str(dateStart), "%Y%m%d")
    receipt_date_end = datetime.datetime.strptime(str(dateEnd), "%Y%m%d")

    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取应收及分销预测汇总中R列的求和
        sum_column = 0
        if "应收及分销预测汇总" in workbook.sheetnames:
            worksheet_receipt = workbook["应收及分销预测汇总"]
            for row_idx, row in enumerate(worksheet_receipt.iter_rows(min_row=2, min_col=10, max_col=22, values_only=True), start=2):
                # 检查是否为空行
                if not any(row):
                    continue
                receipt_sales = row[0]  # J列，直销分销
                receipt_booking = row[1]  # K列，订阅与否
                receipt_date_str = row[6]  # P列，日期
                receipt_amount = row[8]  # R列，收款金额
                receipt_is_amount = row[12]  # V列，是否未回款
                try:
                    receipt_date = parse_datetime(receipt_date_str, row_idx=row_idx)
                    if receipt_date_start <= receipt_date <= receipt_date_end and receipt_sales == '分销' and receipt_is_amount == '未回款'and '非订阅' not in receipt_booking:
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_column -= float(cleaned_value)
                                else:
                                    sum_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
        # 在D4单元格填充求和值
        sum_column = sum_column / ratio / 10000.0
        worksheet_24[unit] = sum_column
    return workbook

def processXlsx_XinQianChongCi_DingYue(workbook, unit, dateStart, dateEnd,ratio,product_name):
    # 解析整数形式的日期为日期对象
    receipt_date_start = datetime.datetime.strptime(str(dateStart), "%Y%m%d")
    receipt_date_end = datetime.datetime.strptime(str(dateEnd), "%Y%m%d")

    if "24年业绩预测-机构" in workbook.sheetnames:
        worksheet_24 = workbook["24年业绩预测-机构"]
        # 获取项目漏斗汇总-签约金额替重中Z列的求和
        sum_column = 0
        if "项目漏斗汇总-签约金额替重" in workbook.sheetnames:
            worksheet_receipt = workbook["项目漏斗汇总-签约金额替重"]
            for row_idx, row in enumerate(worksheet_receipt.iter_rows(min_row=2, min_col=23, max_col=26, values_only=True), start=2):
                # 检查是否为空行
                if not any(row):
                    continue
                receipt_booking = row[0]  # W列，订阅与否
                receipt_date_str = row[1]  # X列，日期
                receipt_amount = row[3]  # Z列，回款

                try:
                    receipt_date = parse_datetime(receipt_date_str, row_idx=row_idx)
                    if receipt_date_start <= receipt_date <= receipt_date_end and '订阅' in receipt_booking:
                        if isinstance(receipt_amount, str):  # 如果值是字符串类型
                            # 删除逗号并尝试将字符串转换为浮点数
                            try:
                                cleaned_value = receipt_amount.replace(',', '')
                                # 如果字符串包含负号，我们在转换之前将其删除
                                if '-' in cleaned_value:
                                    cleaned_value = cleaned_value.replace('-', '')
                                    sum_column -= float(cleaned_value)
                                else:
                                    sum_column += float(cleaned_value)
                            except ValueError as e:
                                # 如果无法转换为浮点数，则打印错误消息并继续下一个值
                                print(f"数值错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
                        elif isinstance(receipt_amount, (int, float)):  # 如果值是数值类型
                            sum_column += receipt_amount
                except ValueError as e:
                    print(f"日期解析错误：{e}，位于工作表 '应收及分销预测汇总' 的第 {row_idx} 行")
        # 在D4单元格填充求和值
        sum_column = sum_column / ratio / 10000.0
        worksheet_24[unit] = sum_column
    return workbook


