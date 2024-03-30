import pandas as pd
import os

def copy_csv_to_excel(csv_filenames, excel_file, csv_directory):
    # 确保目标目录存在
    if not os.path.exists(csv_directory):
        os.makedirs(csv_directory)

    # 创建 Excel 写入器对象
    excel_file_path = os.path.join(csv_directory, excel_file)

    # 遍历每个 CSV 文件
    for csv_filename in csv_filenames:
        csv_file = os.path.join(csv_directory, csv_filename)
        # 从文件名中提取工作表名称
        sheet_name = os.path.splitext(csv_filename)[0]

        # 读取 CSV 文件到 DataFrame，跳过第一行
        df = pd.read_csv(csv_file, skiprows=1)
        
        # 如果 Excel 文件已存在，则尝试读取同名工作表的内容
        if os.path.exists(excel_file_path):
            with pd.ExcelFile(excel_file_path) as xls:
                if sheet_name in xls.sheet_names:
                    # 读取同名工作表的内容到 DataFrame
                    existing_df = pd.read_excel(xls, sheet_name)
                    # 将 CSV 文件的内容添加到同名工作表的下方
                    new_df = pd.concat([existing_df, df], ignore_index=True)
                    # 写入 Excel 文件
                    with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:
                        new_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    continue  # 继续下一个 CSV 文件

        # 如果不存在同名工作表，则直接将 DataFrame 写入到 Excel 文件的对应工作表中
        # with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
        #     df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)

    print("CSV 文件已成功复制到 Excel 文件的对应工作表中")

def main():
    # 固定的 CSV 文件目录
    csv_directory = r"D:\JB\表(1)\原始表"
    # 三个 CSV 文件的文件名
    csv_filenames = ["收款明细表.csv", "项目漏斗汇总-签约金额替重.csv", "应收及分销预测汇总.csv"]
    # 目标 Excel 文件的路径
    excel_file = "浙南temp.xls"

    # 复制 CSV 文件到 Excel 文件的对应工作表中
    copy_csv_to_excel(csv_filenames, excel_file, csv_directory)

if __name__ == "__main__":
    main()
