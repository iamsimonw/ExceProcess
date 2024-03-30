import os
import openpyxl
import tools




if __name__ == "__main__":
    # 固定的CSV文件目录
    csv_directory = r"D:\JB\表(1)\原始表"
    excel_file = "浙南temp.xlsx"
    # 三个CSV文件的文件名
    csv_filename0 ="应收及分销预测汇总.csv"
    csv_filename1 ="项目漏斗汇总-签约金额替重.csv"
    csv_filename2 ="收款明细表.csv"
    # 目标Excel文件的路径

    
    tools.copyCSVtoXlsx(csv_directory,csv_filename0,excel_file)
    tools.copyCSVtoXlsx(csv_directory,csv_filename1,excel_file)
    tools.copyCSVtoXlsx(csv_directory,csv_filename2,excel_file)
