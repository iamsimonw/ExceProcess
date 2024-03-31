import os
import openpyxl
import tools




if __name__ == "__main__":
    # 固定的CSV文件目录
    csv_directory = r"D:\JB\表(1)\原始表"
    excel_file = "浙南temp.xlsx"
    excel_file_full = os.path.join(csv_directory, excel_file)
    # 三个CSV文件的文件名
    csv_filename0 ="应收及分销预测汇总.csv"
    csv_filename1 ="项目漏斗汇总-签约金额替重.csv"
    csv_filename2 ="收款明细表.csv"
    quarter_start = "20240401"
    quarter_end = "20240701"
    year_start="20240101"
    year_end="20241231"
    # tools.copyCSVtoXlsx(csv_directory,csv_filename0,excel_file)
    # tools.copyCSVtoXlsx(csv_directory,csv_filename1,excel_file)
    # tools.copyCSVtoXlsx(csv_directory,csv_filename2,excel_file)
    ratio_all=1.1
    ratio_ZhiXiao=1.11
    ratio_FenXiao=1.13
    
    # 创建或加载Excel工作簿
    if os.path.exists(excel_file_full):
        workbook = openpyxl.load_workbook(excel_file_full)
    else:
        workbook = openpyxl.Workbook()
    '''
    当季整体
    '''
    #税后回款
    workbook = tools.processXlsx_ShuiHouHuiKuan(workbook,"D11",quarter_start,quarter_end,ratio_all)
    #直销应收确保
    workbook = tools.processXlsx_ZhiXiaoYingShouQueBao(workbook,"F11",quarter_start,quarter_end,ratio_all)
    #分销应收确保
    workbook = tools.processXlsx_FenXiaoYingShouQueBao(workbook,"G11",quarter_start,quarter_end,ratio_all)
    #新签确保
    workbook = tools.processXlsx_XinQianQueBao(workbook,"H11",quarter_start,quarter_end,ratio_all)
    #直销应收冲刺
    workbook = tools.processXlsx_ZhiXiaoYingShouChongCi(workbook,"J11",quarter_start,quarter_end,ratio_all)
    #分销应收冲刺
    workbook = tools.processXlsx_FenXiaoYingShouChongCi(workbook,"K11",quarter_start,quarter_end,ratio_all)
    #新签冲刺
    workbook = tools.processXlsx_XinQianChongCi(workbook,"L11",quarter_start,quarter_end,ratio_all)
    '''
    当季直销
    '''
    workbook = tools.processXlsx_ShuiHouHuiKuan_ZhiXiao(workbook,"U11",quarter_start,quarter_end,ratio_all)
    
    
    
    workbook.save(excel_file_full)