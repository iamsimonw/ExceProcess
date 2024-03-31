import os
import openpyxl
import tools



def All_Process(workbook,cell_number,time_start,time_end,ratio_all,ratio_FenXiao,ratio_ZhiXiao,ratio_DingYue,product_name='Full'):
    if product_name=='Full':
        '''
        整体
        '''
        #税后回款
        workbook = tools.processXlsx_ShuiHouHuiKuan(workbook,f"D{cell_number}",time_start,time_end,ratio_all,product_name)
        #直销应收确保
        workbook = tools.processXlsx_ZhiXiaoYingShouQueBao(workbook,f"F{cell_number}",time_start,time_end,ratio_all,product_name)
        #分销应收确保
        workbook = tools.processXlsx_FenXiaoYingShouQueBao(workbook,f"G{cell_number}",time_start,time_end,ratio_all,product_name)
        #新签确保
        workbook = tools.processXlsx_XinQianQueBao(workbook,f"H{cell_number}",time_start,time_end,ratio_all,product_name)
        #直销应收冲刺
        workbook = tools.processXlsx_ZhiXiaoYingShouChongCi(workbook,f"J{cell_number}",time_start,time_end,ratio_all,product_name)
        #分销应收冲刺
        workbook = tools.processXlsx_FenXiaoYingShouChongCi(workbook,f"K{cell_number}",time_start,time_end,ratio_all,product_name)
        #新签冲刺
        workbook = tools.processXlsx_XinQianChongCi(workbook,f"L{cell_number}",time_start,time_end,ratio_all,product_name)
        '''
        直销
        '''
        workbook = tools.processXlsx_ShuiHouHuiKuan_ZhiXiao(workbook,f"U{cell_number}",time_start,time_end,ratio_ZhiXiao,product_name)
        '''
        分销
        '''
        workbook = tools.processXlsx_ShuiHouHuiKuan_FenXiao(workbook,f"AF{cell_number}",time_start,time_end,ratio_FenXiao,product_name)
        '''
        订阅
        '''
        workbook = tools.processXlsx_ShuiHouHuiKuan_DingYue(workbook,f"AO{cell_number}",time_start,time_end,ratio_DingYue,product_name)
        workbook = tools.processXlsx_ZhiXiaoYingShouQueBao_DingYue(workbook,f"AQ{cell_number}",time_start,time_end,ratio_DingYue,product_name)
        workbook = tools.processXlsx_FenXiaoYingShouQueBao_DingYue(workbook,f"AR{cell_number}",time_start,time_end,ratio_DingYue,product_name)
        workbook = tools.processXlsx_XinQianQueBao_DingYue(workbook,f"AS{cell_number}",time_start,time_end,ratio_DingYue,product_name)
        workbook = tools.processXlsx_ZhiXiaoYingShouChongCi_DingYue(workbook,f"AU{cell_number}",time_start,time_end,ratio_DingYue,product_name)
        workbook = tools.processXlsx_FenXiaoYingShouChongCi_DingYue(workbook,f"AV{cell_number}",time_start,time_end,ratio_DingYue,product_name)
        workbook = tools.processXlsx_XinQianChongCi_DingYue(workbook,f"AW{cell_number}",time_start,time_end,ratio_DingYue,product_name)
    elif product_name=='YS':
        1
    
    
    return workbook