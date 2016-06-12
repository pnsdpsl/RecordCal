"""
@author silva
本程序主要去统计POS贷的数据，根据地区和进件客户经理进行分类
"""
# coding: utf-8
import xlrd
import time
import xlsxwriter
import sys
import os.path


TODAYTIMEFORMAT = "%Y%m%d"
ISOTIMEFORMAT="%Y-%m-%d %X"

local_timer = time.strftime(ISOTIMEFORMAT, time.localtime())
today_file = time.strftime(TODAYTIMEFORMAT, time.localtime())+".xls"
today = time.strftime(TODAYTIMEFORMAT, time.localtime())
today_file_grade_file = time.strftime(TODAYTIMEFORMAT, time.localtime())+".xlsx"

organization_name = ['公司总部(特商)']
loan_row_names = ['序号', '事业部', '客户姓名', '贷款渠道', '证件号码', '申请来源', '申请日期', '发放日期', '还款日期', '到期日期', '还款方式', '合同金额', '合同号', '团队长名称', '主办客户经理名称', '单位电话', '单位电话区号', '放款金额', '当前余额', '贷款品种', '贷款专案', '执行利率', '期限（月数）', '省/直辖市', '市', '区', '详细地址', '省/直辖市', '市', '区', '详细地址', '省/直辖市', '市', '区', '详细地址', '进件来源（可包含pad、委托贷款、PC）', '进件通路    ', '贷款用途', '性别', '婚姻状况', '教育程度', '年龄', '单位名称', '部门', '单位性质', '行业', '职业', '职务', '逾期期数', '逾期天数']


def generateexcelfile(inputdict, fangkuandict):
    grade_file = 'D:\\02-特商业务部\\18-生产\\01-业绩报表\\特商POS+S白名单业绩统计模板.xlsx'
    if not os.path.isfile(grade_file):
        print("File not Exist")
        sys.exit()
    data = xlrd.open_workbook(grade_file)
    grade_table = data.sheet_by_index(0)
    nrows = grade_table.nrows
    ncols = grade_table.ncols

    new_book = "D:\\02-特商业务部\\18-生产\\01-业绩报表\\" + today_file_grade_file
    workbook = xlsxwriter.Workbook(new_book)
    # 增加sheet1和sheet2
    worksheet = workbook.add_worksheet()
    worksheet2 = workbook.add_worksheet()

    cell_format = workbook.add_format({'align': 'center', 'font_size': '14'})
    for i in range(nrows-1):
        for j in range(ncols):
            cell_value = grade_table.cell_value(i, j)
            worksheet2.write(i, j, cell_value)
            # 在sheet1中写入表头文件
            if i >= 0 and i <= 4 and j >=0 and j <= 12:
                worksheet.write(i, j, cell_value, cell_format)
            # to do 表头的格式设置

            if j == 0:
                worksheet.write(i, j, cell_value, cell_format)

            # 除去表头和最后一行，需要在表格中增加公式
            if i >= 5 and i < nrows-1:
                if j in [2, 3, 4]:
                    # chr(97)=a
                    myformula = '=b' + str(i+1) + '+sheet2!' + chr(j+97) + str(i+1)
                elif j in [7]:
                    myformula = '=f' + str(i+1) + '+sheet2!h' + str(i+1)
                elif j in [8]:
                    myformula = '=g' + str(i+1) + '+sheet2!i' + str(i+1)
                elif j in [9]:
                    myformula = '=f' + str(i+1) + '+sheet2!j' + str(i+1)
                elif j in [10]:
                    myformula = '=g' + str(i+1) + '+sheet2!k' + str(i+1)
                elif j in [11]:
                    myformula = '=f' + str(i+1) + '+sheet2!l' + str(i+1)
                elif j in [12]:
                    myformula = '=g' + str(i+1) + '+sheet2!m' + str(i+1)
                else:
                    myformula = ''
                worksheet.write(i, j, myformula)

    # 填写今天的进件数据
    fenhang_bank = grade_table.col_values(0)
    temprow = nrows-1
    for k in inputdict:
        if k in fenhang_bank:
            row_number = fenhang_bank.index(k)
            worksheet.write(row_number, 1, inputdict[k], cell_format)
        else:
            print("%s 不在进件列表中，请增加"%(k))
            worksheet.write(temprow, 0, k, cell_format)
            worksheet.write(temprow, 1, inputdict[k], cell_format)
            temprow += 1
    # TODO
    # 处理每周开始，每月开始，每年开始

    # 填写今天的放款笔数和金额
    for k in fangkuandict:
        if k in fenhang_bank:
            row_number = fenhang_bank.index(k)
            worksheet.write(row_number, 5, fangkuandict[k][0], cell_format)
            fangkuanedu = fangkuandict[k][1]/10000
            worksheet.write(row_number, 6, fangkuanedu, cell_format)
        else:
            print("%s 不在放款列表中，请增加"%(k))
            worksheet.write(temprow, 0, k, cell_format)
            worksheet.write(temprow, 5, fangkuandict[k][0], cell_format)
            fangkuanedu = fangkuandict[k][1]/10000
            worksheet.write(temprow, 6, fangkuanedu, cell_format)
            temprow += 1
    # TODO
    # 处理每周开始，每月开始，每年开始

    # 最后一行增加sum公式
    for j in range(ncols):
        if j == 0:
            worksheet.write(temprow, j, "总计")
        else:
            myformula = '=sum('+ chr(j+97)+'6:'+chr(j+97)+str(temprow-1) +')'
            worksheet.write(temprow, j, myformula)
    print(fenhang_bank)
    # 设置格式
    worksheet.set_column('A:A', 18)
    title_format = workbook.add_format({'align': 'center',
                                   'valign': 'vcenter',
                                    'font_size':'14',
                                        'bold':'1'})
    worksheet.merge_range('A1:M2', "", title_format)
    worksheet.write_rich_string('A1', 'POS贷业绩日报（'+today+')', title_format)
    # 保存生成的文件
    workbook.close()

# 读取当天的放款数据的Excel文档
try:
    print("Start Reading Today's loan XLS file ...")
    loan_data = xlrd.open_workbook('D:\\02-特商业务部\\18-生产\\03-生产数据\\01放款数据\\'+today_file)
    loan_table = loan_data.sheets()[0]
    loan_rows = loan_table.nrows
    result_dict = {}
    for i in range(loan_rows):
        if i in range(3):
            continue
        if loan_table.row_values(i)[loan_row_names.index('事业部')] in organization_name:
            cust_manager_name = loan_table.row_values(i)[loan_row_names.index('主办客户经理名称')]
            loan_money = loan_table.row_values(i)[loan_row_names.index('合同金额')]
            # 如果客户经理的明细是代码的，则截取，如果不是代码，保持原样
            if "1" in cust_manager_name or "2" in cust_manager_name or "0" in cust_manager_name:
                bank_name = cust_manager_name[:-2]
            else:
                bank_name = cust_manager_name
            if bank_name not in result_dict:
                result_dict[bank_name] = [1, loan_money]
            else:
                result_dict[bank_name][0] += 1
                result_dict[bank_name][1] += loan_money
    print(result_dict)
except IOError:
    print("Reading loan xls file Failed")


# 读取当天的进件报表数据
jinjian_file = 'D:\\02-特商业务部\\18-生产\\03-生产数据\\02进件数据\\'+today_file
try:
    print("Starting reading the jinjian xls file")
    jinjian_data = xlrd.open_workbook(jinjian_file)
    jinjian_table = jinjian_data.sheets()[0]
    jinjian_rows = jinjian_table.nrows
    jinjian_dict = {}
    for i in range(jinjian_rows):
        if i in range(3):
            continue
        if jinjian_table.row_values(i)[1] in organization_name:
            cust_manager_name = jinjian_table.row_values(i)[4]
            # 如果客户经理的明细是代码的，则截取，如果不是代码，保持原样
            if "1" in cust_manager_name or "2" in cust_manager_name or "0" in cust_manager_name:
                bank_name = cust_manager_name[:-2]
            else:
                bank_name = cust_manager_name
            if bank_name not in jinjian_dict:
                jinjian_dict[bank_name] = 1
            else:
                jinjian_dict[bank_name] += 1
    print(jinjian_dict)
except IOError:
    print("Reading jinjian file failed")

generateexcelfile(jinjian_dict, result_dict)
"""
读取业绩报告，首先拷贝现有的数据至原有备份sheet
然后将本日的数据放入到统计的sheet中，进行计算
"""