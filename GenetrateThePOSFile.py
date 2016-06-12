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

resultdict = {}
inputdata = {}
fangkuandata = {}

"""
生成Excel文档
"""
def generateExcelFile():
    print("Starting Generating Excel File")
    grade_file = 'tongji/POS贷业绩统计.xlsx'
    if not os.path.isfile(grade_file):
        print("File not Exist")
        sys.exit()
    data = xlrd.open_workbook(grade_file)
    grade_table = data.sheet_by_index(0)
    nrows = grade_table.nrows
    ncols = grade_table.ncols
    new_book = "tongji/" + today_file_grade_file
    workbook = xlsxwriter.Workbook(new_book)
    # 增加sheet1和sheet2
    worksheet = workbook.add_worksheet()
    worksheet2 = workbook.add_worksheet()
    # 拷贝现有的数据至sheet2
    for i in range(nrows):
        for j in range(ncols):
            cell_value = grade_table.cell_value(i, j)
            worksheet2.write(i, j, cell_value)
    # todo 生成表头
    # 设置格式
    worksheet.set_column('A:A', 18)
    title_format = workbook.add_format({'align': 'center',
                                   'valign': 'vcenter',
                                    'font_size': '14',
                                        'bold': '1',
                                        'bg_color': 'gray'})
    worksheet.merge_range('A1:M2', "", title_format)
    worksheet.write_rich_string('A1', 'POS贷业绩日报（'+today+')', title_format)

    # TODO 组合生成综合字典

    # TODO 数据写入Excel文档

    # TODO 写入公式
    # 保存生成的文件
    workbook.close()

"""
读取当天的进件数据，并且填写进件dict中去
"""
def gettheinputdata():
    inputfile = 'jinjian/'+today_file
    try:
        print("Starting reading the jinjian xls file")
        jinjian_data = xlrd.open_workbook(inputfile)
        jinjian_table = jinjian_data.sheets()[0]
        jinjian_rows = jinjian_table.nrows
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
                if bank_name not in inputdata:
                    inputdata[bank_name] = 1
                else:
                    inputdata[bank_name] += 1
        print(inputdata)
    except IOError:
        print("Reading jinjian file failed")


"""
获取放款数据，并填写到放款dict中去
"""
def getthefangkuandata():
    print("Start Getting the Fang Kuan Data")
    try:
        print("Start Reading Today's loan XLS file ...")
        loan_data = xlrd.open_workbook('fangkuan/'+today_file)
        loan_table = loan_data.sheets()[0]
        loan_rows = loan_table.nrows
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
                if bank_name not in fangkuandata:
                    fangkuandata[bank_name] = [1, loan_money]
                else:
                    fangkuandata[bank_name][0] += 1
                    fangkuandata[bank_name][1] += loan_money
        print(fangkuandata)
    except IOError:
        print("Reading loan xls file Failed")


"""
生成一个整体的dict，其中包含了进件数据和放款数据
"""
def generatecombineddict():
    print("Starting Generating the Combined Dictionary")


gettheinputdata()
getthefangkuandata()
generatecombineddict()
generateExcelFile()