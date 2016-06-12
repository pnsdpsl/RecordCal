# coding=utf-8
import xlrd
import time
import sys

ISOTIMEFORMAT="%Y-%m-%d %X"
TODAYTIMEFORMAT = "%Y%m%d"
THISMONTH = "%Y%m"
WEEK = '%w'
local_timer = time.strftime(ISOTIMEFORMAT, time.localtime())
today_file = time.strftime(TODAYTIMEFORMAT, time.localtime())+".xls"
dayofWeek = time.strftime(WEEK, time.localtime())


def getCityRate():
    print("get the xls file")
    this_file = xlrd.open_workbook('D:\\02-特商业务部\\18-生产\\03-生产数据\\04逾期客户明细\\'+today_file)
    try:
        table = this_file.sheets()[0]
        table_nrows = table.nrows
        for i in range(table_nrows):
            cust_address = table.row_values(i)[23]
            print(cust_address)
    except IOError:
        print("读取今日逾期数据文件失败")


'''
如果是星期一，则需要读取三天前的逾期客户明细报表，否则就读取一天前的逾期客户明细报表
如果是星期六，则需要读取一天前，
如果是星期日，则需要读取两天前
'''
if int(dayofWeek) == 1:
    yesterday_file = time.strftime(TODAYTIMEFORMAT, time.localtime(time.time() - 24*60*60*3)) + ".xls"
elif int(dayofWeek) == 0:
    yesterday_file = time.strftime(TODAYTIMEFORMAT, time.localtime(time.time() - 24*60*60*2)) + ".xls"
else:
    yesterday_file = time.strftime(TODAYTIMEFORMAT, time.localtime(time.time() - 24*60*60)) + ".xls"
this_month_file = time.strftime(THISMONTH, time.localtime())+".xls"

# 读取当天的逾期数据的Excel文档
print("Start Reading Today's XLS file...")
organization_name = ["公司总部", "公司总部(特商)"]
try:
    data = xlrd.open_workbook('D:\\02-特商业务部\\18-生产\\03-生产数据\\04逾期客户明细\\'+today_file)
    table = data.sheets()[0]
    today_nrows = table.nrows
    print("Today All is %d rows" % (today_nrows))
    # 存入数组
    arr = []
    # 计算逾期合同余额，组成逾期客户的申请编号数组
    overdue_contract_number = 0
    for i in range(today_nrows):
        if i in range(3):
            continue
        if table.row_values(i)[1] in organization_name:
            arr.append(table.row_values(i)[3])
            overdue_contract_number = overdue_contract_number + table.row_values(i)[12]
except IOError:
    print("今日逾期文件读取失败")
    sys.exit()

# 读取前天的逾期数据的Excel文档
print("Start Reading Yesterday's XLS File...")
try:
    data_before = xlrd.open_workbook('D:\\02-特商业务部\\18-生产\\03-生产数据\\04逾期客户明细\\'+yesterday_file)
    table_before = data_before.sheets()[0]
    rows_before = table_before.nrows
    print("Yesterday All is %d Rows" %(rows_before))

    # 存入数组
    arr_before = [];
    for i in range(rows_before):
        if i in range(3):
            continue
        if table_before.row_values(i)[1] in organization_name:
            arr_before.append(table_before.row_values(i)[3])
except IOError:
    print("昨日逾期文件读取失败")
    sys.exit()

record = []
for i in arr_before:
    index_row_before = arr_before.index(i)
    # 修正indexrow，因为统计申请编号的时候已经删除最初文档前面的三行
    index_row_before = index_row_before + 3
    # 获取到昨天的逾期金额
    overdue_before = table_before.row_values(index_row_before)[13]
    cust_name = table_before.row_values(index_row_before)[4]
    product_type = table_before.row_values(index_row_before)[6]
    merchant = table_before.row_values(index_row_before)[32]

    # 获取到今天的逾期金额 并与昨天的逾期金额做比较
    if i in arr:
        index_row = arr.index(i)
        index_row = index_row + 3
        overdue = table.row_values(index_row)[13]
        if overdue_before != overdue:
            payback = round((overdue_before-overdue), 2)
            if payback > 0:
                if merchant  == '':
                    record.append("[未结清]产品为： "+product_type+" 的客户： "+cust_name+" 今日还款金额为："+str(payback)+" 当前逾期金额为： "+str(overdue))
                else:
                    record.append("[未结清]产品为： "+product_type +" 商户： "+ merchant +" 的客户： "+cust_name+" 今日还款金额为："+str(payback)+" 当前逾期金额为： "+str(overdue))
    else:
        if merchant == '':
            record.append("[已结清]产品为： "+product_type+" 的客户： "+cust_name + " 已结清")
        else:
            record.append("[已结清]产品为： "+product_type +" 商户： "+ merchant +" 的客户： "+cust_name+" 已结清")

# 寻找新逾期客户
for i in arr:
    if i not in arr_before:
        index_row = arr.index(i)
        index_row += 3
        overdue = table.row_values(index_row)[13]
        cust_name = table.row_values(index_row)[4]
        product_type = table.row_values(index_row)[6]
        merchant = table.row_values(index_row)[32]
        if merchant == '':
            record.append("[新增逾期]产品为： "+product_type+" 的客户： "+cust_name + " 逾期金额为："+ str(overdue))
        else:
            record.append("[新增逾期]产品为： "+product_type +" 商户： "+ merchant +" 的客户： "+cust_name+" 逾期金额为： " +str(overdue))

# 对记录进行排序
record.sort()

'''
计算逾期率，首先获取当月待还款客户合同余额，
然后再计算当日的逾期客户合同余额，相除，即可获得
'''
try:
    ready_payback = xlrd.open_workbook('D:\\02-特商业务部\\18-生产\\03-生产数据\\03待还款客户明细\\'+this_month_file)
    table_ready_payback = ready_payback.sheets()[0]
    rows_ready_payback = table_ready_payback.nrows
    all_need_payback = 0
    for i in range(rows_ready_payback):
        if i in range(3):
            continue
        if table_ready_payback.row_values(i)[4] in organization_name:
            all_need_payback = all_need_payback + table_ready_payback.row_values(i)[18]
    overdue_rate = overdue_contract_number / all_need_payback
    print("实时逾期率为%.2f%%"%(overdue_rate*100))
except IOError:
    print("本月代还款文件读取失败")
    sys.exit()

# 将最后的结果写入到文档中
print("Writing to the file")
newfile = open(r'D:\02-特商业务部\18-生产\03-生产数据\04逾期客户明细\recentfile.txt', 'a')
try:
    # 与上一条记录保持三行的距离
    for i in range(3):
        newfile.write('\n')
    newfile.write(local_timer+'\n')
    for i in record:
        print(i)
        newfile.write(i+'\n')
    newfile.write("实时逾期率为 %.2f%%"%(overdue_rate*100))
finally:
    newfile.close()
