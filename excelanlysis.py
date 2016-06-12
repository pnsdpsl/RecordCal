import xlrd
import time

TODAYTIMEFORMAT = "%Y%m%d"

local_timer = time.strftime(ISOTIMEFORMAT, time.localtime())
today_file = time.strftime(TODAYTIMEFORMAT, time.localtime())+".xls"

# 读取当天的代还款数据的Excel文档
print("Start Reading Today's XLS file...")
organization_name = ["公司总部", "公司总部(特商)"]

