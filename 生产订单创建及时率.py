import xlrd
import pandas as pd
import calendar
import datetime
from datetime import timedelta
import openpyxl

from numpy import datetime64
from openpyxl import load_workbook
data_time = datetime64("2000-01-02")  # 设置工艺路线版本日期的最早期限
now = datetime.date.today()

# 获取当月首尾日期
this_month_start = datetime.datetime(now.year, now.month, 1)
this_month_end = datetime.datetime(now.year, now.month, calendar.monthrange(now.year, now.month)[1])

# 获取上月首尾日期
last_month_end = this_month_start - timedelta(days=1)
last_month_start = datetime.datetime(last_month_end.year, last_month_end.month, 1)

# 将上月首尾日期切割
_last_month_end = str(last_month_end).split(" ")[0].replace("", "")
last_month_start = str(last_month_start).split(" ")[0].replace("-", "")
last_month_end = str(last_month_end).split(" ")[0].replace("-", "")
# 材料出库及时率
Production_data = pd.read_excel(f"./KPI/PROD/生产订单列表-{last_month_start}-{last_month_end}.XLSX",
                                usecols=['生产订单号', '行号', '物料编码', '物料名称', '生产批号', '制单时间', '类型'],
                                converters={'生产订单号': str, '制单时间': datetime64, '物料编码': str})
Production_data = Production_data[Production_data["类型"] == "标准"]
# '母件编码', '母件名称', '旧子件编码',
# '旧子件名称', '旧子件基本用量', '新子件编码', '新子件名称', '新子件基本用量',

Material_data = pd.read_excel(f"./KPI/SCM/OM/存货档案{_last_month_end}.XLSX",
                              usecols=['存货编码', '计划默认属性', '启用日期'],
                              converters={'启用日期': datetime64, "存货编码": str})
Production_this_month_data = Production_data[Production_data["制单时间"] > this_month_start]

Material_data = Material_data.rename(columns={'存货编码': '物料编码'})

this_month_merge_data = pd.merge(Production_this_month_data, Material_data, how="left", on=['物料编码'])
this_month_merge_data.to_excel('./KPI/SCM/OM/demo.xlsx', sheet_name="demo", index=False)

Routing_data = pd.read_excel(f"./KPI/SCM/OM/工艺路线资料表--含资源.xlsx",
                             usecols=[0, 4, 6], header=3, names=["物料编码", "版本代号", "版本日期"],
                             converters={'版本日期': str, '物料编码': str})
Routing_data = Routing_data.dropna(subset=['版本代号'])  # 去除nan的列
Routing_data["版本日期"] = Routing_data["版本日期"].str.replace("/", "-").astype("datetime64")
Routing_data = Routing_data[Routing_data["版本日期"] > data_time]
Routing_data = Routing_data.drop_duplicates(subset=["物料编码"])  # 去重
# Routing_data.to_excel('./KPI/SCM/OM/demo.xlsx', index=False)
old_Material_data = this_month_merge_data[this_month_merge_data["启用日期"].isnull()]  # 旧物料
del old_Material_data['计划默认属性']
del old_Material_data['启用日期']
new_Material_data = this_month_merge_data[this_month_merge_data["启用日期"].notnull()]  # 新物料

new_Production_data = pd.merge(new_Material_data, Routing_data, on=["物料编码"], how="left")
new_Production_data = new_Production_data.dropna(subset=["版本日期"])  # 去nan
new_Production_data['下单延时/H'] = ((new_Production_data['制单时间'] - new_Production_data['版本日期']) / pd.Timedelta(1, 'H')).astype(
    int)
new_Production_data.loc[new_Production_data["下单延时/H"] > 72, "创建及时率"] = "超时"  # 计算出来的审批延时大于3天为超时
new_Production_data.loc[new_Production_data["下单延时/H"] <= 72, "创建及时率"] = "正常"  # 小于等于3天为正常
BOM_data = pd.read_excel(f"./KPI/SCM/OM/BOM集成时间表.xlsx", header=1,
                         usecols=['子件编码', '计划默认属性', '集成时间'],
                         converters={"子件编码": str})
BOM_data = BOM_data.rename(columns={'子件编码': '物料编码'})
BOM_data = BOM_data.dropna(subset=["集成时间"])  # 去nan
BOM_data["集成时间"] = BOM_data["集成时间"].astype("datetime64")

old_Production_data = pd.merge(old_Material_data, BOM_data, on=["物料编码"], how="left")
old_Production_data = old_Production_data.dropna(subset=["集成时间"])  # 去nan
old_Production_data['下单延时/H'] = (
            (old_Production_data['制单时间'] - old_Production_data['集成时间']) / pd.Timedelta(1, 'H')).astype(
    int)
old_Production_data = old_Production_data.drop_duplicates(subset=["生产订单号", "行号", "物料编码"])  # 去重
old_Production_data.loc[old_Production_data["下单延时/H"] > 24, "创建及时率"] = "超时"  # 计算出来的审批延时大于1天为超时
old_Production_data.loc[old_Production_data["下单延时/H"] <= 24, "创建及时率"] = "正常"  # 小于等于1天为正常
# 输出新旧物料及时率
new_Production_data.to_excel('./KPI/SCM/OM/生产订单创建及时率.xlsx', sheet_name="新物料", index=False)
book = load_workbook('./KPI/SCM/OM/生产订单创建及时率.xlsx')
writer = pd.ExcelWriter("./KPI/SCM/OM/生产订单创建及时率.xlsx", engine='openpyxl')
writer.book = book
old_Production_data.to_excel(writer, "旧物料", index=False)
writer.save()
