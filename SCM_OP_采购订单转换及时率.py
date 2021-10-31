import pandas as pd
import calendar
import datetime
from datetime import timedelta
from numpy import datetime64
from openpyxl import load_workbook

all_merge_data_mrp = []  # 筛选合并的mrp数据
all_data_mrp = []  # 所有的mrp数据
pd.set_option('display.max_columns', None)
all_list = []

def ReformDays(Days):
    now_work_days = []
    for day in Days:
        if day < 10:
            now_work_days.append("0" + str(day))
        else:
            now_work_days.append(str(day))
    return now_work_days


def CheckDataMRP(first, two):
    global all_merge_data_mrp
    first = first.dropna(subset=['物料编码'])  # 去除nan的列
    two = two.dropna(subset=['物料编码'])  # 去除nan的列
    out_data = pd.merge(first, two.drop(labels=['抓取时间'], axis=1), on=['物料编码', '物料名称', '需求跟踪号', '需求跟踪行号', '物料属性'])
    all_merge_data_mrp.append(out_data)


# 获取当月工作日函数
def WorkDays(year, month):
    # 利用日历函数，创建截取工作日日期
    cal = calendar.Calendar()
    Work_Days = []  # 创建工作日数组
    for week in cal.monthdayscalendar(int(year), int(month)):
        for i, day in enumerate(week):
            # 为0或者大于等于5的为休息日
            if day == 0 or i >= 5:
                continue
            # 否则加入数组
            Work_Days.append(day)
    return Work_Days


if __name__ == '__main__':
    now = datetime.date.today()
    # 获取当月首尾日期
    this_month_start = datetime.datetime(now.year, now.month, 1)
    this_month_end = datetime.datetime(now.year, now.month, calendar.monthrange(now.year, now.month)[1])
    # 获取上月首尾日期
    last_month_end = this_month_start - timedelta(days=1)
    last_month_start = datetime.datetime(last_month_end.year, last_month_end.month, 1)
    # 获取截取这个月份、年、上个月
    this_month_start = str(this_month_start).split(" ")[0]
    this_month_check = this_month_start
    this_month_end = str(this_month_end).split(" ")[0]
    this_month = this_month_start.split("-")[1]

    year = this_month_start.split("-")[0]

    last_month_start = str(last_month_start).split(" ")[0]
    last_month = last_month_start.split("-")[1]

    last_work_days = WorkDays(year, last_month)  # 获取上个月工作日
    this_work_days = WorkDays(year, this_month)  # 获取这个月工作日
    work_days7 = []  # 设设置到上月7天
    work_days = []  # 设置到上月3天
    work_days.extend(last_work_days[-3:])
    work_days.extend(this_work_days)  # 将上个月最后三天和这个月工作日相合并

    work_days7.extend(last_work_days[-7:])
    work_days7.extend(this_work_days)  # 将上个月最后七天和这个月工作日相合并
    work_days = ReformDays(work_days)  # 改造
    work_days7 = ReformDays(work_days7)  # 改造

    base_data_mrp = []

    flag = 0
    for x in work_days:
        if flag < 3:
            try:
                base_data = pd.read_excel(f"./DATA/SCM/OP/MRP计划维护--全部{year}-{last_month}-{x}.XLSX",
                                          usecols=['物料编码', '物料名称', '需求跟踪号', '需求跟踪行号', '物料属性'],
                                          converters={'物料编码': int, '需求跟踪行号': int}
                                          )
                base_data = base_data.loc[base_data["物料属性"] == "采购"]
                catch_data = datetime64(f"{year}-{last_month}-{x}")
                base_data['抓取时间'] = catch_data
                base_data_mrp.append(base_data)
                flag = flag + 1
                continue
            except:
                flag = flag + 1
                continue
        else:
            try:
                now_data = pd.read_excel(f"./DATA/SCM/OP/MRP计划维护--全部{year}-{this_month}-{x}.XLSX",
                                         usecols=['物料编码', '物料名称', '需求跟踪号', '需求跟踪行号', '物料属性'],
                                         converters={'物料编码': int, '需求跟踪行号': int}
                                         )
                all_data_mrp.append(now_data)
            except:
                continue
            catch_data = datetime64(f"{year}-{this_month}-{x}")
            now_data['抓取时间'] = catch_data
            now_data = now_data.loc[now_data["物料属性"] == "采购"]
            base_data_mrp.append(now_data)  # 新添加新的base
            CheckDataMRP(base_data_mrp[0], now_data)  # 合并检查是否存在一样的
            del (base_data_mrp[0])  # 删除第一个base
    all_mrp = pd.concat(all_merge_data_mrp, axis=0, ignore_index=True)
    for name1, group in all_mrp.groupby(['物料编码', '需求跟踪号', '需求跟踪行号']):
        group = pd.DataFrame(group)  # 新建pandas
        group = group.sort_values(by='抓取时间', ascending=True)  # 升序排序
        max_data_list = group.head(1)  # 取最早的抓取时间，就是排序后的第一列
        all_list.append(max_data_list)  # 加入list
        flag = flag + 1
    complete_merge_data = pd.concat(all_list, axis=0, ignore_index=True)

    end_merge_data = pd.concat(all_merge_data_mrp, axis=0, ignore_index=True)
    end_merge_data = end_merge_data.drop_duplicates()

# 未转换请购订单清单
    this_month_start = str(this_month_start).split(" ")[0].replace("-", "")
    last_month_start = str(last_month_start).split(" ")[0].replace("-", "")
    this_month_end = str(this_month_end).split(" ")[0].replace("-", "")

    po_data = pd.read_excel(f"./DATA/SCM/OP/采购订单列表-{last_month_start}-{this_month_end}.XLSX",
                            usecols=['请购单号', '存货编码', '存货名称', '行号'],
                            converters={'请购单号': str, '订单编号': int, '行号': int, '存货编码': int}
                            )
    pr_data = pd.read_excel(f"./DATA/SCM/OP/请购单列表-{this_month_start}-{this_month_end}.XLSX",
                            usecols=['单据号', '存货编码', '存货名称', '行号'],
                            converters={'单据号': str, '行号': int, '存货编码': int}
                            )

    pr_data = pr_data.rename(columns={'单据号': '请购单号'})  # 修改单据号为请购单号
    pr_data = pr_data.dropna(subset=['存货编码'])  # 去除nan的列
    po_data = po_data.dropna(subset=['存货编码'])  # 去除nan的列

    pr_data = pr_data.append(po_data)
    pr_data = pr_data.append(po_data)

    pr_data.drop_duplicates(keep=False, inplace=True)
    pr_data.reset_index()
    #  小于当月的历史未转换MRP数据筛选
    MRPHistory_data = complete_merge_data[complete_merge_data['抓取时间'] < datetime64(this_month_check)]
    #  当月大于2天的未转换MRP数据筛选
    MRPNow_data = complete_merge_data[complete_merge_data['抓取时间'] >= datetime64(this_month_check)]

    book = load_workbook('./RESULT/SCM/OP/采购订单转换及时率.xlsx')
    writer = pd.ExcelWriter("./RESULT/SCM/OP/采购订单转换及时率.xlsx", engine='openpyxl')
    writer.book = book
    MRPNow_data.to_excel(writer, "当月未转换MRP清单", index=False)
    writer.save()

    pr_data.to_excel('./RESULT/SCM/OP/采购订单转换及时率.xlsx', sheet_name="当月未转换PR清单", index=False)
    book = load_workbook('./RESULT/SCM/OP/采购订单转换及时率.xlsx')
    writer = pd.ExcelWriter("./RESULT/SCM/OP/采购订单转换及时率.xlsx", engine='openpyxl')
    writer.book = book
    MRPHistory_data.to_excel(writer, "历史未转换MRP清单", index=False)
    writer.save()



