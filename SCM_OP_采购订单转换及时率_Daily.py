import pandas as pd
import calendar
import datetime
from datetime import timedelta, datetime
from numpy import datetime64
from openpyxl import load_workbook

import Func

pd.set_option('display.max_columns', None)


class OrderConversion:
    def __init__(self):
        self.YesMRPData = None
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = Func.Path()
        self.MRPScreenList = []  # 筛选合并的mrp数据
        self.MRPNewDataList = []  # 本月所有的mrp数据
        self.GroupMRPList = []  # 分组后的mrp数据
        self.today = datetime.today()  # datetime类型当前日期
        yesterday = self.today + timedelta(days=-3)  # 减去三天
        while 1:  # 获取前三天MRP数据，没有则继续往前推一天，直到有为止
            try:
                self.YesMRPData = pd.read_excel(f"{self.path}/DATA/SCM/OP/MRP计划维护--全部{str(yesterday)[:10]}.XLSX",
                                                usecols=['物料编码', '物料名称', '需求跟踪号', '需求跟踪行号', '物料属性', '是否客供料', '开工日期'],
                                                converters={'物料编码': int, '需求跟踪行号': int, '开工日期': str}
                                                )
                break
            except Exception as e:
                print(e)
                yesterday = yesterday + timedelta(days=-1)
        self.MRPData = pd.read_excel(f"{self.path}/DATA/SCM/OP/MRP计划维护--全部{str(self.today)[:10]}.XLSX",
                                     usecols=['物料编码', '物料名称', '需求跟踪号', '需求跟踪行号', '物料属性', '是否客供料', '开工日期'],
                                     converters={'物料编码': int, '需求跟踪行号': int, '开工日期': str}
                                     )
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0]
        self.LastMonthStart = str(self.LastMonthStart).split(" ")[0]
        self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0]
        self.this_month_check = self.ThisMonthStart

        self.POData = pd.read_excel(f"{self.path}/DATA/SCM/OP/采购订单列表.XLSX",
                                    usecols=['订单编号', '实际到货日期', '行号', '制单时间'],
                                    converters={'订单编号': str, '实际到货日期': datetime64, '制单时间': datetime64, '行号': str}
                                    )
        self.POData = self.POData.rename(columns={'行号': '采购订单行号', '订单编号': '采购订单号', '制单时间': '采购订单制单时间'})

        self.PRProcessData = pd.read_excel(f"{self.path}/DATA/SCM/OP/请购执行进度表.XLSX",
                                           usecols=[1, 6, 7, 8, 9, 10, 14, 15, 20, 23], header=4,
                                           names=["请购单号", "请购单行号", "存货编码", "存货名称", "规格型号", "数量", "采购订单号", "采购订单行号",
                                                  "计划到货日期", "采购员"],
                                           converters={'请购单号': str, '请购单行号': str, '存货编码': str, '数量': float,
                                                       '采购订单号': str, '采购订单行号': str, '计划到货日期': datetime64})

        self.PRData = pd.read_excel(f"{self.path}/DATA/SCM/OP/请购单列表.XLSX",
                                    usecols=['单据号', '行号', '行关闭人', '建议订货日期', '审核时间', '制单时间', '存货编码', '存货名称', '规格型号',
                                             '数量'],
                                    converters={'单据号': str, '行号': str, '审核时间': datetime64, '制单时间': datetime64,
                                                '建议订货日期': datetime64, '存货编码': str, '数量': float})
        self.PRData = self.PRData.rename(columns={'行号': '请购单行号', '单据号': '请购单号', '制单时间': '请购单制单时间', '审核时间': '请购单审核时间'})

        self.PRData = self.PRData.loc[self.PRData["行关闭人"].isnull()]  # 筛选 行关闭人 为空的行
        self.PRApproveNotTime = self.PRData.loc[self.PRData["请购单审核时间"].isnull()]  # 筛选 审核时间 为空的行
        self.PRApproveNotTime = self.PRApproveNotTime.loc[  # 筛选 建议订货日期 小于 当天 的行
            self.PRApproveNotTime['建议订货日期'] < datetime64(str(self.today)[:10])]

        self.PRData = self.PRData.drop(['存货编码', '存货名称', '规格型号', '数量'], axis=1)

    def mkdir(self, path):
        self.func.mkdir(path)

    def ADDSheet(self, df, name):
        book = load_workbook(f'{self.path}/RESULT/SCM/OP/采购订单转换及时率.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/SCM/OP/采购订单转换及时率.xlsx", engine='openpyxl')
        writer.book = book
        df.to_excel(writer, f"{name}", index=False)
        writer.save()

    def PRProcessConvert(self):  #
        df = pd.merge(self.PRProcessData, self.POData, on=['采购订单号', '采购订单行号'], how='left')
        # df.to_excel('./demo2.xlsx', index=False)
        df2 = pd.merge(df, self.PRData, on=['请购单行号', '请购单号'], how='left')
        Approve = df2.loc[df2["请购单审核时间"].notnull()]  # 筛选 请购单审核时间 不为空的值
        ApproveNotTime = df2.loc[df2["请购单审核时间"].isnull()]  # 筛选 请购单审核时间 为空的值
        df3 = Approve.loc[Approve["采购订单号"].notnull()]  # 筛选 采购订单号 不为空的值
        df4 = Approve.loc[Approve["采购订单号"].isnull()]  # 筛选 采购订单号 为空的值

        df3['转化延时'] = ((df3['采购订单制单时间'] - df3['请购单审核时间']) / pd.Timedelta(1, 'H')).astype(int)
        df3.loc[df3["转化延时"] > 48, "单据状态"] = "超时"
        df3.loc[df3["转化延时"] <= 48, "单据状态"] = "正常"

        df4 = df4.loc[df4['建议订货日期'] < datetime64(str(self.today)[:10])]

        df3.to_excel(f'{self.path}/RESULT/SCM/OP/采购订单转换及时率.xlsx', sheet_name="三天内未及时转化请购单", index=False)
        self.ADDSheet(df4, '历史未转化请购单')
        self.ADDSheet(self.PRApproveNotTime, '历史未审核请购单')
        self.ADDSheet(ApproveNotTime, '未完成采购并且已被关闭请购单')

    def GetOrderConversion(self):  # 三天未转化MRP
        self.YesMRPData = self.YesMRPData.loc[self.YesMRPData["物料属性"] == "采购"]
        self.YesMRPData = self.YesMRPData.loc[self.YesMRPData["是否客供料"].isnull()]

        self.MRPData = self.MRPData.loc[self.MRPData["物料属性"] == "采购"]
        self.MRPData = self.MRPData.loc[self.MRPData["是否客供料"].isnull()]
        df = pd.merge(
            self.MRPData.drop(labels=['物料名称', '物料属性', '是否客供料'],
                              axis=1), self.YesMRPData, on=['物料编码', '需求跟踪号', '需求跟踪行号', '开工日期'])
        df = df.loc[df["需求跟踪号"] == "XS20211223001"]
        df2 = df.loc[df["需求跟踪号"] != "XS20211223001"]
        self.ADDSheet(df, '挤出项目未转换MRP清单')
        self.ADDSheet(df2, '三天内未转化MRP清单')
    def run(self):
        self.PRProcessConvert()
        self.GetOrderConversion()
        # self.HistoryNotConverted()


if __name__ == '__main__':
    OC = OrderConversion()
    OC.run()
