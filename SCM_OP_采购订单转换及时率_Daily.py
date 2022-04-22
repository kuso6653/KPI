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
        self.ThrAgoMRPData = None
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
                self.ThrAgoMRPData = pd.read_excel(f"{self.path}/DATA/SCM/OP/MRP计划维护--全部{str(yesterday)[:10]}.XLSX",
                                                   usecols=['物料编码', '物料名称', '需求跟踪号', '需求跟踪行号', '物料属性', '是否客供料', '开工日期',
                                                            '采购员名称'],
                                                    converters={'物料编码': int, '需求跟踪行号': int, '开工日期': str}
                                                   )
                self.ThrAgoMRPData = self.ThrAgoMRPData.rename(columns={'采购员名称': '默认采购员'})

                break
            except Exception as e:
                print(e)
                yesterday = yesterday + timedelta(days=-1)

        yesterday = yesterday + timedelta(days=-4)  # 减去三天 再减去4天
        while 1:  # 获取前7天MRP数据，没有则继续往前推一天，直到有为止
            try:
                self.SevAgoMRPData = pd.read_excel(f"{self.path}/DATA/SCM/OP/MRP计划维护--全部{str(yesterday)[:10]}.XLSX",
                                                   usecols=['物料编码', '物料名称', '需求跟踪号', '需求跟踪行号', '物料属性', '是否客供料', '开工日期','采购员名称'],
                                                   converters={'物料编码': int, '需求跟踪行号': int, '开工日期': str}
                                                   )
                self.SevAgoMRPData = self.SevAgoMRPData.rename(columns={'采购员名称': '默认采购员'})
                break
            except Exception as e:
                print(e)
                yesterday = yesterday + timedelta(days=-1)

        self.TodayMRPData = pd.read_excel(f"{self.path}/DATA/SCM/OP/MRP计划维护--全部{str(self.today)[:10]}.XLSX",
                                          usecols=['物料编码', '物料名称', '需求跟踪号', '需求跟踪行号', '物料属性', '是否客供料', '开工日期', '采购员名称'],
                                          converters={'物料编码': int, '需求跟踪行号': int, '开工日期': str}
                                          )
        self.TodayMRPData = self.TodayMRPData.rename(columns={'采购员名称': '默认采购员'})
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
                                           usecols=[1, 6, 7, 8, 9, 10, 14, 15, 16, 20, 23], header=4,
                                           names=["请购单号", "请购单行号", "存货编码", "存货名称", "规格型号", "数量", "采购订单号", "采购订单行号",
                                                  "订单日期", "计划到货日期", "采购订单制单人"],
                                           converters={'请购单号': str, '请购单行号': str, '存货编码': str, '数量': float,
                                                       '采购订单号': str, '采购订单行号': str, '计划到货日期': datetime64, '订单日期': datetime64})

        self.PRData = pd.read_excel(f"{self.path}/DATA/SCM/OP/请购单列表.XLSX",
                                    usecols=['单据号', '行号', '行关闭人', '建议订货日期', '审核时间', '制单时间', '存货编码', '存货名称', '规格型号',
                                             '数量', '执行采购员'],
                                    converters={'单据号': str, '行号': str, '审核时间': datetime64, '制单时间': datetime64,
                                                '建议订货日期': datetime64, '存货编码': str, '数量': float})
        self.PRData = self.PRData.rename(
            columns={'行号': '请购单行号', '单据号': '请购单号', '制单时间': '请购单制单时间', '审核时间': '请购单审核时间', '执行采购员': '默认采购员'})

        self.PRData = self.PRData.loc[self.PRData["行关闭人"].isnull()]  # 筛选 行关闭人 为空的行
        self.PRApproveNotTime = self.PRData.loc[self.PRData["请购单审核时间"].isnull()]  # 筛选 审核时间 为空的行
        self.PRApproveNotTime = self.PRApproveNotTime.loc[  # 筛选 建议订货日期 小于 当天 的行
            self.PRApproveNotTime['建议订货日期'] < datetime64(str(self.today)[:10])]

        self.PRData = self.PRData.drop(['存货编码', '存货名称', '规格型号', '数量'], axis=1)

    def mkdir(self, path):
        self.func.mkdir(path)

    def ADDSheet(self, df, name):
        book = load_workbook(f'{self.path}/RESULT/SCM/OP/采购订单转换及时率{str(self.today)[:10]}.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/SCM/OP/采购订单转换及时率{str(self.today)[:10]}.xlsx", engine='openpyxl')
        writer.book = book
        df.to_excel(writer, f"{name}", index=False)
        writer.save()

    def PRProcessConvert(self):  #
        df = pd.merge(self.PRProcessData, self.POData, on=['采购订单号', '采购订单行号'], how='left')
        df2 = pd.merge(df, self.PRData, on=['请购单行号', '请购单号'], how='left')
        Approve = df2.loc[df2["请购单审核时间"].notnull()]  # 筛选 请购单审核时间 不为空的值
        ApproveNotTime = df2.loc[df2["请购单审核时间"].isnull()]  # 筛选 请购单审核时间 为空的值
        df3 = Approve.loc[Approve["采购订单号"].notnull()]  # 筛选 采购订单号 不为空的值
        df4 = Approve.loc[Approve["采购订单号"].isnull()]  # 筛选 采购订单号 为空的值
        # df3.to_excel(f'./demo.xlsx', index=False)
        df3['转化延时'] = ((df3['采购订单制单时间'] - df3['请购单审核时间']) / pd.Timedelta(1, 'H')).astype(int)
        df4 = df4.loc[df4['建议订货日期'] < datetime64(str(self.today)[:10])]

        df3.to_excel(f'{self.path}/RESULT/SCM/OP/采购订单转换及时率{str(self.today)[:10]}.xlsx', sheet_name="二月份至今的请购单列表",
                     index=False)
        self.ADDSheet(df4, '二月份至今的未转化请购单')
        self.ADDSheet(self.PRApproveNotTime, '二月份至今的未审核请购单')
        self.ADDSheet(ApproveNotTime, '二月份至今未完成采购并且已被关闭请购单')

    def GetOrderConversion(self):  # 三天未转化MRP
        self.ThrAgoMRPData = self.ThrAgoMRPData.loc[self.ThrAgoMRPData["物料属性"] == "采购"]
        self.ThrAgoMRPData = self.ThrAgoMRPData.loc[self.ThrAgoMRPData["是否客供料"].isnull()]

        self.SevAgoMRPData = self.SevAgoMRPData.loc[self.SevAgoMRPData["物料属性"] == "采购"]
        self.SevAgoMRPData = self.SevAgoMRPData.loc[self.SevAgoMRPData["是否客供料"].isnull()]

        self.TodayMRPData = self.TodayMRPData.loc[self.TodayMRPData["物料属性"] == "采购"]
        self.TodayMRPData = self.TodayMRPData.loc[self.TodayMRPData["是否客供料"].isnull()]

        ThrAgoMRPData = pd.merge(
            self.TodayMRPData.drop(labels=['物料名称', '物料属性', '是否客供料', '默认采购员'],
                                   axis=1), self.ThrAgoMRPData, on=['物料编码', '需求跟踪号', '需求跟踪行号', '开工日期'])

        SevAgoMRPData = pd.merge(
            self.TodayMRPData.drop(labels=['物料名称', '物料属性', '是否客供料', '默认采购员'],
                                   axis=1), self.SevAgoMRPData, on=['物料编码', '需求跟踪号', '需求跟踪行号', '开工日期'])
        Extruder = ThrAgoMRPData.loc[ThrAgoMRPData["需求跟踪号"] == "XS20211223001"]
        ThrAgoMRPData = ThrAgoMRPData.loc[ThrAgoMRPData["需求跟踪号"] != "XS20211223001"]
        SevAgoMRPData = SevAgoMRPData.loc[SevAgoMRPData["需求跟踪号"] != "XS20211223001"]

        ThrAgoMRPData['转化延迟时间'] = '三天前'
        SevAgoMRPData['转化延迟时间'] = '七天前'

        ThrAgoMRPData = ThrAgoMRPData[
            ~ ThrAgoMRPData['物料名称'].str.contains('包装费用|碳素钢板|包装箱|电气图纸|备选部件|电气外部辅材|润滑管道系统|气控管道系统')]
        ThrAgoMRPData = ThrAgoMRPData[
            ~ ThrAgoMRPData['物料编码'].isin(
                ['888012200', '888500351', '888500352', '888500353', '888500354', '888500355', '888500356'])]

        SevAgoMRPData = SevAgoMRPData[
            ~ SevAgoMRPData['物料名称'].str.contains('包装费用|碳素钢板|包装箱|电气图纸|备选部件|电气外部辅材|润滑管道系统|气控管道系统')]
        SevAgoMRPData = SevAgoMRPData[
            ~ SevAgoMRPData['物料编码'].isin(
                ['888012200', '888500351', '888500352', '888500353', '888500354', '888500355', '888500356'])]

        self.ADDSheet(Extruder, '挤出项目未转换MRP清单')
        self.ADDSheet(ThrAgoMRPData, '三天内未转化MRP清单')
        self.ADDSheet(SevAgoMRPData, '七天内未转化MRP清单')

    def run(self):
        self.PRProcessConvert()
        self.GetOrderConversion()
        # self.HistoryNotConverted()


if __name__ == '__main__':
    OC = OrderConversion()
    OC.run()
