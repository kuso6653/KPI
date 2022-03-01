import pandas as pd
import calendar
import datetime
from datetime import timedelta
from numpy import datetime64
from openpyxl import load_workbook

import Func

pd.set_option('display.max_columns', None)


class MaterialMaintenance:
    def __init__(self):
        self.MMList = []
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = Func.Path()
        self.new_data = pd.DataFrame()

    def mkdir(self, path):
        self.func.mkdir(path)

    def CheckDataStock(self, BaseData, NewData):
        BaseData = BaseData.dropna(subset=['存货编码'])  # 去除nan的列
        NewData = NewData.dropna(subset=['存货编码'])  # 去除nan的列
        out_data = pd.merge(BaseData.drop(labels=['主要供货单位名称', '最低供应量', '采购员名称', '固定提前期', '计划默认属性', '启用日期', '停用日期', '无需采购件'], axis=1),
                            NewData,
                            on=['存货编码', '存货名称'])
        out_data = out_data[out_data.isnull().any(axis=1)]
        out_data = out_data.loc[out_data["固定提前期"] == 0]

        self.MMList.append(out_data)

    def GetMaterialMaintenance(self):  # 当月大于7天未维护的采购物料清单
        # 获取截取这个月份、年、上个月
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0]
        this_month = self.ThisMonthStart.split("-")[1]

        year = self.ThisMonthStart.split("-")[0]

        self.LastMonthStart = str(self.LastMonthStart).split(" ")[0]
        last_month = self.LastMonthStart.split("-")[1]

        last_work_days = self.func.WorkDays(year, last_month)  # 获取上个月工作日
        this_work_days = self.func.WorkDays(year, this_month)  # 获取这个月工作日
        work_days7 = []  # 设设置到上月7天

        work_days7.extend(last_work_days[-7:])
        work_days7.extend(this_work_days)  # 将上个月最后七天和这个月工作日相合并
        work_days7 = self.func.ReformDays(work_days7)  # 改造

        base_data_stock = []

        flag = 0
        for work_day in work_days7:
            if flag < 7:
                try:  # 存货档案-20211001
                    base_data = pd.read_excel(f"{self.path}/DATA/SCM/存货档案{year}-{last_month}-{work_day}.XLSX",
                                              usecols=['存货编码', '存货名称', '主要供货单位名称', '采购员名称', '最低供应量', '固定提前期', '计划默认属性',
                                                       '启用日期', '停用日期', '无需采购件', '计划方法'],
                                              converters={'最低供应量': int, '固定提前期': int}
                                              )
                    base_data = base_data.loc[base_data["计划默认属性"] == "采购"]
                    base_data = base_data.loc[base_data["计划方法"] != "N"]
                    base_data = base_data[
                        (base_data["停用日期"].isnull()) & (base_data["无需采购件"].isnull())]
                    base_data_stock.append(base_data)
                    flag = flag + 1
                    continue
                except:
                    continue
            else:
                try:
                    self.new_data = pd.read_excel(f"{self.path}/DATA/SCM/存货档案{year}-{this_month}-{work_day}.XLSX",
                                                  usecols=['存货编码', '存货名称', '主要供货单位名称', '采购员名称', '最低供应量', '固定提前期',
                                                           '计划默认属性', '启用日期', '停用日期', '无需采购件', '计划方法'],
                                                  converters={'最低供应量': int, '固定提前期': int, '启用日期': datetime64}
                                                  )
                    self.new_data = self.new_data.loc[self.new_data["计划默认属性"] == "采购"]
                    self.new_data = self.new_data.loc[self.new_data["计划方法"] != "N"]
                    self.new_data = self.new_data[
                        (self.new_data["停用日期"].isnull()) & (self.new_data["无需采购件"].isnull())]
                except:
                    continue
                base_data_stock.append(self.new_data)  # 新添加新的base
                self.CheckDataStock(base_data_stock[0], self.new_data)  # 合并检查是否存在一样的
                del (base_data_stock[0])  # 删除第一个base

    def ThisMonthNotMaintained(self):
        res = pd.concat(self.MMList, axis=0, ignore_index=True)
        res = res.drop_duplicates()
        self.new_data = self.new_data[self.new_data.isnull().any(axis=1)]
        #  小于当月的历史未维护物料数据筛选
        self.new_data = self.new_data.loc[self.new_data["固定提前期"] == 0]
        self.new_data = self.new_data[self.new_data['启用日期'] < datetime64(self.ThisMonthStart)]
        #  当月大于7天的未维护物料数据筛选
        res = res.loc[res["固定提前期"] == 0]
        res = res[res['启用日期'] >= datetime64(self.ThisMonthStart)]
        self.mkdir(self.path+'/RESULT/SCM/SP')
        res.to_excel(f'{self.path}/RESULT/SCM/SP/采购物料维护及时率.xlsx', sheet_name="当月大于7天未维护的采购物料清单", index=False)

    def HistoryNotMaintained(self):  # 历史未维护数据清单
        self.mkdir(self.path + '/RESULT/SCM/SP')
        book = load_workbook(f'{self.path}/RESULT/SCM/SP/采购物料维护及时率.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/SCM/SP/采购物料维护及时率.xlsx", engine='openpyxl')
        writer.book = book
        self.new_data.to_excel(writer, "历史未维护数据清单", index=False)
        writer.save()

    def run(self):
        self.GetMaterialMaintenance()
        self.ThisMonthNotMaintained()
        self.HistoryNotMaintained()


if __name__ == '__main__':
    MM = MaterialMaintenance()
    MM.run()
