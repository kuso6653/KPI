import pandas as pd
from numpy import datetime64
from openpyxl import load_workbook
import Func

pd.set_option('display.max_columns', None)


class SelfMaterial:
    def __init__(self):
        self.new_data = pd.DataFrame()
        self.SelfMaterialList = []
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = "//10.56.164.127/it&m/KPI"

    def ContrastData(self, BaseData, NewData):
        BaseData = BaseData.dropna(subset=['存货编码'])  # 去除nan的列
        NewData = NewData.dropna(subset=['存货编码'])  # 去除nan的列
        out_data = pd.merge(BaseData.drop(labels=['生产部门名称', '变动提前期', '变动基数', '固定提前期', '计划默认属性', '启用日期', '停用日期', '无需采购件'], axis=1),
                            NewData,
                            on=['存货编码', '存货名称'])
        out_data = out_data[out_data.isnull().any(axis=1)]
        out_data = out_data.loc[out_data["固定提前期"] == 0]

        self.SelfMaterialList.append(out_data)

    def mkdir(self, path):
        self.func.mkdir(path)

    def GetThisMonthData(self):  # 当月大于3天未维护的自制物料清单
        # 获取截取这个月份、年、上个月
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0]
        ThisMonth = self.ThisMonthStart.split("-")[1]

        NowYear = self.ThisMonthStart.split("-")[0]

        self.LastMonthStart = str(self.LastMonthStart).split(" ")[0]
        LastMonth = self.LastMonthStart.split("-")[1]

        last_work_days = self.func.WorkDays(NowYear, LastMonth)  # 获取上个月工作日
        this_work_days = self.func.WorkDays(NowYear, ThisMonth)  # 获取这个月工作日

        work_days = []  # 设置到上月3天
        work_days.extend(last_work_days[-3:])
        work_days.extend(this_work_days)  # 将上个月最后三天和这个月工作日相合并
        work_days = self.func.ReformDays(work_days)  # 改造

        BaseDataList = []
        flag = 0
        for work_day in work_days:
            if flag < 3:
                try:  # 存货档案-20211001
                    base_data = pd.read_excel(f"{self.path}/DATA/SCM/存货档案{NowYear}-{LastMonth}-{work_day}.XLSX",
                                              usecols=['存货编码', '存货名称', '计划默认属性', '固定提前期', '生产部门名称', '变动提前期', '变动基数',
                                                       '启用日期', '停用日期', '无需采购件'],
                                              converters={'最低供应量': int, '变动提前期': int, '变动基数': float}
                                              )
                    base_data = base_data.loc[base_data["计划默认属性"] == "自制"]
                    base_data = base_data[
                        (base_data["停用日期"].isnull()) & (base_data["无需采购件"].isnull())]
                    BaseDataList.append(base_data)
                    flag = flag + 1
                    continue
                except:
                    continue
            else:
                try:
                    self.new_data = pd.read_excel(f"{self.path}/DATA/SCM/存货档案{NowYear}-{ThisMonth}-{work_day}.XLSX",
                                                  usecols=['存货编码', '存货名称', '计划默认属性', '固定提前期', '生产部门名称', '变动提前期', '变动基数',
                                                           '启用日期', '停用日期', '无需采购件'],
                                                  converters={'最低供应量': int, '变动提前期': int, '变动基数': float}
                                                  )
                    self.new_data = self.new_data.loc[self.new_data["计划默认属性"] == "自制"]
                    self.new_data = self.new_data[
                        (self.new_data["停用日期"].isnull()) & (self.new_data["无需采购件"].isnull())]
                except:
                    continue
                BaseDataList.append(self.new_data)  # 新添加新的base
                self.ContrastData(BaseDataList[0], self.new_data)  # 合并检查是否存在一样的
                del (BaseDataList[0])  # 删除第一个base

        res = pd.concat(self.SelfMaterialList, axis=0, ignore_index=True)
        res = res.drop_duplicates()

        #  当月大于7天的未维护订单数据筛选
        res = res.loc[res["固定提前期"] == 0]
        res = res[res['启用日期'] >= datetime64(self.ThisMonthStart)]
        self.mkdir(self.path + '/RESULT/SCM/OM')
        res.to_excel(f'{self.path}/RESULT/SCM/OM/自制物料维护及时率.xlsx', sheet_name="当月大于3天未维护的自制物料清单", index=False)

    def GetHistoryData(self):  # 历史未维护数据清单
        self.new_data = self.new_data[self.new_data.isnull().any(axis=1)]
        #  小于当月的历史未维护订单数据筛选
        self.new_data = self.new_data.loc[self.new_data["固定提前期"] == 0]
        self.new_data = self.new_data[self.new_data['启用日期'] < datetime64(self.ThisMonthStart)]
        self.mkdir(self.path + '/RESULT/SCM/OM')
        book = load_workbook(f'{self.path}/RESULT/SCM/OM/自制物料维护及时率.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/SCM/OM/自制物料维护及时率.xlsx", engine='openpyxl')
        writer.book = book
        self.new_data.to_excel(writer, "历史未维护数据清单", index=False)
        writer.save()

    def run(self):
        self.GetThisMonthData()
        self.GetHistoryData()


if __name__ == '__main__':
    SM = SelfMaterial()
    SM.run()
