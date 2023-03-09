import pandas as pd
import Func
from numpy import datetime64
from openpyxl import load_workbook

pd.set_option('display.max_columns', None)


# 工序派工及时率
class ProcessDispatch:
    def __init__(self):
        self.AllDataList = []
        self.DispatchDataList = []
        self.BaseDataList = []
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = Func.Path()

        self.WorkCenterData = pd.read_excel(
            f"{self.path}/DATA/PROD/工作中心维护.XLSX",
            usecols=['工作中心代号', '部门名称'],
            converters={'工作中心代号': str})
        self.WorkCenterData = self.WorkCenterData.rename(columns={'工作中心代号': '工作中心', '部门名称': '生产部门'})

    def mkdir(self, path):
        self.func.mkdir(path)

    def CheckDataWork(self, base_data, new_data):
        base_data = base_data.dropna(subset=['物料编码'])  # 去除nan的列
        new_data = new_data.dropna(subset=['物料编码'])  # 去除nan的列
        out_data = pd.merge(base_data.drop(labels=['派工标识', '工作中心', '生产部门', '工序开工日', '工序完工日'], axis=1),
                            new_data, on=['物料编码', '生产订单', '工序行号', '行号', '物料名称'])
        out_data = out_data[out_data['工序开工日'] >= datetime64(self.ThisMonthStart)]
        out_data = out_data[out_data['工序开工日'] <= datetime64(self.ThisMonthEnd)]
        Dispatch_data = out_data
        out_data = out_data.loc[out_data['派工标识'] != "*"]
        self.AllDataList.append(out_data)
        self.DispatchDataList.append(Dispatch_data)

    def SaveFile(self):
        res = pd.concat(self.AllDataList, axis=0, ignore_index=True)
        res = res.drop_duplicates()
        Dispatch = pd.concat(self.DispatchDataList, axis=0, ignore_index=True)
        Dispatch = Dispatch.drop_duplicates()

        DispatchResult = (1 - res.shape[0] / Dispatch.shape[0])
        DispatchResult = format(float(DispatchResult), '.2%')
        dict = {'当月未派工总数': [res.shape[0]], '当月需派工单总数': [Dispatch.shape[0]], '当月派工及时率': [DispatchResult]}
        DispatchResult_Sheet = pd.DataFrame(dict)

        self.func.mkdir(self.path + '/RESULT/PROD')
        DispatchResult_Sheet.to_excel(f'{self.path}/RESULT/PROD/工序派工及时率.xlsx', sheet_name="当月派工及时率", index=False)
        book = load_workbook(f'{self.path}/RESULT/PROD/工序派工及时率.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/PROD/工序派工及时率.xlsx", engine='openpyxl')
        writer.book = book
        res.to_excel(writer, "当月未派工工序清单", index=False)
        writer.save()

    def GetProcessDispatch(self):
        # 获取截取这个月份、年、上个月
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0]
        ThisMonth = self.ThisMonthStart.split("-")[1]

        ThisYear = self.ThisMonthStart.split("-")[0]

        self.LastMonthEnd = str(self.LastMonthEnd).split(" ")[0]
        LastMonth = self.LastMonthEnd.split("-")[1]

        last_work_days = self.func.WorkDays(ThisYear, LastMonth)  # 获取上个月工作日
        this_work_days = self.func.WorkDays(ThisYear, ThisMonth)  # 获取这个月工作日
        WorkDaysList = []  # 设置到上月3天
        WorkDaysList.extend(last_work_days[-3:])
        WorkDaysList.extend(this_work_days)  # 将上个月最后三天和这个月工作日相合并

        WorkDaysList = self.func.ReformDays(WorkDaysList)  # 改造

        flag = 0
        for work_day in WorkDaysList:
            if flag < 3:
                try:
                    BaseData = pd.read_excel(f"{self.path}/DATA/PROD/工序派工资料维护{ThisYear}-{LastMonth}-{work_day}.XLSX",
                                             usecols=['物料编码', '生产订单', '工序行号', '派工标识', '行号', '物料名称', '工作中心',
                                                      '工序开工日', '工序完工日'],
                                             converters={'物料编码': int, '工序行号': int, '工序开工日': datetime64,
                                                         '工序完工日': datetime64})
                    BaseData = pd.merge(BaseData, self.WorkCenterData, how="left", on=['工作中心'])  # 匹配生产部门
                    self.BaseDataList.append(BaseData)
                    flag = flag + 1
                    continue
                except:
                    flag = flag + 1
                    continue
            else:
                try:
                    NewData = pd.read_excel(f"{self.path}/DATA/PROD/工序派工资料维护{ThisYear}-{ThisMonth}-{work_day}.XLSX",
                                            usecols=['物料编码', '生产订单', '工序行号', '派工标识', '行号', '物料名称', '工作中心',
                                                     '工序开工日', '工序完工日'],
                                            converters={'物料编码': int, '工序行号': int, '工序开工日': datetime64,
                                                        '工序完工日': datetime64})
                    NewData = pd.merge(NewData, self.WorkCenterData, how="left", on=['工作中心'])  # 匹配生产部门
                except:
                    continue
                self.BaseDataList.append(NewData)  # 新添加新的base
                self.CheckDataWork(self.BaseDataList[0], NewData)  # 合并检查是否存在一样的
                del (self.BaseDataList[0])  # 删除第一个base

    def run(self):
        self.GetProcessDispatch()
        self.SaveFile()


if __name__ == '__main__':
    PD = ProcessDispatch()
    PD.run()
