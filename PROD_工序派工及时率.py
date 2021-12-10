import pandas as pd
import Func

pd.set_option('display.max_columns', None)


# 工序派工及时率
class ProcessDispatch:
    def __init__(self):
        self.AllDataList = []
        self.BaseDataList = []
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = Func.Path()

    def mkdir(self, path):
        self.func.mkdir(path)

    def CheckDataWork(self, base_data, new_data):
        base_data = base_data.dropna(subset=['物料编码'])  # 去除nan的列
        new_data = new_data.dropna(subset=['物料编码'])  # 去除nan的列
        out_data = pd.merge(base_data.drop(labels=['派工标识'], axis=1), new_data, on=['物料编码', '生产订单', '工序行号', '行号'])
        out_data = out_data.loc[out_data['派工标识'] != "*"]
        self.AllDataList.append(out_data)

    def SaveFile(self):
        res = pd.concat(self.AllDataList, axis=0, ignore_index=True)
        res = res.drop_duplicates()
        self.func.mkdir(self.path + '/RESULT/PROD')
        res.to_excel(f'{self.path}/RESULT/PROD/工序派工及时率.xlsx', sheet_name="工序派工及时率", index=False)

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
                                             usecols=['物料编码', '生产订单', '工序行号', '派工标识', '行号'],
                                             converters={'物料编码': int, '工序行号': int}
                                             )
                    self.BaseDataList.append(BaseData)
                    flag = flag + 1
                    continue
                except:
                    continue
            else:
                try:
                    NewData = pd.read_excel(f"{self.path}/DATA/PROD/工序派工资料维护{ThisYear}-{ThisMonth}-{work_day}.XLSX",
                                            usecols=['物料编码', '生产订单', '工序行号', '派工标识', '行号'],
                                            converters={'物料编码': int, '工序行号': int}
                                            )
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
