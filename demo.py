import Func
import pandas as pd
from numpy import datetime64
class Demo:
    def __init__(self):
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = Func.Path()
        self.ProductionData = pd.read_excel(
            f"{self.path}/DATA/PROD/生产订单列表.XLSX",
            usecols=['生产订单号', '行号', '物料编码', '物料名称', '生产批号', '制单时间', '类型'],
            converters={'生产订单号': str, '物料编码': str, '生产批号': str})
    def run(self):
        self.ProductionData['制单时间'] = self.ProductionData['制单时间'].astype(datetime64)
if __name__ == '__main__':
    demo = Demo()
    demo.run()
