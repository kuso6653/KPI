import pandas as pd
from numpy import datetime64
from openpyxl import load_workbook
import Func


class Warehouse:
    def __init__(self):
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = "//10.56.164.127/it&m/KPI"

        # 将本月月首尾日期切割
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0].replace("-", "")
        self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0].replace("-", "")
        self.PurchaseInData = pd.read_excel(
            f"{self.path}/DATA/SCM/采购时效性统计表-{self.ThisMonthStart}-{self.ThisMonthEnd}.XLSX",
            usecols=[1, 6, 7, 12, 22, 26, 30], header=3,
            names=["订单号", "存货编码", "存货名称", "订单制单时间", "报检审核时间", "检验审核时间", "入库制单时间"],
            converters={'订单制单时间': datetime64, '报检审核时间': datetime64,
                        '检验审核时间': datetime64,
                        '入库制单时间': datetime64, '存货编码': float})
        self.MaterialOutData = pd.read_excel(
            f"{self.path}/DATA/SCM/WM/材料出库单列表-{self.ThisMonthStart}-{self.ThisMonthEnd}.XLSX",
            usecols=['出库单号', '材料编码', '物料描述', '审核时间', '制单时间'],
            converters={'材料编码': str, '出库单号': str})

    def mkdir(self, path):
        self.func.mkdir(path)

    def GetWarehouse(self):
        self.PurchaseInData = self.PurchaseInData.dropna(axis=0, how='any')  # 去除所有nan的列
        self.PurchaseInData['审批延时'] = ((self.PurchaseInData['入库制单时间'] - self.PurchaseInData['检验审核时间'])
                                       / pd.Timedelta(1, 'H')).astype(int)  # 制单时间相减，然后减去 质检的审核时间
        # 将天数转化为小时数
        self.PurchaseInData.loc[self.PurchaseInData["审批延时"] > 72, "单据状态"] = "超时"  # 计算出来的审批延时大于72为超时
        self.PurchaseInData.loc[self.PurchaseInData["审批延时"] <= 72, "单据状态"] = "正常"  # 小于等于72为正常

        # 材料出库及时率

        self.MaterialOutData = self.MaterialOutData.dropna(axis=0, how='any')  # 去除所有nan的列

        self.MaterialOutData['审批延时'] = (
                (self.MaterialOutData['审核时间'] - self.MaterialOutData['制单时间']) / pd.Timedelta(1, 'H')).astype(
            int)
        # 将天数转化为小时数
        self.MaterialOutData.loc[self.MaterialOutData["审批延时"] > 72, "单据状态"] = "超时"  # 计算出来的审批延时大于72为超时
        self.MaterialOutData.loc[self.MaterialOutData["审批延时"] <= 72, "单据状态"] = "正常"  # 小于等于72为正常

        self.MaterialOutData = self.MaterialOutData.drop_duplicates()  # 去重

    def PurchaseIn(self):  # 仓库入库及时率
        self.mkdir(self.path+"/RESULT/SCM/WM")
        self.PurchaseInData.to_excel(f'{self.path}/RESULT/SCM/WM/仓库出入库及时率.xlsx', sheet_name="仓库入库及时率", index=False)

    def MaterialOut(self):  # 仓库出库及时率
        book = load_workbook(f'{self.path}/RESULT/SCM/WM/仓库出入库及时率.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/SCM/WM/仓库出入库及时率.xlsx", engine='openpyxl')
        writer.book = book
        self.MaterialOutData.to_excel(writer, "仓库出库及时率", index=False)
        writer.save()

    def run(self):
        self.GetWarehouse()
        self.PurchaseIn()
        self.MaterialOut()


if __name__ == '__main__':
    W = Warehouse()
    W.run()
