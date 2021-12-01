import pandas as pd
from numpy import datetime64

import Func


class MaterialInspection:
    def __init__(self):
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = "//10.56.164.127/it&m/KPI"
        # 将上月首尾日期切割
        self.LastMonthStart = str(self.LastMonthStart).split(" ")[0].replace("-", "")
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0].replace("-", "")
        self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0].replace("-", "")

    def mkdir(self, path):
        self.func.mkdir(path)

    def GetPurchaseIn(self):
        PurchaseInData = pd.read_excel(f"{self.path}/DATA/SCM/采购时效性统计表-{self.LastMonthStart}-{self.ThisMonthEnd}.XLSX",
                                       usecols=[1, 6, 7, 12, 19, 22, 26, 30], header=2,
                                       names=["订单号", "存货编码", "存货名称", "订单制单时间", "报检单号", "报检审核时间", "检验审核时间", "入库制单时间"],
                                       converters={'订单制单时间': datetime64, '报检审核时间': datetime64, '检验审核时间': datetime64,
                                                   '入库制单时间': datetime64, '存货编码': float, "订单号": str})
        PurchaseInData = PurchaseInData.dropna(subset=['报检审核时间'])  # 去除nan的列
        PurchaseInData['审批延时'] = ((PurchaseInData['检验审核时间'] - PurchaseInData['报检审核时间']) / pd.Timedelta(1, 'H')).astype(
            int)
        PurchaseInData.loc[PurchaseInData["审批延时"] > 24, "单据状态"] = "超时"  # 计算出来的质检的审批延时大于24为超时
        PurchaseInData.loc[PurchaseInData["审批延时"] <= 24, "单据状态"] = "正常"  # 小于等于24为正常
        self.SaveFile(PurchaseInData)

    def SaveFile(self, PurchaseInData):
        self.mkdir(self.path + "/RESULT/QM")
        PurchaseInData.to_excel(f'{self.path}/RESULT/QM/来料检验及时率.xlsx', sheet_name="来料检验及时率", index=False)

    def run(self):
        self.GetPurchaseIn()


if __name__ == '__main__':
    MI = MaterialInspection()
    MI.run()
