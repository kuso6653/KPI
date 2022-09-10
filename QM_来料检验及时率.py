import pandas as pd
from numpy import datetime64
from openpyxl import load_workbook
import Func


class MaterialInspection:
    def __init__(self):
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = Func.Path()
        # 将上月首尾日期切割
        self.LastMonthStart = str(self.LastMonthStart).split(" ")[0]  # .replace("-", "")
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0]
        self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0]

    def mkdir(self, path):
        self.func.mkdir(path)

    def GetPurchaseIn(self):
        PurchaseInData = pd.read_excel(f"{self.path}/DATA/SCM/采购时效性统计表.XLSX",
                                       usecols=[0, 1, 6, 7, 12, 20, 23, 27, 31], header=2,
                                       names=["行号", "订单号", "存货编码", "存货名称", "订单审核时间", "报检单号", "报检审核时间", "检验审核时间", "入库制单时间"],
                                       converters={'订单审核时间': datetime64, '报检审核时间': datetime64, '检验审核时间': datetime64,
                                                   '入库制单时间': datetime64, '存货编码': float, "订单号": str, "行号": str})
        PurchaseInData = PurchaseInData.dropna(subset=['报检审核时间'])  # 去除nan的列
        ApproveData = PurchaseInData[PurchaseInData['检验审核时间'].isnull()]  # 筛选出已报检未检验的数据 （报检审核时间有，检验审核时间有）
        ApproveData = ApproveData[ApproveData['报检审核时间'] <= datetime64(self.ThisMonthEnd)]  # 筛选出本月的单据
        ApproveData['单据状态'] = '超时'  # 已报检未检验的数据默认为 超时
        PurchaseInData = PurchaseInData.dropna(subset=['检验审核时间'])  # 去除nan的列
        PurchaseInData = PurchaseInData[PurchaseInData['报检审核时间'] >= datetime64(self.ThisMonthStart)]  # 筛选出本月的单据
        PurchaseInData = PurchaseInData[PurchaseInData['报检审核时间'] <= datetime64(self.ThisMonthEnd)]  # 筛选出本月的单据
        PurchaseInData['审批延时'] = ((PurchaseInData['检验审核时间'] - PurchaseInData['报检审核时间']) / pd.Timedelta(1, 'H')).astype(
            int)
        PurchaseInData.loc[PurchaseInData["审批延时"] > 48, "单据状态"] = "超时"  # 计算出来的质检的审批延时大于48为超时
        PurchaseInData.loc[PurchaseInData["审批延时"] <= 48, "单据状态"] = "正常"  # 小于等于48为正常
        self.SaveFile(PurchaseInData, ApproveData)

    def SaveFile(self, PurchaseInData, ApproveData):
        self.mkdir(self.path + "/RESULT/QM")
        PurchaseInData.to_excel(f'{self.path}/RESULT/QM/来料检验及时率.xlsx', sheet_name="来料检验及时率", index=False)

        book = load_workbook(f'{self.path}/RESULT/QM/来料检验及时率.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/QM/来料检验及时率.xlsx", engine='openpyxl')
        writer.book = book
        ApproveData.to_excel(writer, "已报检未检验数据", index=False)
        writer.save()

    def run(self):
        self.GetPurchaseIn()


if __name__ == '__main__':
    MI = MaterialInspection()
    MI.run()
