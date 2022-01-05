import re

import Func
import openpyxl
from openpyxl import load_workbook, Workbook
import pandas as pd


# -*- coding:utf-8 -*-
# EqualRe = re.findall("(^=)", "dsfasdf")
# BOMRe = re.findall("[B][O][M]", "Password:TH@123456，FunType:OM，Mode：保存，Data")
from OracleHelper import OracleHelper


def HaveLog(strRE):  # 判断该行是否有等号
    EqualRe = re.findall("(^=)", strRE)
    try:
        if EqualRe[0] == "=":
            return True
    except:
        return False


def HaveBOM(strRE):  # 判断该行是否有BOM字段
    BOMRe = re.findall("[F][u][n][T][y][p][e][:][B][O][M]", strRE)
    try:
        if BOMRe[0] == "FunType:BOM":
            return True
    except:
        return False


def HaveRow(strRE):  # 判断该行是否有iRowNo字段
    RowRe = re.findall("[i][R][o][w][N][o]", strRE)
    try:
        if RowRe[0] == "iRowNo":
            return True
    except:
        return False


class BOM:
    def __init__(self):
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()

        self.BOMList = []
        self.oneStr = 1
        self.cursor = 0  # 设置游标

        self.BOMData = pd.DataFrame(columns=['cProNo', 'cFaInvCode', 'time'])
        self.TagTime = ""
        self.wb = Workbook()
        self.wb.save('./BOM.xlsx')
        sql_PRO = "select t.code,t.name,t.customername from TN_A_PROJECT t"
        SqlOracle = OracleHelper("T5_ENTITY", "thsoft", "10.56.164.22:1521/THPLM")
        sql_PRO_list = SqlOracle.find_sql(sql_PRO)
        self.PROData = pd.DataFrame(sql_PRO_list, columns=['cProNo', 'name', 'customername'])

    def getTXT(self, lines, this_date):
        flag = 1
        for index, value in enumerate(lines):
            if HaveLog(value) and flag == 0:  # 有等号并且第一次遇到
                self.cursor = 0
                #  = pd.DataFrame(columns=['cProNo', 'cFaInvCode', 'time'])
                flag = 1  # 设置第二次遇到直接跳过直到遇到BOM 字段
            if HaveBOM(value):  # 有BOM
                self.cursor = 1  # 将游标置为 1
            if self.cursor == 1:
                self.TagTime = lines[index - 2]
                self.cursor = 2
            if self.cursor == 2:
                flag = 0
                if HaveRow(value):
                    self.BOMData = self.BOMData.append({"cProNo": lines[index + 1].replace('cProNo', '').replace(',',
                                                                                                                 '').replace(
                        '"', '').replace(':', '').replace(' ', '').replace('\n', ''),
                                                        "cFaInvCode": lines[index + 2].replace('cFaInvCode',
                                                                                               '').replace(',',
                                                                                                           '').replace(
                                                            '"', '').replace(':', '').replace(' ', '').replace('\n',
                                                                                                               ''),
                                                        "time": self.TagTime.replace('时间:', '').replace('\n', '')},
                                                       ignore_index=True)

    def run(self):
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0]
        this_month = self.ThisMonthStart.split("-")[1]
        year = self.ThisMonthStart.split("-")[0]
        month = self.ThisMonthStart.split("-")[1]
        EveryDays = self.func.WorkDays(year, this_month)  # 取本月所有日期
        EveryDays = self.func.ReformDays(EveryDays)  # 改造

        for i in EveryDays:
            this_date = str(year) + str(month) + str(i)
            try:
                FileOpen = open(f'./U8接口{this_date}_u8log.txt', "rt", encoding="utf-8")
                lines = FileOpen.readlines()
                self.getTXT(lines, this_date)
            except:
                continue
        self.BOMData = pd.merge(self.BOMData, self.PROData, how="left", on="cProNo")
        max_data_list = []
        for name, group in self.BOMData.groupby(["cProNo", "cFaInvCode"]):
            group = pd.DataFrame(group)  # 新建pandas
            group = group.sort_values(by='time', ascending=False)  # 降序排序
            max_data_list.append(group.head(1))
        self.BOMData = pd.concat(max_data_list, axis=0, ignore_index=True)
        self.BOMData.to_excel('./BOM.xlsx', index=False)


if __name__ == '__main__':
    bom = BOM()
    bom.run()
