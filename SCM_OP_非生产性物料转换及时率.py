import re
import time
from datetime import datetime
from openpyxl import load_workbook
import pandas as pd
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
# from gevent import monkey
from numpy import datetime64
import gevent, time, requests
from gevent.queue import Queue
import Func
from dateutil.relativedelta import relativedelta
from get_holiday_cn.client import getHoliday

today = datetime.datetime.today()  # datetime类型当前日期
yesterday2 = today - relativedelta(months=2)  # 往时间推迟两个月，取当前和前两个月共三个月数据
yesterday2 = str(yesterday2).split(' ')[0].replace('-', '')[:6]
EarlyTime = int(yesterday2)


def GetHoliday(date):
    client = getHoliday()
    # 指定日期获取数据
    date = client.assemble_holiday_data(date)
    """
            {
              "code": 0,              // 0服务正常。-1服务出错
              "type": {
                "type": enum(0, 1, 2, 3), // 节假日类型，分别表示 工作日、周末、节日、调休
    """
    return date["code"], date["type"]["type"], date["type"]["week"]


class GetOAFunc:
    def __init__(self, Cookie):
        self.Cookie = Cookie
        self.PR = 'SelectPurchase'
        self.headers = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                          'Chrome/71.0.3578.98 Safari/537.36',
            'Cookie': self.Cookie
        }
        self.list = []
        self.func = Func
        self.path = self.func.Path()
        self.Purchase_in_data = pd.read_excel(f"{self.path}/DATA/SCM/OP/采购订单列表.XLSX",
                                              usecols=['存货编码', '存货名称', '订单编号', '主计量', '数量', '制单时间'],
                                              converters={'存货编码': str, '订单编号': str, '制单时间': datetime64}
                                              )
        self.Purchase_in_data = self.Purchase_in_data.dropna(subset=['存货编码'])  # 去除nan的列
        self.Purchase_in_data = self.Purchase_in_data.rename(columns={'订单编号': 'U8采购单号'})
        self.PurchaseConversionList = pd.DataFrame(columns=["申请单号", "申请人", "申请日期", "U8采购单号", "审批日期"])
        self.work = Queue()  # 创建队列对象，并赋值给work。

    def GetPRUidList(self, url):  # 获取PR目录中的时间以及详细页面的uid
        flag = True
        res = self.GetRequest(url)
        uid = re.findall(r'unid="(.*?)"', res.text)
        code = re.findall(r'<entrydata columnnumber="0" name="FormID">\n<text>(.*?)</text>', res.text)
        for i, j in zip(uid, code):
            try:
                if int(j[:6]) >= EarlyTime:
                    self.list.append(i)
                    # self.work.put_nowait(i)
                else:
                    flag = False
            except Exception as e:
                print(e)
                continue
        return flag

    def GetRequest(self, url):  # 用于网络申请返回res
        try:
            res = requests.get(url, headers=self.headers)
        except Exception as e:
            print(e)
            time.sleep(3)
            res = requests.get(url, headers=self.headers)
        return res

    def GetPRData(self):  # 获取PR非生产性物料
        # while not self.work.empty():
        for uid in self.list:
            # uid = self.work.get_nowait()
            url1 = f'http://portal.chemchina.com/oa08/dept509/{self.PR}.nsf/vwAll/{uid}?opendocument'
            UserInfo = self.GetRequest(url1).content.decode('utf-8')
            # pid用于获取流转意见的request拼接
            pid = re.findall(r'<div class="btn-group btn-corner" data-addType="append" data-ajax="(.*?)">',
                             UserInfo)
            pid = str(pid[0]).split('&')[1]
            # 是否是结束流程
            is_end = re.findall(r'<input name="TFCurNodeName" type="hidden" value="(.*?)">', UserInfo)[0]
            Application_name = re.findall(
                r'<input name="ApplyPsnCN" value="(.*?)" id="ApplyPsnCN"',
                UserInfo)  # 申请人
            code = re.findall(
                r'<input name="FormID" value="(.*?)"', UserInfo)  # OA申请单号
            U8_code = re.findall(
                r'<input name="fldU8Nun" value="(.*?)" class="form-control"',  # U8采购单号
                UserInfo)
            Application_date = re.findall(
                r'<input name="ApplyDate" value="(.*?)" id="ApplyDate"',  # OA申请日期
                UserInfo)

            if U8_code[0] != "" and is_end == "结束":
                url2 = f'http://portal.chemchina.com/oa08//flow/flowprocess2021.nsf/FlowMind?readform&{pid}'
                WorkTime = self.GetRequest(url2).content.decode('utf-8')
                time_list = re.findall(r'<div class="col-md-4"><b>时间：</b>(.*?)</div>', WorkTime)
                approval_list = re.findall(r"""<div class="col-md-4"><b>审批节点：</b>(.*?)</div>""", WorkTime)  # 审批节点
                for i, j in zip(approval_list, time_list):
                    if self.func.RelevantPersonnel(i):
                        # ["申请单号", "申请人", "申请日期", "U8采购单号", "审批日期"
                        self.PurchaseConversionList = self.PurchaseConversionList.append(
                            {"申请单号": code[0], "申请人": Application_name[0],
                             "申请日期": Application_date[0], "U8采购单号": U8_code[0],
                             "审批日期": j},
                            ignore_index=True)
                        break

    def PRToTime(self):  # 计算PR的时间，并保存
        self.PurchaseConversionList["审批日期"] = self.PurchaseConversionList["审批日期"].astype(datetime64)
        self.merge_data = pd.merge(self.PurchaseConversionList, self.Purchase_in_data, on=["U8采购单号"])
        self.merge_data['未及时率/H'] = (
                (self.merge_data['制单时间'] - self.merge_data['审批日期']) / pd.Timedelta(1, 'H')).astype(
            int)
        self.merge_data.loc[self.merge_data["未及时率/H"] > 384, "创建及时率"] = "超时"  # 计算出来的审批延时大于16天为超时
        self.merge_data.loc[self.merge_data["未及时率/H"] <= 384, "创建及时率"] = "正常"  # 小于等于16天为正常

        path = f"{self.path}/RESULT/SCM/OP"
        file_path = path + '/' + '非生产性物料转换及时率' + '.xlsx'
        self.merge_data.to_excel(file_path, index=False)

    def run(self):

        url = f'http://portal.chemchina.com/oa08/dept509/{self.PR}.nsf/vwAll?ReadViewEntries&start=1&count=20'
        flag = self.GetPRUidList(url)
        time.sleep(1)
        page = 1
        while flag:  # RPUidList中未超出设定时间为True,否则跳出循环
            # 从第二页开始，需要获取page对应的start
            page_url = f'http://portal.chemchina.com/oa08/flow/homepage.nsf/agtFixViewPosition?openagent&db=dept509/{self.PR}.nsf&vw=vwAll&cat=&page={page}&count=20'
            page_num = self.GetRequest(page_url).text.replace('\n', '')  # 截取对应的start，系统根据不同的账号会返回不同的start
            url3 = f'http://portal.chemchina.com/oa08/dept509/{self.PR}.nsf/vwAll?ReadViewEntries&start={page_num}&count=20'  # 这个才是真正的目录列表url
            flag = self.GetPRUidList(url3)
            page = page + 1
        # 暂时去掉异步调用
        # task_list = []
        # for x in range(4):  # 创立四个线程
        #     task = gevent.spawn(self.GetPRData)
        #     task_list.append(task)
        # gevent.joinall(task_list)
        self.GetPRData()
        self.PRToTime()


# 读取指定文件
def ReadTxT():
    with open("//10.56.164.228/KPI/员工手册.txt") as f:
        txt = f.read()
        f.close()
    return txt


if __name__ == '__main__':
    driver = webdriver.Chrome()  # 创建对象，启动谷歌浏览器
    driver.implicitly_wait(20)  # 隐式等待
    txt = ReadTxT()  # 获取密码
    driver.get("http://portal.chemchina.com/")  # 请求url
    driver.find_element(By.NAME, 'username').send_keys('fjthadmin')  # 输入账号密码
    driver.find_element(By.NAME, 'password').send_keys(txt)  # 输入账号密码
    driver.implicitly_wait(20)
    driver.find_element(By.XPATH, '//*[@id="warp"]/form/div[2]/div[3]/div[4]/input[3]').click()  # 点击登录
    driver.implicitly_wait(20)
    # 获取和拼接cookie
    cookie = driver.get_cookies()
    cookie = cookie[2]['name'] + '=' + cookie[2]['value'] + ';' + cookie[1]['name'] + '=' + cookie[1][
        'value'] + ';' + cookie[0]['name'] + '=' + cookie[0]['value']
    driver.quit()  # 退出浏览器
    getOA = GetOAFunc(cookie)
    getOA.run()
