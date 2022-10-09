import re
import time
from datetime import datetime
from openpyxl import load_workbook
import pandas as pd
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from gevent import monkey

import gevent, time, requests
from gevent.queue import Queue

# monkey.patch_all()
import Func
from dateutil.relativedelta import relativedelta

today = datetime.datetime.today()  # datetime类型当前日期
yesterday2 = today - relativedelta(months=2)  # 往时间推迟两个月，取当前和前两个月共三个月数据
yesterday2 = str(yesterday2).split(' ')[0].replace('-', '')[:6]
EarlyTime = int(20190709)


class GetOAFunc:
    def __init__(self):
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
        self.file_name = "/RESULT/SCM/OP/询比价流程"
        self.Cookie = cookie
        self.func = Func
        self.path = Func.Path()
        self.IP = 'InquiryParity'
        self.headers = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                          'Chrome/71.0.3578.98 Safari/537.36',
            'Cookie': self.Cookie
        }
        self.PR_uid_list = []
        self.PR_Productive = pd.DataFrame(columns=('申请单号', '申请部门', '申请人', '确认供应商', '比价明细'))
        # self.work = Queue()  # 创建队列对象，并赋值给work。

    def GetRequest(self, url):  # 用于网络申请返回res
        try:
            res = requests.get(url, headers=self.headers)
        except Exception as e:
            print(e)
            time.sleep(3)
            res = requests.get(url, headers=self.headers)
        return res

    def GetPRUidList(self, url):
        flag = True
        res = self.GetRequest(url)
        uid = re.findall(r'unid="(.*?)"', res.text)
        code = re.findall(r'<entrydata columnnumber="0" name="AppCreated">\n<text>(.*?)</text>', res.text)
        for i, j in zip(uid, code):
            j = j.replace('-', '')
            try:
                if int(j[:8]) > EarlyTime:
                    self.PR_uid_list.append(i)
                else:
                    self.PR_uid_list.append(i)
                    flag = False
            except Exception as e:
                print(e)
                continue

        return flag

    def GetPRData(self):
        for uid in self.PR_uid_list:
            # while not self.work.empty():
            #     uid = self.work.get_nowait()
            url1 = f'http://portal.chemchina.com/oa08/dept509/{self.IP}.nsf/vwAll/{uid}?opendocument'
            UserInfo = self.GetRequest(url1).content.decode('utf-8')
            # 订单编号
            uid = re.findall(r'<input name="FormID" value="(.*?)" ', UserInfo)[0] # 申请单号
            Department = re.findall(r'<input name="ApplyDept" value="(.*?)"', UserInfo)[0]  # 申请部门
            Applicant = re.findall(r'<input name="ApplyPsnCN" value="(.*?)"', UserInfo)[0]  # 申请人
            CostCenter = re.findall(r'<input name="fldQrGys" value="(.*?)" ', UserInfo)[0]  # 确认供应商
            detailed1 = str(re.findall(r'<input name="XJData1" value="(.*?)"', UserInfo)[0]).split('$^$')  #  比价明细供应商
            detailInfo = []
            for i in detailed1:
                if i not in detailInfo:
                    detailInfo.append(i)
            self.PR_Productive = self.PR_Productive.append(
                    [{'申请单号': uid, '申请部门': Department, '申请人': Applicant, '确认供应商': CostCenter,
                      '比价明细': ','.join(detailInfo)}], ignore_index=True)

    def save(self):
        self.PR_Productive.to_excel('./询比价流程详细信息表.xlsx', index=False)

    def run(self):

        url = f'http://portal.chemchina.com/oa08/dept509/{self.IP}.nsf/vwAll?ReadViewEntries&start=1&count=20'
        flag = self.GetPRUidList(url)
        time.sleep(1)
        page = 1
        while flag:  # 当uid_list 不为20时候跳出循环
            # 从第二页开始，需要获取page对应的start
            page_url = f'http://portal.chemchina.com/oa08/flow/homepage.nsf/agtFixViewPosition?openagent&db=dept509/{self.IP}.nsf&vw=vwAll&cat=&page={page}&count=20'
            page_num = self.GetRequest(page_url).text.replace('\n', '')
            url3 = f'http://portal.chemchina.com/oa08/dept509/{self.IP}.nsf/vwAll?ReadViewEntries&start={page_num}&count=20'
            flag = self.GetPRUidList(url3)
            page = page + 1
        self.GetPRData()
        self.save()


# 读取指定文件
def ReadTxT():
    with open("//10.56.164.228/KPI/员工手册.txt") as f:
        txt = f.read()
        f.close()
    return txt


if __name__ == '__main__':

    getOA = GetOAFunc()
    getOA.run()
