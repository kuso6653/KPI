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
EarlyTime = int(20220609)


class GetOAFunc:
    def __init__(self, Cookie):
        self.file_name = "/RESULT/SCM/OP/出差申请"
        self.Cookie = Cookie
        self.func = Func
        self.path = Func.Path()
        self.IP = 'StaffTravel'
        self.headers = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                          'Chrome/71.0.3578.98 Safari/537.36',
            'Cookie': self.Cookie
        }
        self.PR_uid_list = []
        self.PR_Productive = pd.DataFrame(columns=('经办人', '经办部门', '出差人员', '申请日期', '出差地域'))

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
        if len(uid) >= 20:
            self.PR_uid_list = self.PR_uid_list + uid
            return flag
        else:
            self.PR_uid_list = self.PR_uid_list + uid
            flag = False
            return flag

    def GetPRData(self):
        for uid in self.PR_uid_list:
            # while not self.work.empty():
            #     uid = self.work.get_nowait()
            url1 = f'http://portal.chemchina.com/oa08/dept509/{self.IP}.nsf/vwAll/{uid}?opendocument'
            UserInfo = self.GetRequest(url1).content.decode('utf-8')
            # 订单编号
            OrderUserName = re.findall(r'<input name="ApplyPsnCN" value="(.*?)" ', UserInfo)[0]  # 经办人
            Department = re.findall(r'<input name="ApplyDept" value="(.*?)"', UserInfo)[0]  # 申请部门
            ApplicantUserName = re.findall(r'<input name="fldChuchaiRen" value="(.*?)" readonly type="text">', UserInfo)[0]  # 申请人
            ApplicantDate = re.findall(r'<input name="ApplyDate" value="(.*?)" id="ApplyDate"', UserInfo)[0]  # 申请日期
            Address = re.findall(r'<option value="(.*?)" selected>', UserInfo)[1]  # 出差地域

            self.PR_Productive = self.PR_Productive.append(
                [{'经办人': OrderUserName, '经办部门': Department, '出差人员': ApplicantUserName, '申请日期': ApplicantDate,
                  '出差地域': Address}], ignore_index=True)

            # self.PR_Productive = pd.DataFrame(columns=('经办人', '经办部门', '出差人员', '申请日期', '出差地域'))

    def save(self):
        self.PR_Productive.to_excel('./出差申请.xlsx', index=False)

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
