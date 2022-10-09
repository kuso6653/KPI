import datetime
import re
import requests
import time
from numpy import datetime64
import pandas as pd
from dateutil.relativedelta import relativedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
import Func

# 销售采购订单下达及时率
today = datetime.datetime.today()  # datetime类型当前日期
yesterday2 = today - relativedelta(months=2)  # 往时间推迟两个月，取当前和前两个月共三个月数据
yesterday2 = str(yesterday2).split(' ')[0].replace('-', '')[:6]
EarlyTime = int(202201)


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
        self.file_name = "/RESULT/SCM/OP/生产性&非生产性物料采购订单执行效率"
        self.Cookie = cookie
        self.func = Func
        self.path = Func.Path()
        self.PR = 'ContractSale'
        self.headers = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                          'Chrome/71.0.3578.98 Safari/537.36',
            'Cookie': self.Cookie
        }
        self.url_list = []
        self.PR_uid_list = []
        self.SalesOrderDataOA = pd.DataFrame(columns=('U8销售合同单号', '申请人', '申请部门', 'OA合同审批时间'))
        self.SalesOrderDataU8 = pd.read_excel(f"{self.path}/DATA/SDCS/销售订单列表.XLSX", usecols=['订单号', '制单时间'],
                                              converters={'制单时间': datetime64})
        self.SalesOrderDataU8 = self.SalesOrderDataU8.rename(columns={'制单时间': 'U8系统创建时间', '订单号': 'U8销售合同单号'})
        self.SalesOrderDataU8 = self.SalesOrderDataU8.drop_duplicates(subset=["U8销售合同单号"])  # 去重关键列
        self.SalesOrderDataU8 = self.SalesOrderDataU8.dropna(subset=['U8销售合同单号'])  # 去除nan的列

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
        code_data = re.findall(r'<entrydata columnnumber="0" name="AppCreated">\n<text>(.*?)</text>', res.text)
        for i, j in zip(uid, code_data):
            j = j.replace("-", "")[:6]
            try:
                if int(j) >= EarlyTime:
                    self.PR_uid_list.append(i)
                else:
                    flag = False
            except Exception as e:
                print(e)
                continue

        return flag

    def GetPRData(self):
        TimeList = []
        for uid in self.PR_uid_list:
            # while not self.work.empty():
            #     uid = self.work.get_nowait()
            url1 = f'http://portal.chemchina.com/oa08/dept509/{self.PR}.nsf/vwAll/{uid}?opendocument'
            UserInfo = self.GetRequest(url1).content.decode('utf-8')
            # pid用于获取流转意见的request拼接
            pid = re.findall(r'<div class="btn-group btn-corner" data-addType="append" data-ajax="(.*?)">', UserInfo)
            pid = str(pid[0]).split('&')[1]
            # 截取当前流程是否结束
            try:
                is_end = re.findall(r'<input name="TFCurNodeName" type="hidden" value="(.*?)">', UserInfo)[0]
            except Exception as e:
                print(e)
                is_end = ''
            code = re.findall(r'<input name="FormID" value="(.*?)"', UserInfo)[0]  # 申请单号
            Contract_code = re.findall(r'<input name="fldHeTongBH" value="(.*?)"', UserInfo)  # 合同单号

            if str(Contract_code[0]).startswith('XS'):  # 模糊查询以XS开头的合U8单号
                U8_code = str(Contract_code[0]).replace(' ', '')
            else:
                U8_code = re.findall(r'<input name="fldU8HeTongBH" value="(.*?)"', UserInfo)[0]  # U8销售合同单号
            name = re.findall(r'<input name="ApplyPsnCN" value="(.*?)" id="ApplyPsnCN" ', UserInfo)[0]  # 申请人
            dept = re.findall(r'<input name="ApplyDept" value="(.*?)" id="ApplyDept"', UserInfo)[0]  # 申请部门
            if U8_code != "" and is_end == "结束":
                url2 = f'http://portal.chemchina.com/oa08//flow/flowprocess2021.nsf/FlowMind?readform&{pid}'
                WorkTime = self.GetRequest(url2).content.decode('utf-8')
                name_list = re.findall(r'<div class="col-md-4"><b>审批人：</b>(.*?)</div>', WorkTime)  # 审批人
                time_list = re.findall(r'<div class="col-md-4"><b>时间：</b>(.*?)</div>', WorkTime)  # 时间
                opinion_list = re.findall(r'<pre class="prettyprint linenums prettyprinted">(.*?)</pre>',
                                          WorkTime)  # 审批意见
                approval_list = re.findall(r"""<div class="col-md-4"><b>审批节点：</b>(.*?)</div>""", WorkTime)  # 审批节点
                for i, j, k, n in zip(approval_list, name_list, time_list, opinion_list):
                    if self.func.GeneralOffice(i):
                        TimeList.append(k)  # 获取 综合管理部盖章 的时间
                    # if self.func.OrderRegistration(i):
                    #     TimeList.append(k)
            if len(TimeList) == 1:
                self.SalesOrderDataOA = self.SalesOrderDataOA.append(
                    {'U8销售合同单号': U8_code, '申请人': name, '申请部门': dept, 'OA合同审批时间': TimeList[0],
                     }, ignore_index=True)
            TimeList.clear()

    def save(self):
        if len(self.SalesOrderDataOA) != 0:
            self.SalesOrderDataOA = pd.merge(self.SalesOrderDataU8, self.SalesOrderDataOA, on=["U8销售合同单号"])
            self.SalesOrderDataOA["OA合同审批时间"] = self.SalesOrderDataOA["OA合同审批时间"].astype(datetime64)
            self.SalesOrderDataOA["U8系统创建时间"] = self.SalesOrderDataOA['U8系统创建时间'].astype(datetime64)

            self.SalesOrderDataOA['审批延时/H'] = (
                    (self.SalesOrderDataOA['U8系统创建时间'] - self.SalesOrderDataOA['OA合同审批时间'])
                    / pd.Timedelta(1, 'H')).astype(int)
            self.SalesOrderDataOA.loc[self.SalesOrderDataOA["审批延时/H"] > 48, "下达及时率"] = "超时"  # 计算出来的审批延时大于3天为超时
            self.SalesOrderDataOA.loc[self.SalesOrderDataOA["审批延时/H"] <= 48, "下达及时率"] = "正常"  # 小于等于3天为正常
            self.SalesOrderDataOA.to_excel(f"{self.path}/RESULT/SDCS/销售订单下达及时率.xlsx", index=False)

    def run(self):
        url = f'http://portal.chemchina.com/oa08/dept509/{self.PR}.nsf/vwAll?ReadViewEntries&start=1&count=20'
        flag = self.GetPRUidList(url)
        time.sleep(1)
        page = 1
        while flag:  # 当uid_list 不为20时候跳出循环
            # 从第二页开始，需要获取page对应的start
            page_url = f'http://portal.chemchina.com/oa08/flow/homepage.nsf/agtFixViewPosition?openagent&db=dept509/{self.PR}.nsf&vw=vwAll&cat=&page={page}&count=20'
            page_num = self.GetRequest(page_url).text.replace('\n', '')
            url3 = f'http://portal.chemchina.com/oa08/dept509/{self.PR}.nsf/vwAll?ReadViewEntries&start={page_num}&count=20'
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
