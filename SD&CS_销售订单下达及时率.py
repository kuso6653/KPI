import re
import time
from numpy import datetime64
from selenium import webdriver
import pandas as pd
import Func

row = 1

title_list = ['U8销售合同单号']


class SalesOrder:
    def __init__(self):
        self.func = Func
        self.txt = self.func.ReadTxT()
        self.path = Func.Path()

        self.driver = webdriver.Chrome()  # 创建对象，启动谷歌浏览器
        js = 'window.open("");'  # 通过执行js，开启一个新的窗口
        self.driver.execute_script(js)
        self.driver.implicitly_wait(20)
        self.n = self.driver.window_handles  # 获取窗口句柄
        self.driver.switch_to.window(self.n[0])  # 切换至第一个窗口
        self.url_list = []
        self.link_name_list = []
        self.SalesOrderData = pd.DataFrame(columns=('U8销售合同单号', 'OA合同审批时间', 'U8系统创建时间'))

    def mkdir(self, path):
        self.func.mkdir(path)

    def first_step(self):

        self.driver.get("http://portal.chemchina.com/")  # 请求url
        self.driver.find_element_by_name('username').send_keys("fjthadmin")  # 输入账号密码
        self.driver.find_element_by_name('password').send_keys(self.txt)

        self.driver.implicitly_wait(20)
        self.driver.find_element_by_xpath('//*[@id="warp"]/form/div[2]/div[3]/div[4]/input[3]').click()  # 点击登录
        time.sleep(10)
        self.driver.implicitly_wait(20)
        self.driver.get(
            "http://portal.chemchina.com/oa08/flow/homepage.nsf/ViewData?readform&draft=no&db=dept509/ContractSale.nsf&vw=vwAll")
        time.sleep(10)
        self.driver.implicitly_wait(20)

    def get_url(self):
        html = self.driver.find_element_by_xpath(
            '/html/body/div[2]/div[2]/div/div/div[2]/div[1]/div[2]/div[3]/table/tbody').get_attribute("outerHTML")
        # 获取url的拼接code
        self.url_list.clear()
        url_code = re.findall(r'<tr unid="(.*?)">', html)
        for i in url_code:
            self.url_list.append(i)

    def get_data(self):
        TimeList = []
        self.driver.switch_to.window(self.n[1])  # 切至第二个窗口
        time.sleep(2)
        for u in self.url_list:
            u = "http://portal.chemchina.com/oa08/dept509/ContractSale.nsf/vwAll/" + u + "?opendocument"
            self.driver.get(u)
            self.driver.implicitly_wait(20)  # 隐式等待，20秒内未找到元素则超时
            time.sleep(3)
            title_html = self.driver.find_element_by_xpath(  # 基本信息
                '//*[@id="BasicInfo"]/div[1]').get_attribute("outerHTML")
            try:
                link = re.findall(r'</button>(.*?)</div>', title_html)[0]  # 当前流程
            except:
                self.driver.get(u)
                self.driver.implicitly_wait(20)  # 隐式等待，20秒内未找到元素则超时
                time.sleep(3)
                title_html = self.driver.find_element_by_xpath(  # 基本信息
                    '//*[@id="BasicInfo"]/div[1]').get_attribute("outerHTML")
                link = re.findall(r'</button>(.*?)</div>', title_html)[0]  # 当前流程
            is_end = str(link).split(" ")[0].split("：")[1]  # 截取当前环节

            information_html = self.driver.find_element_by_xpath(  # 基本信息
                '//*[@id="BasicInfo"]/div[4]').get_attribute("outerHTML")
            # information_html为获取到的基本信息里头的数据，因为基本信息、附件、流转意见等等都在同一个界面展示，所以information_html为基本信息的html
            code = re.findall(
                r'<input name="FormID" value="(.*?)"', information_html)  # 申请单号
            U8_code = re.findall(
                r'<input name="fldU8HeTongBH" value="(.*?)"',  # U8销售合同单号
                information_html)

            if len(code[0]) == 0:  # 单号为空
                continue
            elif int(code[0][:6]) <= 202107:
                return True
            try:
                wander_html = self.driver.find_element_by_xpath('//*[@id="DFlow_MindList"]').get_attribute("outerHTML")
            except:
                self.driver.get(u)
                self.driver.implicitly_wait(20)  # 隐式等待，20秒内未找到元素则超时
                time.sleep(3)
                wander_html = self.driver.find_element_by_xpath('//*[@id="DFlow_MindList"]').get_attribute("outerHTML")
                print("有进来")
            if U8_code[0] != "" and is_end == "结束":
                name_list = re.findall(r'<div class="col-md-4"><b>审批人：</b>(.*?)</div>', wander_html)  # 审批人
                time_list = re.findall(r'<div class="col-md-4"><b>时间：</b>(.*?)</div>', wander_html)  # 时间
                opinion_list = re.findall(r'<pre class="prettyprint linenums prettyprinted">(.*?)</pre>',
                                          wander_html)  # 审批意见
                approval_list = re.findall(r"""<div class="col-md-4"><b>审批节点：</b>(.*?)</div>""", wander_html)  # 审批节点
                for i, j, k, n in zip(approval_list, name_list, time_list, opinion_list):
                    if self.func.GeneralOffice(i):
                        TimeList.append(k)   # 获取 综合管理部盖章 的时间
                    if self.func.OrderRegistration(i):
                        TimeList.append(k)
            if len(TimeList) == 2:
                self.SalesOrderData = self.SalesOrderData.append(
                    {'U8销售合同单号': U8_code[0], 'OA合同审批时间': TimeList[0], 'U8系统创建时间': TimeList[1]}, ignore_index=True)
            TimeList.clear()
        time.sleep(1)
        self.driver.switch_to.window(self.n[0])  # 切换至第一个窗口
        self.driver.find_element_by_class_name("next").click()  # 点击下一页
        self.driver.implicitly_wait(20)
        time.sleep(1)
        return False

    def DataFor(self):
        if len(self.SalesOrderData) != 0:
            self.SalesOrderData["OA合同审批时间"] = self.SalesOrderData["OA合同审批时间"].astype(datetime64)
            self.SalesOrderData["U8系统创建时间"] = self.SalesOrderData['U8系统创建时间'].astype(datetime64)

            self.SalesOrderData['审批延时/H'] = (
                    (self.SalesOrderData['U8系统创建时间'] - self.SalesOrderData['OA合同审批时间']) / pd.Timedelta(1, 'H')).astype(
                int)
            self.SalesOrderData.loc[self.SalesOrderData["审批延时/H"] > 48, "下达及时率"] = "超时"  # 计算出来的审批延时大于3天为超时
            self.SalesOrderData.loc[self.SalesOrderData["审批延时/H"] <= 48, "下达及时率"] = "正常"  # 小于等于3天为正常

    def SaveFile(self):
        path = f"{self.path}/RESULT/SDCS"
        self.mkdir(path)
        file_path = path + '/' + '销售订单下达及时率' + '.xlsx'
        self.SalesOrderData.to_excel(file_path, index=False)
        print("success")

    def run(self):
        self.first_step()
        while 1:
            # 主动暂停查找元素，稳定等待
            self.get_url()
            flag = self.get_data()
            if flag:
                break
        self.DataFor()
        self.SaveFile()
        self.driver.quit()


if __name__ == '__main__':
    oa_get = SalesOrder()
    oa_get.run()
