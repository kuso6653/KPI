import re
import time
from selenium import webdriver
import pandas as pd
from numpy import datetime64
import Func

title_list = ['U8采购单号']


class PurchaseConversion:
    def __init__(self):
        self.func = Func
        self.txt = self.func.ReadTxT()
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = Func.Path()

        # 将上月首尾日期切割
        self.LastMonthStart = str(self.LastMonthStart).split(" ")[0].replace("-", "")
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0].replace("-", "")
        self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0].replace("-", "")

        self.Purchase_in_data = pd.read_excel(f"{self.path}/DATA/SCM/OP/采购订单列表.XLSX",
                                              usecols=['存货编码', '存货名称', '订单编号', '主计量', '数量', '制单时间'],
                                              converters={'存货编码': str, '订单编号': str, '制单时间': datetime64}
                                              )
        self.Purchase_in_data = self.Purchase_in_data.dropna(subset=['存货编码'])  # 去除nan的列
        self.Purchase_in_data = self.Purchase_in_data.rename(columns={'订单编号': 'U8采购单号'})
        self.one_list = []
        self.driver = webdriver.Chrome()  # 创建对象，启动谷歌浏览器
        js = 'window.open("");'  # 通过执行js，开启一个新的窗口
        self.driver.execute_script(js)
        self.driver.implicitly_wait(20)
        self.n = self.driver.window_handles  # 获取窗口句柄
        self.driver.switch_to.window(self.n[0])  # 切换至第一个窗口

        self.url_list = []
        self.link_name_list = []
        self.PurchaseConversionList = pd.DataFrame(columns=["申请单号", "申请人", "申请日期", "U8采购单号", "审批日期"])
        self.merge_data = pd.DataFrame

    def mkdir(self, path):
        self.func.mkdir(path)

    def first_step(self):
        self.driver.get("http://portal.chemchina.com/")  # 请求url
        self.driver.find_element_by_name('username').send_keys("fjthadmin")  # 输入账号密码
        self.driver.find_element_by_name('password').send_keys(self.txt)

        self.driver.implicitly_wait(20)
        self.driver.find_element_by_xpath('//*[@id="warp"]/form/div[2]/div[3]/div[4]/input[3]').click()  # 点击登录

        self.driver.implicitly_wait(20)
        self.driver.get(
            "http://portal.chemchina.com/oa08/flow/homepage.nsf/ViewData?readform&draft=no&db=dept509/SelectPurchase.nsf&vw=vwAll")
        self.driver.implicitly_wait(20)

    def get_url(self):
        time.sleep(3)
        html = self.driver.find_element_by_xpath(
            '/html/body/div[2]/div[2]/div/div/div[2]/div[1]/div[2]/div[3]/table/tbody').get_attribute("outerHTML")
        # 获取url的拼接code
        self.url_list.clear()
        url_code = re.findall(r'<tr unid="(.*?)">', html)
        for i in url_code:
            self.url_list.append(i)

    def get_data(self):
        self.driver.switch_to.window(self.n[1])  # 切至第二个窗口

        for u in self.url_list:
            u = "http://portal.chemchina.com/oa08/dept509/SelectPurchase.nsf/vwAll/" + u + "?opendocument"
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
            Application_name = re.findall(
                r'<input name="ApplyPsnCN" value="(.*?)" id="ApplyPsnCN" class="form-control" type="text" readonly="readonly">',
                information_html)  # 申请人
            code = re.findall(
                r'<input name="FormID" value="(.*?)"', information_html)  # OA申请单号
            U8_code = re.findall(
                r'<input name="fldU8Nun" value="(.*?)" class="form-control"',  # U8采购单号
                information_html)
            Application_date = re.findall(
                r'<input name="ApplyDate" value="(.*?)" id="ApplyDate" class="form-control" readonly="readonly" type="text">',
                # OA申请日期
                information_html)
            if len(code[0]) == 0:  # 单号为空
                continue
            elif int(code[0][:6]) <= 202109:
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
                    if self.func.RelevantPersonnel(i):
                        # ["申请单号", "申请人", "申请日期", "U8采购单号", "审批日期"
                        self.PurchaseConversionList = self.PurchaseConversionList.append(
                            {"申请单号": code[0], "申请人": Application_name[0],
                             "申请日期": Application_date[0], "U8采购单号": U8_code[0],
                             "审批日期": k},
                            ignore_index=True)
                        break
                self.one_list.clear()
        self.driver.switch_to.window(self.n[0])  # 切换至第一个窗口
        self.driver.find_element_by_class_name("next").click()  # 点击下一页
        self.driver.implicitly_wait(20)
        time.sleep(1)

        return False

    def MergeDataFor(self):
        self.PurchaseConversionList["审批日期"] = self.PurchaseConversionList["审批日期"].astype(datetime64)
        self.merge_data = pd.merge(self.PurchaseConversionList, self.Purchase_in_data, on=["U8采购单号"])
        self.merge_data['未及时率/H'] = (
                (self.merge_data['制单时间'] - self.merge_data['审批日期']) / pd.Timedelta(1, 'H')).astype(
            int)
        self.merge_data.loc[self.merge_data["未及时率/H"] > 384, "创建及时率"] = "超时"  # 计算出来的审批延时大于16天为超时
        self.merge_data.loc[self.merge_data["未及时率/H"] <= 384, "创建及时率"] = "正常"  # 小于等于16天为正常

    def save_data(self):
        path = f"{self.path}/RESULT/SCM/OP"
        self.mkdir(path)
        file_path = path + '/' + '非生产性物料转换及时率' + '.xlsx'
        self.merge_data.to_excel(file_path, index=False)
        print("success")

    def run(self):
        self.first_step()
        while 1:
            # 主动暂停查找元素，稳定等待
            self.get_url()
            flag = self.get_data()
            if flag:
                break
        self.MergeDataFor()
        self.save_data()

        self.driver.quit()


if __name__ == '__main__':
    oa_get = PurchaseConversion()
    oa_get.run()
