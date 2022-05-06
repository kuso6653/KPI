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

monkey.patch_all()
import Func
from dateutil.relativedelta import relativedelta
from get_holiday_cn.client import getHoliday

today = datetime.datetime.today()  # datetime类型当前日期
yesterday2 = today - relativedelta(months=2)  # 往时间推迟两个月，取当前和前两个月共三个月数据
yesterday2 = str(yesterday2).split(' ')[0].replace('-', '')[:6]
EarlyTime = int(yesterday2)


# 	生产性物料采购订单执行效率（是否有效推进订单办结）：
# 	    开始时间：OA采购合同审批开始时间；
# 	    结束时间：OA采购合同审批结束时间；(如果没有结束时间，按脚本执行当天17:00算)
# 	    采购订单效率=结束时间-开始时间；1）


#   PR-采购申请 / PO-采购合同
# 	非生产性物料采购订单执行效率（是否及时下单/是否有效推进订单办结）：
# a)	开始时间：OA 采购申请审批开始时间；
# b)	结束时间：OA采购合同审批结束时间；
# c)	PR->PO 时长 = 采购合同开始 - 采购申请结束
# d)    PO时长 = 采购合同开始时间 - 采购合同结束时间(如果没有结束时间，按脚本执行当天17:00算)

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


class GetWorkHour:
    def __init__(self, StartTime, EndTime):
        self.StartTime = StartTime
        self.EndTime = EndTime

    def StartTimeFunc(self):  # 取年月日进行判断
        YMDTime = self.StartTime
        flag = 0
        while True:
            code, type_code, week = GetHoliday(str(YMDTime).split(" ")[0])  # 判断是否是节假日
            if code == 0 and type_code != 0:
                flag = 1  # 是在节假日进行的审批，打标记然后日期 加1天 ，继续判断直到是工作日为止
                YMDTime = pd.to_datetime(YMDTime) + pd.to_timedelta('24:00:00')
            else:
                break
        if flag == 1:  # 如果是节假日审批，则移动到下个工作日，并清除 时分秒
            YMDTime = pd.to_datetime(YMDTime) - pd.to_timedelta(str(YMDTime).split(" ")[1])
            return YMDTime
        else:
            return YMDTime

    def EndTimeFunc(self):  # 取年月日进行判断，在周末审批则往前移动到
        YMDTime = self.EndTime
        flag = 0
        while True:
            code, type_code, week = GetHoliday(str(YMDTime).split(" ")[0])  # 判断是否是节假日
            if code == 0 and type_code != 0:
                flag = 1  # 是在节假日进行的审批，打标记然后日期 减1天 ，继续判断直到是工作日为止
                YMDTime = pd.to_datetime(YMDTime) - pd.to_timedelta('24:00:00')
            else:
                break
        if flag == 1:  # 如果是节假日审批，则移动到上个工作日，并定位到上个工作日最后时间 添加 时分秒
            YMDTime = pd.to_datetime(str(YMDTime).split(" ")[0]) + pd.to_timedelta('23:59:59')
            return YMDTime
        else:
            return YMDTime

    def GetHolidayDays(self):  # 计算获取需要扣减的周末时间天数
        start = self.StartTime.split(" ")[0]
        end = self.EndTime.split(" ")[0]
        StartCount = 0
        EndCount = 0
        WorkCount = len(pd.bdate_range(start, end))  # 工作日总天数
        ALLCount = len(pd.date_range(start, end))  # 总天数包括周末
        data1, type1, week1 = GetHoliday(start)  # 是周六则减2，周天减1
        data2, type2, week2 = GetHoliday(end)  # 是周六则减1，周天减2
        if week1 == 6:
            StartCount = 2
        elif week1 == 7:
            StartCount = 1
        if week2 == 6:
            EndCount = 1
        elif week2 == 7:
            EndCount = 2
        return ALLCount - WorkCount - StartCount - EndCount

    def run(self):
        StartTime = self.StartTimeFunc()
        EndTime = self.EndTimeFunc()

        d1 = datetime.datetime.strptime(str(EndTime), '%Y-%m-%d %H:%M:%S')
        d2 = datetime.datetime.strptime(str(StartTime), '%Y-%m-%d %H:%M:%S')
        delta = d1 - d2
        HolidayDays = self.GetHolidayDays()
        delta2 = round(delta.days * 24 + delta.seconds / 3600, 2) - HolidayDays * 24  # 计算时间再减去连续跨周末
        return delta2


class GetOAFunc:
    def __init__(self, Cookie):
        self.file_name = "/RESULT/SCM/OP/生产性&非生产性物料采购订单执行效率"
        self.Cookie = Cookie
        self.func = Func
        self.path = Func.Path()
        self.PO = 'ProcureContract'
        self.PR = 'SelectPurchase'
        self.headers = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                          'Chrome/71.0.3578.98 Safari/537.36',
            'Cookie': self.Cookie
        }
        self.PO_Productive = pd.DataFrame(columns=(
            '申请单号', '部门', '经办人', '合同标的物名称', '合同标的额', '开始时间', '结束时间', '流转状态', '流转时长'))

        self.PO_UnProductive = pd.DataFrame(columns=(
            '采购订单号', '采购申请单号', '部门', '经办人', '合同标的物名称', '合同标的额', '采购订单审批状态',
            '采购订单开始时间', '采购订单结束时间', '采购申请转采购订单时长', '采购订单审批时长'))

        self.PR_UnProductive = pd.DataFrame(columns=('采购申请单号', '采购申请结束时间'))
        self.PRToPo = pd.DataFrame
        self.PRToPo2 = pd.DataFrame(columns=(
            '采购订单号', '采购申请单号', '部门', '经办人', '合同标的物名称', '合同标的额', '采购订单审批状态',
            '采购申请结束时间', '采购订单开始时间', '采购订单结束时间', '采购申请转采购订单时长', '采购订单审批时长'))

        now = datetime.date.today()
        self.TodayTime = str(now).split(' ')[0]
        TodayLastTime = pd.to_datetime(str(now)) + pd.to_timedelta('17:00:00')
        self.TodayLastTime = str(TodayLastTime)  # 获取今天下班时间 转为str

        self.work = Queue()  # 创建队列对象，并赋值给work。

    def GetPOUidList(self, url):
        flag = True
        res = self.GetRequest(url)
        uid = re.findall(r'unid="(.*?)"', res.text)
        code = re.findall(r'<entrydata columnnumber="3" name="FormID">\n<text>(.*?)</text>', res.text)
        for i, j in zip(uid, code):
            try:
                if int(j[:6]) >= EarlyTime:
                    # self.PO_uid_list.append(i)
                    self.work.put_nowait(i)
                else:
                    flag = False
            except Exception as e:
                print(e)
                continue
        return flag

    def GetPRUidList(self, url):
        flag = True
        res = self.GetRequest(url)
        uid = re.findall(r'unid="(.*?)"', res.text)
        code = re.findall(r'<entrydata columnnumber="0" name="FormID">\n<text>(.*?)</text>', res.text)
        for i, j in zip(uid, code):
            try:
                if int(j[:6]) >= EarlyTime:
                    # self.PR_uid_list.append(i)
                    self.work.put_nowait(i)
                else:
                    flag = False
            except Exception as e:
                print(e)
                continue

        return flag

    def GetPOData(self):  # 获取采购合同的生产和非生产性物料
        while not self.work.empty():
            uid = self.work.get_nowait()
            url1 = f'http://portal.chemchina.com/oa08/dept509/{self.PO}.nsf/vwAll/{uid}?opendocument'
            UserInfo = self.GetRequest(url1).content.decode('utf-8')
            # 当前流程环节
            is_end = re.findall(r'<input name="TFCurNodeName" type="hidden" value="(.*?)">', UserInfo)[0]
            # pid用于获取流转意见的request拼接
            pid = re.findall(r'<div class="btn-group btn-corner" data-addType="append" data-ajax="(.*?)">',
                             UserInfo)
            pid = str(pid[0]).split('&')[1]
            # 订单编号
            uid = re.findall(r'<input name="FormID" value="(.*?)"', UserInfo)
            # 经办人
            name = re.findall(r'<input name="ApplyPsnCN" value="(.*?)" id="ApplyPsnCN"', UserInfo)
            # 部门
            dept = re.findall(r'<input name="ApplyDept" value="(.*?)" id="ApplyDept"', UserInfo)
            # 标的物金额
            money = re.findall(r'<input name="fldHeTongBDEx" value="(.*?)" ', UserInfo)
            # 标的物名称
            subject = re.findall(r'<input name="fldHeTongName" value="(.*?)"', UserInfo)
            applyNo = re.findall(r'<input name="fldCcsqd" value="(.*?)"', UserInfo)  # 抓取采购申请ID
            product = re.findall(r'<option value="(.*?)" selected', UserInfo)  # 抓取是否是生产性物料

            url2 = f'http://portal.chemchina.com/oa08//flow/flowprocess2021.nsf/FlowMind?readform&{pid}'
            WorkTime = self.GetRequest(url2).content.decode('utf-8')
            time_list = re.findall(r'<div class="col-md-4"><b>时间：</b>(.*?)</div>', WorkTime)
            if '生产性物料采购' in product:
                if is_end == '结束':
                    hour = GetWorkHour(time_list[0], time_list[-1])
                    status = '结束'
                else:
                    hour = GetWorkHour(time_list[0], self.TodayLastTime)
                    status = '流转中'
                delta2 = hour.run()
                self.PO_Productive = self.PO_Productive.append([{'申请单号': uid[0], '部门': dept[0], '经办人': name[0],
                                                                 '合同标的物名称': subject[0], '合同标的额': money[0],
                                                                 '开始时间': time_list[0],
                                                                 '结束时间': self.TodayLastTime,
                                                                 '流转状态': status, '流转时长': delta2
                                                                 }],
                                                               ignore_index=True)
            elif '非生产性物料采购' in product:
                if is_end == '结束':
                    hour = GetWorkHour(time_list[0], time_list[-1])
                    status = '结束'
                else:
                    hour = GetWorkHour(time_list[0], self.TodayLastTime)
                    status = '流转中'
                delta2 = hour.run()
                self.PO_UnProductive = self.PO_UnProductive.append(
                    [{'采购订单号': uid[0], '采购申请单号': applyNo[0], '部门': dept[0], '经办人': name[0],
                      '合同标的物名称': subject[0],
                      '合同标的额': money[0], '采购订单审批状态': status, '采购订单开始时间': time_list[0],
                      '采购订单结束时间': self.TodayLastTime, '采购订单审批时长': delta2}], ignore_index=True)
            else:
                continue

    def GetPRData(self):
        while not self.work.empty():
            uid = self.work.get_nowait()
            url1 = f'http://portal.chemchina.com/oa08/dept509/{self.PR}.nsf/vwAll/{uid}?opendocument'
            UserInfo = self.GetRequest(url1).content.decode('utf-8')
            # pid用于获取流转意见的request拼接
            pid = re.findall(r'<div class="btn-group btn-corner" data-addType="append" data-ajax="(.*?)">',
                             UserInfo)
            pid = str(pid[0]).split('&')[1]
            # 订单编号
            uid = re.findall(r'<input name="FormID" value="(.*?)"', UserInfo)
            url2 = f'http://portal.chemchina.com/oa08//flow/flowprocess2021.nsf/FlowMind?readform&{pid}'
            WorkTime = self.GetRequest(url2).content.decode('utf-8')
            time_list = re.findall(r'<div class="col-md-4"><b>时间：</b>(.*?)</div>', WorkTime)
            self.PR_UnProductive = self.PR_UnProductive.append(
                [{'采购申请单号': uid[0], '采购申请结束时间': time_list[-1]}], ignore_index=True)

    def PRFromPOToTime(self):  # 计算PR转PO的时间，并保存
        self.PRToPo = pd.merge(self.PO_UnProductive, self.PR_UnProductive, on=['采购申请单号'])
        for index, row in self.PRToPo.iterrows():
            hour = GetWorkHour(row['采购申请结束时间'], row['采购订单开始时间'])
            PRtoPoTime = hour.run()
            # row['采购申请转采购订单时长'] = time  # 不知道为什么不能赋值
            self.PRToPo2 = self.PRToPo2.append(
                [{'采购订单号': row[0], '采购申请单号': row[1], '部门': row[2], '经办人': row[3],
                  '合同标的物名称': row[4], '合同标的额': row[5], '采购订单审批状态': row[6],
                  '采购申请结束时间': row[11],
                  '采购订单开始时间': row[7], '采购订单结束时间': row[8], '采购申请转采购订单时长': PRtoPoTime, '采购订单审批时长': row[10]
                  }],
                ignore_index=True)

        self.PRToPo2.to_excel(f'{self.path}{self.file_name}{self.TodayTime}.xlsx', sheet_name="非生产性物料采购效率统计", index=False)

        SaveBook = load_workbook(f'{self.path}{self.file_name}{self.TodayTime}.xlsx')
        writer = pd.ExcelWriter(f'{self.path}{self.file_name}{self.TodayTime}.xlsx', engine='openpyxl')
        writer.book = SaveBook
        self.PO_Productive.to_excel(writer, '生产性物料采购效率统计', index=False)
        writer.save()

    def GetRequest(self, url):
        try:
            res = requests.get(url, headers=self.headers)
        except Exception as e:
            print(e)
            time.sleep(3)
            res = requests.get(url, headers=self.headers)
        return res

    def run(self):

        url = f'http://portal.chemchina.com/oa08/dept509/{self.PO}.nsf/vwAll?ReadViewEntries&start=1&count=20'
        flag = self.GetPOUidList(url)
        time.sleep(1)
        page = 1
        # 第一次获取UID,无需判断page
        while flag:  # 当uid_list 不为20时候跳出循环
            # 从第二页开始，需要获取page对应的start
            page_url = f'http://portal.chemchina.com/oa08/flow/homepage.nsf/agtFixViewPosition?openagent&db=dept509/{self.PO}.nsf&vw=vwAll&cat=&page={page}&count=20 '
            page_num = self.GetRequest(page_url).text.replace('\n', '')
            url3 = f'http://portal.chemchina.com/oa08/dept509/{self.PO}.nsf/vwAll?ReadViewEntries&start={page_num}&count=20 '
            flag = self.GetPOUidList(url3)
            page = page + 1

        task_list = []
        for x in range(4):
            task = gevent.spawn(self.GetPOData)
            task_list.append(task)
        gevent.joinall(task_list)

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
        task_list = []
        for x in range(4):
            task = gevent.spawn(self.GetPRData)
            task_list.append(task)
        gevent.joinall(task_list)

        self.PRFromPOToTime()


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
