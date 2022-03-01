import os
import calendar
from datetime import timedelta
import datetime


# 创建文件夹路径
def mkdir(path):
    folder = os.path.exists(path)
    if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)  # makedirs 创建文件时如果路径不存在会创建这个路径
    else:
        pass


# "--------------------------------------------------------------------------------------------"
# OA 截取字段匹配判断逻辑
def GeneralOffice(approval):  # 新旧oa 的部门管理签字盖章修改
    if approval.find("提交") != -1 and (approval.find("综合管理部盖章") != -1 or approval.find("人事行政部盖章") != -1):
        return True
    else:
        return False


def OrderRegistration(approval):
    if approval.find("提交") != -1 and approval.find("U8销售订单号登记") != -1:
        return True
    else:
        return False


def RelevantPersonnel(approval):
    if approval.find("提交") != -1 and approval.find("相关人员办理") != -1:
        return True
    else:
        return False


# "--------------------------------------------------------------------------------------------"
# 读取指定文件
def ReadTxT():
    with open("//10.56.164.228/KPI/员工手册.txt") as f:
        txt = f.read()
        f.close()
    return txt


# "--------------------------------------------------------------------------------------------"
# 添加日期小于10时不全01、 02 等
def ReformDays(Days):
    now_work_days = []
    for day in Days:
        if day < 10:
            now_work_days.append("0" + str(day))
        else:
            now_work_days.append(str(day))
    return now_work_days


# 获取当月工作日函数
def WorkDays(year, month):
    # 利用日历函数，创建截取工作日日期
    cal = calendar.Calendar()
    WorkDay = []  # 创建工作日数组
    for week in cal.monthdayscalendar(int(year), int(month)):
        for i, day in enumerate(week):
            # 为0或者大于等于5的为休息日
            if day == 0 or i >= 5:
                continue
            # 否则加入数组
            WorkDay.append(day)
    return WorkDay


# 获取当月每天日期函数
def EveryDays(year, month):
    # 利用日历函数，创建截取工作日日期
    cal = calendar.Calendar()
    WorkDay = []  # 创建工作日数组
    for week in cal.monthdayscalendar(int(year), int(month)):
        for i, day in enumerate(week):
            WorkDay.append(day)
    return WorkDay


# 获取当月、上月日期
def GetDate():
    now = datetime.date.today()
    # 获取当月首尾日期
    ThisMonthStart = datetime.datetime(now.year, now.month, 1)
    ThisMonthEnd = datetime.datetime(now.year, now.month, calendar.monthrange(now.year, now.month)[1])
    # 获取上月首尾日期
    LastMonthEnd = ThisMonthStart - timedelta(days=1)
    LastMonthStart = datetime.datetime(LastMonthEnd.year, LastMonthEnd.month, 1)
    return ThisMonthStart, ThisMonthEnd, LastMonthEnd, LastMonthStart
    # 返回
    # 2022-02-01 00:00:00
    # 2022-02-28 00:00:00
    # 2022-01-31 00:00:00
    # 2022-01-01 00:00:00


def str2sec(x):
    # 字符串时分秒转换
    h, m, s = x.strip().split(':')  # .split()函数将其通过':'分隔开，.strip()函数用来除去空格
    return int(h) + int(m) / 60  # int()函数转换成整数运算


# "--------------------------------------------------------------------------------------------"
def Path():
    return "//10.56.164.228/KPI"


# "--------------------------------------------------------------------------------------------"
# U8SQL查询语句

def ReturnInventorySql():  # 返回存货档案查询SQL
    ThisMonthStart, ThisMonthEnd, LastMonthEnd, LastMonthStart = GetDate()
    InventorySql = f"""select Inventory.cinvcode,
      Inventory.cinvdefine7,
      Vendor.cvenname,
      Person.cPersonName,
      Department.cdepname,
      Inventory.iplandefault,
      Inventory.cplanmethod,
      Inventory.iinvadvance,
      Inventory.ialteradvance,
      Inventory.falterbasenum,
      Inventory.fminsupply,
      Inventory.dsdate,
      Inventory.dedate ,
      (inventory.dmodifydate) as dmodifydate,cidefine4
      from Inventory  Left Join InventoryClass on Inventory.cInvCCode=InventoryClass.cInvCCode  Left Join Vendor on Inventory.cVenCode=Vendor.cVenCode  Left Join Position on Inventory.cPosition=Position.cPosCode  Left Join computationGroup on 
    Inventory.cGroupcode=computationGroup.cGroupCode  Left Join  ComputationUnit on Inventory.cComUnitCode=ComputationUnit.cComUnitCode  Left Join Warehouse on Inventory.cDefWareHouse=Warehouse.cWhCode  Left Join Person   on 
    Inventory.cInvPersonCode=Person.cPersonCode  Left Join Person  as PurPerson on Inventory.cPurPersonCode=PurPerson.cPersonCode  Left Join department   on Inventory.cInvDepCode=department.cDepCode  Left join qmCheckProject on  
    Inventory.iQTMethod=qmCheckProject.Id  Left Join Inventory_Sub on Inventory.cInvCode=Inventory_Sub.cInvSubCode  Left Join Inventory_extradefine on Inventory.cInvCode=Inventory_extradefine.cInvCode  where 
    (1=1) and  1=1   And ((Inventory.dSDate >= N'{ThisMonthStart}') And (Inventory.dSDate <= N'{ThisMonthEnd}')) and Inventory.cInvCode not in (select top 0 cInvCode from Inventory where (1=1) and  1=1   And ((Inventory.dSDate >= N'{ThisMonthStart}') And (Inventory.dSDate <= N'{ThisMonthEnd}'))  
    Order by inventory.cinvcode asc)  Order by inventory.cinvcode asc;"""
    return InventorySql


def ReturnMoRoutingDetailSql():  # 返回工序派工委外订单
    ThisMonthStart, ThisMonthEnd, LastMonthEnd, LastMonthStart = GetDate()
    MoRoutingDetailSql = f"""select distinct mh.MoCode as MoCode,md.SortSeq, ii.cInvCode AS InvCode, rd.OpSeq, case IsNull(rs.Moroutingshiftid,0) when 0 then '' else '*' end as DPgFlag
        from sfc_moroutingdetail rd 
        inner join sfc_morouting rh on rd.moroutingid = rh.moroutingid 
        inner join mom_order mh on mh.MoId = rd.MoId 
        inner join mom_orderdetail as md on md.modid = rh.modid  
        inner join bas_part as ip with (nolock) on md.partid = ip.partid  
        inner join inventory as ii with (nolock) on ii.cinvcode = ip.invcode 
        inner join ComputationUnit as iu with (nolock) on ii.cComunitCode = iu.cComUnitCode  
        left outer join mom_morder as mm with (nolock) on mm.modid = md.modid  
        left outer join Factory as mf with (nolock) on mf.cFactoryCode = md.FactoryCode 
        left outer join ComputationUnit mu with (nolock) on md.AuxUnitCode = mu.cComUnitCode 
        left outer join Department dd with (nolock) on dd.cdepcode = md.MDeptcode  
        left outer join AA_RequirementClass as mdc with (nolock) ON mdc.cRClassCode= md.SoCode and md.SoType = 4 
        left outer join mom_remorder as mr with (nolock) on mr.modid = md.modid  
        left outer join UFSystem..UA_User as uc on uc.cUser_Id = rh.CreateUser 
        left outer join UFSystem..UA_User as um on um.cUser_Id = rh.ModifyUser  
        left outer join (select MoRoutingDId as Moroutingshiftid, SumReportQty = sum(ReportQty),SumCompletedQty  = sum(CompletedQty) from sfc_moroutingshift  group by MoRoutingDId) rs 
        on rs.Moroutingshiftid = rd.moroutingdid 
        left outer join mom_motype as mt on mt.motypeid = md.motypeid  
        left outer join reason rm on rm.creasoncode = md.reasoncode 
        left outer join sfc_workcenter as ml on ml.wcid = mr.wcid  
        inner join sfc_workcenter as wc on rd.WcId =wc.WcId 
        left outer join sfc_operation oo on rd.Operationid = oo.Operationid  
        left outer join Vendor as vv on vv.cVenCode = rd.SVendorCode 
        left outer join ComputationUnit as ru on rd.AuxUnitCode = ru.cComunitCode  
        left outer join SO_SOMain as sm on md.OrderType = 1 and md.OrderCode = sm.cSoCode 
        left outer join ex_order as em on md.OrderType = 3 and md.OrderCode = em.cCode  
        LEFT OUTER JOIN Customer as cm ON cm.cCusCode= coalesce(sm.cCusCode,em.cCusCode)  
        where md.SfcFlag = 1 and  1=1   And ((rd.StartDate >= N'{LastMonthStart}') And (rd.StartDate <= N'{ThisMonthEnd}')) And  (md.MoClass = 1 And IsNull(mm.MoDId,0) > 0) And md.status = 3 
        order by mh.MoCode, md.SortSeq,rd.OpSeq;"""
    return MoRoutingDetailSql


def ReturnMRPSql():  # 返回U8数据库查询SQL 的count 和具体的查询数据
    CountSql1 = """
                if object_id('tempdb..##NetDemandPagesData132895399245185579_5d47b5994d584bb8a6249248256aaac1') is null  
        BEGIN CREATE TABLE ##NetDemandPagesData132895399245185579_5d47b5994d584bb8a6249248256aaac1(SortNo int IDENTITY(1,1) NOT NULL,
        TargetId int NOT NULL,  CONSTRAINT PK_##NetDemandPagesData132895399245185579_5d47b5994d584bb8a6249248256aaac1 PRIMARY KEY CLUSTERED (SortNo ASC))  
        Create index idx_1 on ##NetDemandPagesData132895399245185579_5d47b5994d584bb8a6249248256aaac1(TargetId) END

        
    """
    CountSql2 = """
    truncate table ##NetDemandPagesData132895399245185579_5d47b5994d584bb8a6249248256aaac1 Insert into ##NetDemandPagesData132895399245185579_5d47b5994d584bb8a6249248256aaac1 (TargetId) 
        Select n.DemandId  from mps_netdemand n 
        inner join bas_part p on n.PartId = p.PartId 
        Left Outer join mps_planproject j on n.ProjectId=j.ProjectId 
        inner join v_bas_inventory v on p.InvCode = v.InvCode 
        inner join Inventory i on v.InvCode=i.cInvCode 
        Left Outer join mps_plancode c on j.PlanCodeId=c.PlanCodeId 
        Left Outer join mps_netdemandbak nbak on n.DemandId=nbak.DemandId 
        Left Outer join AA_RequirementClass rc on rc.cRClassCode=n.DemandCode 
        left outer join person psp on i.cPurPersonCode=psp.cPersonCode 
        left outer join person psi on i.cInvPersonCode=psi.cPersonCode 
        left outer join person pcu on n.CloseUser=pcu.cPersonCode  Where n.delflag = 0 
        And  ((coalesce(n.SupplyingRCode,'') = '') Or (n.PlanCode=n.SupplyingRCode 
        And n.PlanCode=n.SupplyingPCode))  And  1=1  And n.Status in (1,2,4) and c.PlanCode= N'1'
        and ( n.SupplyType In (0,1,2,3,7))  Order by n.PlanCode  Select Count(SortNo) From ##NetDemandPagesData132895399245185579_5d47b5994d584bb8a6249248256aaac1
    """
    MRPDataSql = """
    
exec sp_executesql N'exec Usp_MP_MP04005_data @StartNo, @EndNo,##NetDemandPagesData132895399245185579_5d47b5994d584bb8a6249248256aaac1',N'@StartNo int,@EndNo int',@StartNo=1,@EndNo=1921"""
    return CountSql1, CountSql2, MRPDataSql
