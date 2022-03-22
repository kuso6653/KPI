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


/******MRP计划维护(新加)**********/
ALTER procedure [dbo].[Usp_MP_MP04005_data]
	@v_StartId int,
	@v_EndId int,
	@v_TableName nvarchar(200)

as

	declare @l_s nvarchar(500), @l_projectid int
	
	create table #tmp_data(SortNo int, TargetId int, DemandId int, PlanCode nvarchar(30),Projectid int)
		
	select @l_s = 'insert into #tmp_data select t.SortNo, t.TargetId, DemandId = n.CopyDemandId, PlanCode = n.PlanCode,n.ProjectId
	  from ' + @v_tablename + ' t inner join mps_netdemand n on t.TargetId = n.DemandId 
	 where n.DelFlag = 0 and t.SortNo between ' + convert(nvarchar(20),@v_StartId) + ' and ' + convert(nvarchar(20),@v_EndId )
		
	exec sp_executesql @l_s
	
	select @l_projectid = Projectid from #tmp_data

	select d.SortNo,d.TargetId, OrderType = DemandId,PlanCode = PlanCode, OrderCode = @l_s,
			 CustCode = @l_s,CustName = @l_s,
			 SoInvCode = @l_s,SoInvName = @l_s, SoInvSpec = @l_s
	  into #tmp_id
	  from #tmp_data d where 1 = 2
	  
    create index #tmp_id1 on #tmp_id(TargetId,SortNo)
    
	exec Usp_MP_GetNetdemandReleaseOrder
	
	update #tmp_id
	set CustCode = isnull(s.cCusCode,e.cCusCode),
		 SoInvCode = isnull(s1.cInvCode,e1.cInvCode)
  from #tmp_id t 
		 inner join mps_netdemand n on t.TargetId = n.DemandId
		 left outer join so_somain s on n.SoType in (1,5) and n.SoCode = s.cSoCode
		 left outer join ex_order e on n.SoType in (3,6) and n.SoCode = e.cCode
		 left outer join so_sodetails s1 on n.SoType = 1 and n.SoDId = convert(nvarchar(10),s1.isosid)
		 left outer join ex_orderdetail e1 on n.SoType = 3 and n.SoDId = convert(nvarchar(10),e1.autoid)
	where n.SoType in (1,3,5,6) 
 
 	update #tmp_id set CustName = c.cCusName,SoInvName = i.cInvName, SoInvSpec = i.cInvStd
	  from #tmp_id t inner join customer c on t.CustCode = c.cCusCode
	  			 left join inventory i on i.cInvCode = t.SoInvCode
	  			 
	
	select distinct n.FactoryCode,n.PartId,p.InvCode,p.MinQty,p.MulQty,p.FixQty,p.SafeQty into #t_factorypart
	  from #tmp_id dlf inner join mps_netdemand n on n.DemandId = dlf.TargetId
			 inner join bas_part p on n.PartId = p.PartId
			 
	if dbo.MultiFactoryEnable() = 1
		update #t_factorypart
			set MinQty = p.MinQty, MulQty = p.MulQty,FixQty = p.FixQty, SafeQty = p.SafeQty
		  from #t_factorypart f inner join bas_factorypart p on f.InvCode = p.InvCode and f.FactoryCode = p.FactoryCode
	
	update #t_factorypart
	   set SafeQty = s.DocQty * -1
	  from #t_factorypart fp inner join mps_schedule s on s.ProjectId = @l_projectId and fp.PartId = s.PartID and docType = 22 
	
	Select n.DemandId,n.PartId,coalesce(n.SoType,0) As SoType,coalesce(n.SoDId,'') As SoDId,
			 coalesce(n.SoId,0) As SoId,n.SoCode,n.SoSeq,n.DemandCode,rc.cRClassName As DemandCodeDesc,
			 n.PlanCode,n.DueDate,n.StartDate,n.LUSD,n.LUCD,n.PlanQty,n.CrdQty,n.SupplyType,n.SchId,n.Ufts,
			 n.ProcQty, n.ManualFlag,n.DelFlag,n.ModifyFlag,n.ProjectId,n.FirmDate,n.FirmUser,n.Status,
			 n.SrpSoDId,n.SrpSoType,convert(decimal(22,6),null) As OnHand,convert(decimal(22,6),null) AS OnOrder,
			 convert(decimal(22,6),null) As OnAllocate,n.SupplyingRCode,n.SupplyingPCode, n.Define22, n.Define23, 
			 n.Define24, n.Define25, n.Define26 , n.Define27, n.Define28, n.Define29, n.Define30, n.Define31, n.Define32,
			 n.Define33, n.Define34, n.Define35, n.Define36 , n.Define37,CONVERT(CHAR,Convert(MONEY,n.Ufts),2) as MpsUfts,
			 i.bATOModel as IsAto,n.CopyDemandId,nbak.PlanQty As OriginalPlanQty, 
			 v.InvCCode,v.InvCode,v.InvAddCode,v.InvName,v.InvStd,v.ComUnitCode As InvUnit,v.ComUnitName As InvUnitName,v.IsRem,v.Policy As Police,v.DemandMergeType As TrackStyle, 
			 p.Free1 As InvFree_1,p.Free2 As InvFree_2,p.Free3 As InvFree_3,p.Free4 As InvFree_4,p.Free5 As InvFree_5, p.Free6 As InvFree_6,p.Free7 As InvFree_7,p.Free8 As InvFree_8,p.Free9 As InvFree_9,p.Free10 As InvFree_10,
			 fp.MinQty,fp.MulQty,fp.FixQty,fp.SafeQty, i.cInvPersonCode As EmplCode,psi.cPersonName As EmplName,n.PurEmplCode As OrgPurEmplCode,coalesce(n.PurEmplCode,i.cPurPersonCode) As PurEmplCode,psp.cPersonName As PurEmplName, 
			 v.InvDefine1 As InvDefine_1,v.InvDefine2 As InvDefine_2,v.InvDefine3 As InvDefine_3,v.InvDefine4 As InvDefine_4,v.InvDefine5 As InvDefine_5, v.InvDefine6 As InvDefine_6,v.InvDefine7 As InvDefine_7,v.InvDefine8 As InvDefine_8,
			 v.InvDefine9 As InvDefine_9,v.InvDefine10 As InvDefine_10, v.InvDefine11 As InvDefine_11,v.InvDefine12 As InvDefine_12,v.InvDefine13 As InvDefine_13,v.InvDefine14 As InvDefine_14,v.InvDefine15 As InvDefine_15,v.InvDefine16 As InvDefine_16, 
			 p.cBasEngineerFigNo as BasEngineerFigNo,i.cPlanMethod as PlanMethod ,dlf.OrderType as RelType,dlf.OrderCode as RelCode,dlf.CustCode,dlf.CustName,dlf.SoInvCode,dlf.SoInvName,dlf.SoInvSpec,n.CloseUser As CloseUser,n.CloseDate,n.CloseTime,
			 n.FactoryCode,f.cFactoryName
	  from mps_netdemand n inner join bas_part p on n.PartId = p.PartId inner join #t_factorypart fp on fp.FactoryCode = n.FactoryCode and p.PartId = fp.PartId
			 Left Outer join mps_planproject j on n.ProjectId=j.ProjectId 
			 Left Outer join v_bas_inventory v on p.InvCode = v.InvCode 
			 Left Outer join Inventory i on v.InvCode=i.cInvCode 
			 Left Outer join mps_plancode c on j.PlanCodeId=c.PlanCodeId 
			 Left Outer join mps_netdemandbak nbak on n.DemandId=nbak.DemandId 
			 Left Outer join AA_RequirementClass rc on rc.cRClassCode=n.DemandCode 
			 left outer join person psp on coalesce(n.PurEmplCode,i.cPurPersonCode)=psp.cPersonCode 
			 left outer join person psi on i.cInvPersonCode=psi.cPersonCode 
			 inner join #tmp_id dlf on n.DemandId = dlf.TargetId 
			 left outer join Factory f on n.FactoryCode = f.cFactoryCode
	 Order by dlf.SortNo;

exec sp_executesql N'exec Usp_MP_MP04005_data @StartNo, @EndNo,##NetDemandPagesData132895399245185579_5d47b5994d584bb8a6249248256aaac1',N'@StartNo int,@EndNo int',@StartNo=1,@EndNo=1921"""
    return CountSql1, CountSql2, MRPDataSql
