import PROD_工序派工及时率
import PROD_报工及时率和完整率
import PROD_计划完成率
import QM_产成品检验及时率
import QM_来料检合格率
import QM_来料检验及时率
import QM_产成品检验合格率
import QM_跨车间工序检验及时率
import SCM_LOGISTIC_发货及时率
import SCM_OM_生产订单维护及时率
import SCM_OM_生产订单创建及时率
import SCM_OM_自制物料维护及时率

import SCM_OP_OA_询比价详细信息
import SCM_OP_OA_采购申请详细信息表
import SCM_OP_OA_非生产性物料转换及时率
import SD_CS_OA_销售订单下达及时率

import SCM_OP_准时到货率
import SCM_OP_采购订单转换及时率
import SCM_SP_采购物料维护及时率
import SCM_WM_仓库出入库及时率
import WorkHour

if __name__ == '__main__':
    try:
        PD = PROD_工序派工及时率.ProcessDispatch()
        PD.run()
    except Exception as E:
        print("PROD_工序派工及时率 not run!")

    try:
        WR = PROD_报工及时率和完整率.WorkReport()
        WR.run()
    except Exception as E:
        print("PROD_报工及时率和完整率 not run!")

    try:
        P = PROD_计划完成率.Plan()
        P.run()
    except Exception as E:
        print("PROD_计划完成率 not run!")

    try:
        FP = QM_产成品检验及时率.FinishedProduct()
        FP.run()
    except Exception as E:
        print("QM_产成品检验及时率 not run!")
    try:
        PQC = QM_产成品检验合格率.ProductQualityControl()
        PQC.run()
    except Exception as E:
        print("QM_产成品检验合格率 not run!")

    try:
        QC = QM_来料检合格率.QualityControl()
        QC.run()
    except Exception as E:
        print("QM_来料检合格率 not run!")
    try:
        MI = QM_来料检验及时率.MaterialInspection()
        MI.run()
    except Exception as E:
        print("QM_来料检验及时率 not run!")
    try:
        CW = QM_跨车间工序检验及时率.CrossWorkshop()
        CW.run()
    except Exception as E:
        print("QM_跨车间工序检验及时率 not run!")
    try:
        D = SCM_LOGISTIC_发货及时率.Deliver()
        D.run()
    except Exception as E:
        print("SCM_LOGISTIC_发货及时率 not run!")

    try:
        OC = SCM_OM_生产订单创建及时率.OrderCreation()
        OC.run()
    except Exception as E:
        print("SCM_OM_生产订单创建及时率 not run!")
    try:
        OM = SCM_OM_生产订单维护及时率.OrderMaintenance()
        OM.run()
    except Exception as E:
        print("SCM_OM_生产订单维护及时率 not run!")

    try:
        SM = SCM_OM_自制物料维护及时率.SelfMaterial()
        SM.run()
    except Exception as E:
        print("SCM_OM_自制物料维护及时率 not run!")
    try:
        SM = SCM_OM_自制物料维护及时率.SelfMaterial()
        SM.run()
    except Exception as E:
        print("SCM_OM_自制物料维护及时率 not run!")
    try:
        AT = SCM_OP_准时到货率.ArriveTime()
        AT.run()
    except Exception as E:
        print("SCM_OP_准时到货率 not run!")
    try:
        OC = SCM_OP_采购订单转换及时率.OrderConversion()
        OC.run()
    except Exception as E:
        print("SCM_OP_采购订单转换及时率 not run!")
    try:
        MM = SCM_SP_采购物料维护及时率.MaterialMaintenance()
        MM.run()
    except Exception as E:
        print("SCM_SP_采购物料维护及时率 not run!")
    try:
        W = SCM_WM_仓库出入库及时率.Warehouse()
        W.run()
    except Exception as E:
        print("SCM_WM_仓库出入库及时率 not run!")
    try:
        WH = WorkHour.WorkHour()
        WH.run()
    except Exception as E:
        print("WorkHour not run!")
    try:
        getOA = SCM_OP_OA_非生产性物料转换及时率.GetOAFunc()
        getOA.run()
    except Exception as E:
        print("SCM_OP_OA_非生产性物料转换及时率 not run!")

    try:
        getOA = SCM_OP_OA_非生产性物料转换及时率.GetOAFunc()
        getOA.run()
    except Exception as E:
        print("SCM_OP_OA_非生产性物料转换及时率 not run!")

    try:
        getOA = SCM_OP_OA_采购申请详细信息表.GetOAFunc()
        getOA.run()
    except Exception as E:
        print("SCM_OP_OA_非生产性物料转换及时率 not run!")
    try:
        getOA = SCM_OP_OA_询比价详细信息.GetOAFunc()
        getOA.run()
    except Exception as E:
        print("SCM_OP_OA_询比价详细信息 not run!")
    try:
        getOA = SD_CS_OA_销售订单下达及时率.GetOAFunc()
        getOA.run()
    except Exception as E:
        print("SD_CS_OA_销售订单下达及时率 not run!")


