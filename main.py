import PROD_工序派工及时率
import PROD_报工及时率和完整率
import PROD_计划完成率

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