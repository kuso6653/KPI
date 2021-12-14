from OracleHelper import OracleHelper
import pandas as pd

def data_gaizao(data):
    global dict_username
    # 对数据进行清洗
    # 导入字典，替换人名
    data.loc[:, "USER_NAME"] = data["USER_NAME"].replace(dict_username)
    return data


if __name__ == '__main__':

    # 链接数据库
    SqlOracle = OracleHelper("T5_ENTITY", "thsoft", "10.56.164.22:1521/THPLM")

    sql_create = "select CODE, NAME from TN_A_PROJECT t"
    sql_data = pd.DataFrame(SqlOracle.find_sql(sql_create),columns=['code', 'name'])
    dict_username = dict(zip(sql_data['code'], sql_data['name']))
    print(sql_data)

