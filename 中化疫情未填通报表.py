import pandas as pd
import os


def data_check(ExcelName, FlagB, AllData):  # 传入表格的日期， 标号flag用于记录是否是最后一个值， 花名册中所有人员数据
    if FlagB:
        ri = ExcelName[9:11] + ','
    else:
        ri = ExcelName[9:11]

    Check_name_data = pd.read_excel(f'./yq/{ExcelName}', usecols=[5], names=['姓名'], header=3)  # 疫情填报的数据
    Check_name_data = AllData[~AllData['姓名'].isin(Check_name_data['姓名'])]  # 筛选取反，输出疫情未填报的人员名单
    Check_name_data['日期'] = ri  # 加上未填写的日期
    return Check_name_data


if __name__ == '__main__':
    folderName = './yq/'
    excelList = []
    DirList = os.listdir(folderName)
    for DirName in DirList:
        if DirName.startswith('每日疫情'):
            excelList.append(DirName)
    excelList.sort()  # 对日期进行排序
    AllData = pd.read_excel('./yq/员工花名册.xlsx', usecols=['姓名', '部门', '子部门'])
    merge_data = pd.DataFrame(columns=['姓名', '部门', '子部门', '日期'])
    flag = True
    for index, ExcelName in enumerate(excelList):
        if index == len(excelList) - 1:
            flag = False
        check_data = data_check(ExcelName, flag, AllData)
        merge_data = pd.concat(merge_data, check_data, keys=['部门', '子部门','姓名'])  #  未填写人员名单合并
        merge_data = merge_data.fillna('')  # 填补nan数据为空
        merge_data['日期'] = merge_data['日期_x'] + merge_data['日期_y']  # 合并两个日期，并删除多余的xy
        merge_data = merge_data.drop(labels=['日期_y', '日期_x'], axis=1)
    merge_data = merge_data.sort_values(by=['部门', '子部门'])
    # print(merge_data)
    DepList = []
    text = ''
#     text = """
#     关于全厂疫情填报情况的通报
#
# 各部门：
# 为了加强疫情管控力度，公司对全厂疫情填报情况进行抽查，发现有以下部门人员未按公司要求进行填报，现给予通报批评如下：
#     """
    for index, row in merge_data.iterrows():
        if row['子部门'] not in ['总装车间', '机加工车间', '铆焊车间']:
            if row['部门'] not in DepList:
                text = text + '\n' + row['部门'] + ':'
                DepList.append(row['部门'])
            text = text + row['姓名'] + '(' + row['日期'] + ')' + '、'
        else:
            if row['子部门'] not in DepList:
                text = text + '\n' + row['子部门']+ ':'
                DepList.append(row['子部门'])
            text = text + row['姓名'] + '(' + row['日期'] + ')' + '、'
    text = text + '\n'
#     text = text + """
#     请各部门负责人教育好本部门员工，严格遵守公司疫情管控制度。
# 特此通报。
#
#                        福建天华智能装备有限公司
#     """
    with open('./yq/demo.txt', encoding='utf-8', mode='w') as f:
        f.write(text)
    print(text)




