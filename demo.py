import pandas as pd
BaseData = pd.read_excel(f"{path}/DATA/工序派工资料维护{ThisYear}-{LastMonth}-{work_day}.XLSX",
                                             usecols=['物料编码', '生产订单', '工序行号', '派工标识', '行号'],
                                             converters={'物料编码': int, '工序行号': int}
                                             )