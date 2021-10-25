import pandas as pd  #######使用内置统计方法聚合数据
import numpy as np

df = pd.DataFrame({'key1': ['A', 'A', 'B', 'B', 'A']
                      , 'key2': ['one', 'two', 'one', 'two', 'one']
                      , 'data1': [2, 3, 4, 6, 8]
                      , 'data2': [3, 5, np.nan, 3, 7]})
print(df)
df = df.groupby(["key1", "key2"])['data1'].sum()
print(df)