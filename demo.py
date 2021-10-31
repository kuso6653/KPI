# -*- coding:utf-8 -*-
import pandas as pd
from pandas import *

df1 = DataFrame([['a', 10, '男'],
                 ['b', 11, '男'],
                 ['c', 11, '女'],
                 ['a', 10, '女'],
                 ['c', 11, '男']],
                columns=['name', 'age', 'sex'])
print("df1:\n%s\n\n" % df1)
df2 = DataFrame([['a', 10, '男'],
                 ['b', 11, '女']],
                columns=['name', 'age', 'sex'])
print("df2:\n%s\n\n" % df2)

df1 = df1.append(df2)
df1 = df1.append(df2)
set_diff_df = df1.drop_duplicates(subset=['name', 'age', 'sex'],keep=False)
print(set_diff_df)