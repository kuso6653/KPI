import Func
import pandas as pd
from datetime import datetime, date, timedelta
import re

df = pd.DataFrame(columns=['行'])
df = df.append({'行': '陈堃1'}, ignore_index=True)
df = df.append({'行': '陈堃2'}, ignore_index=True)
df = df.append({'行': '赖大铖'}, ignore_index=True)
df = df.append({'行': '赖大铖2'}, ignore_index=True)
df = df.append({'行': 'ldc'}, ignore_index=True)
df = df[df['行'].str.contains('陈堃|赖大铖')]
print(df)
