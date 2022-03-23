import Func
import pandas as pd
from datetime import datetime, date, timedelta
import re

str1 = '物料名称1'
str2 = re.findall(r'[\u4e00-\u9fa5]{4}[0-9]', str1)
print(str2)
