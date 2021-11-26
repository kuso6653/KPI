from decimal import Decimal

import openpyxl
import pandas as pd
from numpy import datetime64
from openpyxl import *
data = pd.DataFrame()
wb = Workbook()
data.to_excel('./demo.xlsx', sheet_name="0", index=False)

SaveBook = load_workbook('./demo.xlsx')
writer = pd.ExcelWriter("./demo.xlsx", engine='openpyxl')
writer.book = SaveBook
data.to_excel(writer, f"name", index=False)
writer.save()

