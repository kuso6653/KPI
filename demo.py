# import Func
# import pandas as pd
# from datetime import datetime, date, timedelta
# import re
#
# df = pd.DataFrame(columns=['行'])
# df = df.append({'行': '陈堃1'}, ignore_index=True)
# df = df.append({'行': '陈堃2'}, ignore_index=True)
# df = df.append({'行': '赖大铖'}, ignore_index=True)
# df = df.append({'行': '赖大铖2'}, ignore_index=True)
# df = df.append({'行': 'ldc'}, ignore_index=True)
# df = df[df['行'].str.contains('陈堃|赖大铖')]
# print(df)

def insert_Table_basics(data_list):
    conn = psycopg2.connect(database="aigdb", user="wlaigadmin", password="wlaigadmin", host="localhost", port="5432")
    cursor = conn.cursor()
    for route in data_list:
        route_str = "insert into basics(ip,production_line_number,username,password) values('" + str(route[0]) + "','" + \
                    route[1] + "','" + route[2] + "','" + route[3] + "'" + ")"

        cursor.execute(route_str)

    conn.commit()
    conn.close()

