import pandas as pd
import time
import os
import re

rootPath = './基金经理对比大全'
excelNames = os.listdir(rootPath)  ##获得所有的文件名
# print(excelNames)
merged_df = pd.DataFrame()
for excelName in excelNames:
    print(excelName)
    date = excelName
    date = re.sub('[\u4e00-\u9fa5]', '', date)
    date = re.sub('.xlsx', '', date)
    print(date)
    df = pd.DataFrame(pd.read_excel(rootPath + '/' + excelName, skiprows=1))
    df = df.iloc[:, 1:7]
    # print(df.info())
    df.loc[:, '计算截止日期'] = date
    year = int(pd.to_datetime(date).year)
    month = int(pd.to_datetime(date).month)
    df.loc[:, '年份'] = year
    df.loc[:, '季度'] = month
    df['季度'].replace({3: 1, 6: 2, 9: 3, 12: 4}, inplace=True)
    merged_df = pd.concat([merged_df, df], axis=0)
print(merged_df.info())
merged_df.to_excel('./managers.xlsx')