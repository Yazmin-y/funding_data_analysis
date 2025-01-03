import xlrd
import os
import pandas as pd
import openpyxl
import re


rootPath = './海通数据2018Q4-2022Q4'
# xls_save_as_xlsx(rootPath)
excelNames = os.listdir(rootPath)  ##获得所有的文件名
# print(excelNames)
df_final = pd.DataFrame(columns=['基金公司', '最近一年收益率[%]', '最近一年收益率排名', '最近三年收益率[%]', '最近三年收益率排名', '收益类型', '计算截止日期'])
for excelName in excelNames:
    print(excelName)
    wb = openpyxl.load_workbook(rootPath + '/' + excelName)
    table = wb.worksheets[0]
    date = table['J6'].value
    date = re.sub('[\u4e00-\u9fa5]', '', date)
    date = re.sub('：', '', date)
    print(date)
    df1 = pd.DataFrame(pd.read_excel(rootPath + '/' + excelName, sheet_name='权益类', skiprows=3, header=[0, 1]))
    df2 = pd.DataFrame(pd.read_excel(rootPath + '/' + excelName, sheet_name='固定收益类', skiprows=3, header=[0, 1]))
    # df1.columns = ["".join(x) for x in df1.columns.ravel()]
    year1_1 = df1.filter(regex='最近一年')
    year3_1 = df1.filter(regex='最近三年')
    df_qy = pd.concat([df1.iloc[:, 0], year1_1, year3_1], axis=1)
    df_qy.columns = ['基金公司', '最近一年收益率[%]', '最近一年收益率排名', '最近三年收益率[%]', '最近三年收益率排名']
    df_qy.loc[:, '收益类型'] = '权益类'
    df_qy.loc[:, '计算截止日期'] = date
    print(df_qy.info())

    year1_2 = df2.filter(regex='最近一年')
    year3_2 = df2.filter(regex='最近三年')
    df_gs = pd.concat([df2.iloc[:, 0], year1_2, year3_2], axis=1)
    df_gs.columns = ['基金公司', '最近一年收益率[%]', '最近一年收益率排名', '最近三年收益率[%]', '最近三年收益率排名']
    df_gs.loc[:, '收益类型'] = '固定收益类'
    df_gs.loc[:, '计算截止日期'] = date

    df_tmp = pd.concat([df_qy, df_gs])

    df_final = pd.concat([df_final, df_tmp])

print(df_final.info())
df_final.to_excel('./company_final.xlsx')
