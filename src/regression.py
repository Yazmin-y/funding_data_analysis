# -*- coding:utf-8 -*-
import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor as VIF
import pandas as pd
import pickle
import os
import numpy as np
import re
from statsmodels.iolib.summary2 import summary_col


def looper(limit, data):
    cols = list(data.select_dtypes(include=[int, float]).drop('发行总份额', axis=1).columns)
    for i in range(len(cols)):
        x = sm.add_constant(data[cols])  # 生成自变量
        y = data['发行总份额']  # 生成因变量
        model = sm.OLS(y, x)  # 生成模型
        res = model.fit()  # 模型拟合
        # print(res.summary())
        pvalues = res.pvalues  # 得到结果中所有P值
        pvalues.drop('const', inplace=True)  # 把const取得
        pmax = max(pvalues)  # 选出最大的P值
        if pmax > limit:
            ind = pvalues.idxmax()  # 找出最大P值的index
            cols.remove(ind)  # 把这个index从cols中删除
            print(cols)
        else:
            return res
    # return res


def do_regression():
    rootPath = './grouped/'
    excelNames = os.listdir(rootPath)
    res = []
    model_names = []
    for file_name in excelNames:
        clean_data = pd.DataFrame(pd.read_excel(rootPath + file_name, header=1))

        # 收益率相关数据空值处理
        return_col_list = ['最近一年平均收益率', '最近一年最高收益率', '最近三年平均收益率', '最近三年最高收益率',
                           '基金公司最近一年收益率', '基金公司最近三年收益率', '任职时长']
        clean_data[return_col_list] = clean_data[return_col_list].copy().fillna(value=0)
        # clean_data[['在管基金总规模(亿元)', '基金管理人资产净值合计（非货）']] = clean_data[['在管基金总规模(亿元)', '基金管理人资产净值合计（非货）']].copy().fillna(value=1)

        clean_data.drop(clean_data[clean_data['发行年份.2018'] == 1].index, inplace=True)
        clean_data.drop(clean_data[clean_data['发行总份额'] == 0].index, inplace=True)

        # 不使用的自变量
        clean_data.drop(['最近一年平均收益率', '最近三年平均收益率', '存在最近一年平均收益率', '存在最近三年平均收益率', '发行首日前日上证综指收盘价',
                         '发行年份.2018', '发行年份.2019', '任职时长', '在任基金数', '管理费率', '浮动管理费.是'], axis=1,
                        inplace=True)
        # clean_data.drop(
        #     ['最近一年最高收益率', '最近三年最高收益率', '存在最近一年最高收益率', '存在最近三年最高收益率', '发行首日前日上证综指30天涨跌幅', '发行首日前日上证综指30天成交量', '发行年份.2018',
        #      '发行年份.2019'], axis=1,
        #     inplace=True)
        # clean_data.to_excel('./tmp/temp_' + file_name)

        # 空值数据清理
        clean_data.dropna(axis=0, how='any', inplace=True)

        # 部分列取log
        log_col_list = ['在管基金总规模(亿元)', '基金管理人资产净值合计（非货）', '发行总份额', '发行首日前日上证综指30天成交量']
        clean_data.loc[:, log_col_list] = np.log(clean_data.loc[:, log_col_list])
        clean_data.rename(columns=lambda d: str(d) + ".log" if d in log_col_list else str(d), inplace=True)
        clean_data.to_excel('./tmp/temp_' + file_name)

        # 导出数据统计描述
        # describe_df = pd.DataFrame(clean_data.drop('发行总份额.log', axis=1).describe().T.round(3))
        # describe_df.to_excel('./describe/best/' + re.sub('.xlsx', '', file_name) + '_describe.xlsx')
        # describe_df.to_excel('./describe/avg/' + re.sub('.xlsx', '', file_name) + '_describe.xlsx')

        # pearson = clean_data.corr('pearson', numeric_only=True)
        # spearman = clean_data.corr('spearman', numeric_only=True)
        # pearson.to_excel('./corr/best/' + re.sub('.xlsx', '', file_name) + '_pearson.xlsx')
        # spearman.to_excel('./corr/best/' + re.sub('.xlsx', '', file_name) + '_spearman.xlsx')
        # pearson.to_excel('./corr/avg/' + re.sub('.xlsx', '', file_name) + '_pearson.xlsx')
        # spearman.to_excel('./corr/avg/' + re.sub('.xlsx', '', file_name) + '_spearman.xlsx')

        # 回归模型
        x = sm.add_constant(
            clean_data.select_dtypes(include=[int, float]).drop('发行总份额.log', axis=1))  # 生成自变量
        y = clean_data['发行总份额.log']  # 生成因变量
        model = sm.OLS(y, x)  # 生成模型
        result = model.fit()  # 模型拟合
        print(file_name, 'fund regression result')
        res.append(result)
        model_names.append(re.sub('.xlsx', '', file_name))
        # print(result.summary())  # 模型描述
        # vif = [VIF(x, i) for i in range(x.shape[1])]
        # vif = pd.Series([VIF(x.values, i) for i in range(x.shape[1])],
        #           index=x.columns)
        # print(vif)
        # 导出模型描述结果
        # results_summary = result.summary()
        # results_as_html = results_summary.tables[1].as_html()
        # df = pd.read_html(results_as_html, header=0, index_col=0)[0].round(3)
        # df.to_excel('./results/best/' + re.sub('.xlsx', '', file_name) + '_result.xlsx')
        # df.to_excel('./results/avg/' + re.sub('.xlsx', '', file_name) + '_result.xlsx')

        # 导出模型
        # filepath = r'models/best/' + re.sub('.xlsx', '', file_name) + '_model.pkl'
        # with open(filepath, 'wb') as f:
        #     pickle.dump(result, f)
    regressor_order = ['托管费率', '封闭运作时间', '认购天数', '在管基金总规模(亿元).log',
                       '在任公司经理年限', '最近一年最高收益率', '最近三年最高收益率', '性别.男', '学历.博士', '学历.硕士', '存在最近一年最高收益率', '存在最近三年最高收益率',
                       '基金管理人资产净值合计（非货）.log', '基金公司最近一年收益率', '基金公司最近三年收益率', '存在公司最近一年收益率', '存在公司最近三年收益率',
                       '托管行.股份制商业银行', '托管行.国有银行', '托管行.其他银行', '发行首日前日上证综指30天涨跌幅', '发行首日前日上证综指30天成交量.log', '发行首日前日跟踪指数30天涨跌幅', '发行年份.2020', '发行年份.2021', '发行年份.2022']
    print(summary_col(res, stars=True, model_names=model_names, float_format='%.3f', regressor_order=regressor_order))
    summary = summary_col(res, stars=True, model_names=model_names, float_format='%.3f', regressor_order=regressor_order).as_html()
    df = pd.read_html(summary, header=0, index_col=0)[0]
    # print(df.head(5))
    df.to_excel('./results/test4.xlsx')


do_regression()

# 导入模型
# rootPath = './models/best/'
# excelNames = os.listdir(rootPath)
# res = []
# model_names = []
# for file_name in excelNames:
#     filepath = r'models/best/' + file_name
#     with open(filepath, 'rb') as f:
#         load_model = pickle.load(f)
#         print(load_model.summary2())
#         res.append(load_model)
#         model_names.append(re.sub('_model.pkl', '', file_name))

# regressor_order = ['管理费率', '托管费率', '浮动管理费.是', '封闭运作时间', '认购天数', '在任基金数', '在管基金总规模(亿元).log',
#                    '在任公司经理年限', '最近一年最高收益率', '最近三年最高收益率', '任职时长', '性别.男', '学历.博士', '学历.硕士', '存在最近一年最高收益率', '存在最近三年最高收益率',
#                    '基金管理人资产净值合计（非货）.log', '基金公司最近一年收益率', '基金公司最近三年收益率', '存在公司最近一年收益率', '存在公司最近三年收益率',
#                    '托管行.股份制商业银行', '托管行.国有银行', '托管行.其他银行', '发行首日前日上证综指30天涨跌幅', '发行首日前日上证综指30天成交量.log', '发行首日前日跟踪指数30天涨跌幅', '发行年份.2020', '发行年份.2021', '发行年份.2022']
# print(summary_col(res, stars=True, model_names=model_names, float_format='%.3f', regressor_order=regressor_order))
# summary = summary_col(res, stars=True, model_names=model_names, float_format='%.3f', regressor_order=regressor_order).as_html()
# df = pd.read_html(summary, header=0, index_col=0)[0]
# print(df.head(5))
# df.to_excel('./results/test.xlsx')
# with open('results/test.txt', 'w') as file:
#     file.write(summary)

# clean_data = pd.DataFrame(pd.read_excel('useful/merged3.xlsx', header=1))
#
# # 收益率相关数据空值处理
# return_col_list = ['最近一年平均收益率', '最近一年最高收益率', '最近三年平均收益率', '最近三年最高收益率',
#                    '基金公司最近一年收益率', '基金公司最近三年收益率', '任职时长']
# clean_data[return_col_list] = clean_data[return_col_list].copy().fillna(value=0)
#
# clean_data.drop(clean_data[clean_data['发行年份.2018'] == 1].index, inplace=True)
# clean_data.drop(clean_data[clean_data['发行总份额'] == 0].index, inplace=True)
#
# # 不使用的自变量
# clean_data.drop(['最近一年平均收益率', '最近三年平均收益率', '存在最近一年平均收益率', '存在最近三年平均收益率', '发行首日前日上证综指收盘价',
#                  '发行年份.2018', '发行年份.2019'], axis=1,
#                 inplace=True)
#
# # 空值数据清理
# clean_data.dropna(axis=0, how='any', inplace=True)
#
# # 部分列取log
# log_col_list = ['在管基金总规模(亿元)', '基金管理人资产净值合计（非货）', '发行总份额', '发行首日前日上证综指30天成交量']
# clean_data.loc[:, log_col_list] = np.log(clean_data.loc[:, log_col_list])
# clean_data.rename(columns=lambda d: str(d) + ".log" if d in log_col_list else str(d), inplace=True)
#
# # 导出数据统计描述
# describe_df = pd.DataFrame(clean_data.describe().T.round(2))
# describe_df.to_excel('./describe/best/all_describe.xlsx')


