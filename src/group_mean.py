import pandas as pd
import re
from fuzzywuzzy import fuzz, process

pd.set_option('display.max_columns', None)

# 计算平均和最高收益率
df = pd.DataFrame(pd.read_pickle('./fund_final.pkl'))
df[['最近一年收益率[%]', '最近三年收益率[%]']] = df[['最近一年收益率[%]', '最近三年收益率[%]']].apply(pd.to_numeric, errors='coerce')

# df = df.head(100)
group = df.groupby(['现任基金经理', '所属公司', '计算截止日期'], as_index=False)

count = group['最近一年收益率[%]', '最近三年收益率[%]'].agg(['mean', 'max'])
print(count)

count_df = pd.DataFrame(count)
count_df.reset_index(inplace=True)
count_df.columns = ['现任基金经理', '所属公司', '计算截止日期', '最近一年平均收益率', '最近一年最高收益率', '最近三年平均收益率', '最近三年最高收益率']
print(count_df.info())
count_df.to_excel('./maxAndMean.xlsx')
merged = pd.merge(df, count_df, how='left')
print(merged.info())

merged.to_excel('./fund_final_plus.xlsx')

# 加入基金公司全称
df = pd.DataFrame(pd.read_excel('./maxAndMean.xlsx'))
# df = df.head(1000)
full_names = pd.DataFrame(pd.read_excel('./data_files/company.xlsx', usecols=[0], header=1))
full_names.columns = ['基金公司全称']


def fuzzy_merge(df1, df2, key1, key2, new_col_name='基金公司全称', threshold=85, limit=2):
    s = df2[key2].tolist()
    n = df1[key1].apply(lambda x: process.extract(x, s, limit=limit))
    df1[new_col_name] = n
    n2 = df1[new_col_name].apply(
        lambda x: [i[0] for i in x if i[1] >= threshold][0] if len([i[0] for i in x if i[1] >= threshold]) > 0 else '')
    df1[new_col_name] = n2
    return df1


# 字符串模糊匹配方式
df_merge = fuzzy_merge(df, full_names, '所属公司', '基金公司全称')
# df_merge = df.apply(merge_func, axis=1)
print(df_merge.head(5))
print(df_merge.info())
df_merge.to_excel('./test.xlsx')


# 直接匹配方式
df_all = pd.DataFrame(pd.read_excel('./useful/company.xlsx'))
# df_all = pd.DataFrame(pd.read_pickle('./fund_final.pkl'))
df_name = pd.DataFrame(pd.read_excel('./基金简称全称匹配.xlsx'))
# df_merge = df_merge[df_merge.columns.drop(list(df_merge.filter(regex='Unnamed')))]
# df_all = df_all[df_all.columns.drop(list(df_all.filter(regex='Unnamed')))]
df_all["基金公司"].replace('--', '东兴证券', inplace=True)
print(df_all.info())
df_all = pd.merge(df_all, df_name, left_on='基金公司', right_on='所属公司', how='left')
print(df_all.info())
df_all.to_excel('./useful/company_test.xlsx')

# df_all = pd.merge(df_all, df_merge, on=['现任基金经理', '所属公司', '计算截止日期'], how='left')
# print(df_all.head(5))
# print(df_all.info())
# df_all.to_excel('./fund_final_plus2.xlsx')

