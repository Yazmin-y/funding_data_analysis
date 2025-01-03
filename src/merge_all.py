import pandas as pd
import numpy as np
import datetime
from dateutil.relativedelta import relativedelta
import cn2an
import re

pd.set_option('display.max_columns', None)


def data_transform():
    df = pd.DataFrame(pd.read_excel('./company_final.xlsx'))
    # df = pd.DataFrame(pd.read_excel('./test.xlsx'))
    # df = pd.DataFrame(pd.read_excel('./managers.xlsx'))
    # df = pd.DataFrame(pd.read_excel('./useful/fund_processed.xlsx'))
    df = df[df['计算截止日期'] != 0]
    dates = pd.to_datetime(df['计算截止日期'].astype('str'))
    print(df.info())
    for index, date in enumerate(dates):
        quarter_end_day = datetime.date(date.year, date.month - (date.month - 1) % 3 + 2, 1) + relativedelta(months=1,
                                                                                                             days=-1)
        if (quarter_end_day.month - date.month) > 1:
            month = (date.month - 1) - (date.month - 1) % 3 + 1  # 10
            newdate = datetime.datetime(date.year, month, 1)
            newdate = newdate + datetime.timedelta(days=-1)
            quarter_end_day = newdate
        quarter_end_day = quarter_end_day.strftime('%Y/%-m/%-d')
        print(quarter_end_day)
        df.loc[index, '季度最后一天'] = quarter_end_day
        #
        # month_end_day = (date + datetime.timedelta(days=-date.day + 1)) + relativedelta(months=1, days=-1)
        # month_end_day = month_end_day.strftime('%Y/%-m/%-d')
        # df.loc[index, '月度最后一天'] = month_end_day
        #
        # last_day_of_last_month = datetime.date(date.year, date.month, 1) - datetime.timedelta(1)
        # last_day_of_last_month = last_day_of_last_month.strftime('%Y/%-m/%-d')
        # df.loc[index, '发行最近一个月末'] = last_day_of_last_month
    print(df.head(10))
    df.to_excel('./useful/company.xlsx')
    # df.to_excel('./useful/return_rate.xlsx')
    # df.to_excel('./useful/managers.xlsx', index=False)
    # df.to_excel('./useful/fund_processed.xlsx', index=False)


def fund_process():
    df = pd.DataFrame(pd.read_excel('./useful/fund.xlsx'))
    no_trans = df['基金转型说明'].astype(str).str.isdigit()
    df = df[no_trans]
    df['发行日期最近一个季度末'] = pd.to_datetime(df['发行日期最近一个季度末'].astype('str')).dt.strftime('%Y/%-m/%-d')
    df.loc[:, '运作方式.封闭运作基金'] = np.where(df['封闭运作期'] == 0, 0, 1)
    df.loc[:, '运作方式.持有期基金'] = np.where(df['基金全称'].str.contains('持有期'), 1, 0)
    df['基金经理（成立）'] = df['基金经理（成立）'].str.split(',', expand=True)[0]
    df.loc[:, '基金分类'] = df['投资类型（三级分类）'].replace(['普通股票型基金', '增强指数型基金', '偏股混合型基金', '平衡混合型基金', '灵活配置型基金',
                                                  'QDII普通股票型基金', 'QDII增强指数型基金', 'QDII偏股混合型基金', 'QDII平衡混合型基金',
                                                  'QDII灵活混合型基金',
                                                  '国际(QDII)偏股混合型基金', '国际(QDII)普通股票型基金'], '主动权益类') \
        .replace(['偏债混合型基金', '混合债券型二级基金', 'QDII偏债混合型基金'], '固收+类') \
        .replace(['被动指数型基金', 'QDII被动指数型基金', '国际(QDII)被动指数型股票基金'], '被动权益类') \
        .replace(['中长期纯债型基金', '短期纯债型基金', '被动指数债券型基金', '混合债券型一级基金',
                  '被动指数型债券基金', '国际(QDII)普通债券型基金', '可转换债券型基金'], '固收类')
    df = df[df['基金分类'].isin(['主动权益类', '固收+类', '被动权益类', '固收类'])]
    # df = pd.get_dummies(df, prefix=['是否收取浮动管理费', '基金分类', '性别', '学历'], prefix_sep='.', columns=['是否收取浮动管理费', '基金分类', '性别', '学历'])
    df.to_excel('./useful/fund_processed.xlsx')


def merge_all():
    fund_df = pd.DataFrame(pd.read_excel('./useful/fund_processed.xlsx'))
    company_df = pd.DataFrame(pd.read_excel('./useful/company_test.xlsx'))
    manager_df = pd.DataFrame(pd.read_excel('./useful/managers.xlsx'))
    rr_df = pd.DataFrame(pd.read_excel('./useful/return_rate.xlsx'))
    fund_df.loc[:, '基金收益类型'] = fund_df['基金分类'].replace(['主动权益类', '被动权益类'], '权益类') \
        .replace(['固收类', '固收+类'], '固定收益类')
    # merge_df = pd.merge(fund_df, company_df, left_on=['基金管理人', '发行日期最近一个季度末'], right_on=['基金公司全称', '季度最后一天'], how='left').drop(['基金公司', '基金公司全称', '季度最后一天'], axis=1)
    merge_df = pd.merge(fund_df, manager_df, left_on=['基金经理（成立）', '基金管理人', '发行日期最近一个季度末'],
                        right_on=['基金经理', '基金公司', '季度最后一天'], how='left').drop(
        ['基金经理', '基金公司', '季度最后一天', '性别', '学历', '证券从业日期'],
        axis=1)
    merge_df = pd.merge(merge_df, rr_df, left_on=['基金经理（成立）', '基金管理人', '发行最近一个月末'],
                        right_on=['现任基金经理', '基金公司全称', '月度最后一天'], how='left').drop(['现任基金经理', '基金公司全称', '月度最后一天'],
                                                                                  axis=1)
    print(merge_df.info())
    print(merge_df.head(5))
    merge_df.to_excel('./useful/tmp.xlsx')
    # merge_df = pd.DataFrame(pd.read_excel('./useful/tmp.xlsx'))
    merge_df = pd.merge(merge_df, company_df, left_on=['基金管理人', '发行日期最近一个季度末', '基金收益类型'],
                        right_on=['基金公司全称', '季度最后一天', '收益类型'], how='left').drop(
        ['基金公司全称', '季度最后一天', '基金收益类型', '计算截止日期'], axis=1)

    print(merge_df.info())
    print(merge_df.head(5))
    merge_df.to_excel('./useful/merged.xlsx', index=False)


def manager_info_merge():
    df = pd.DataFrame(pd.read_excel('./useful/merged.xlsx'))
    manager_df = pd.DataFrame(pd.read_excel('./基金经理大全--2023-3-14.xlsx', header=1))
    manager_df.rename(columns={'Unnamed: 1': '基金经理'}, inplace=True)
    manager_df = manager_df[['基金经理', '性别', '学历', '证券从业日期', '国籍', '最早任职日期', '基金公司']]
    merge_df = pd.merge(df, manager_df, left_on=['基金经理（成立）', '基金管理人'], right_on=['基金经理', '基金公司'], how='left').drop(
        ['基金经理', '基金公司'], axis=1)
    print(merge_df.info())
    merge_df.replace('--', np.nan, inplace=True)
    merge_df.to_excel('./useful/merged2.xlsx', index=False)


def duration(data):
    data = cn2an.transform(data)
    if '持有期' in data:
        # print(data)
        digit = int(re.sub(u"([^\u0030-\u0039])", "", data))
        # digit = filter(str.isdigit, data)
        if '月' in data:
            return digit
        elif '年' in data:
            return digit * 12
        else:
            return digit / 30
    else:
        return


def bank_transform(data):
    bank_list_1 = ['中国银行股份有限公司', '中国农业银行股份有限公司', '中国工商银行股份有限公司',
                   '中国建设银行股份有限公司', '交通银行股份有限公司', '中国邮政储蓄银行股份有限公司']
    bank_list_2 = ['招商银行股份有限公司', '上海浦东发展银行股份有限公司', '中信银行股份有限公司', '中国光大银行股份有限公司', '华夏银行股份有限公司',
                   '中国民生银行股份有限公司', '广发银行股份有限公司', '兴业银行股份有限公司', '平安银行股份有限公司', '浙商银行股份有限公司', '恒丰银行股份有限公司', '渤海银行股份有限公司']
    if data in bank_list_1:
        return '国有银行'
    elif data in bank_list_2:
        return '股份制商业银行'
    elif '银行' in data:
        return '其他银行'
    else:
        return '证券公司'


def dummy_process():
    df = pd.DataFrame(pd.read_excel('./useful/merged2.xlsx'))
    # date_df = df[df['发行日期'] != 0]
    # date2_df = df[df['发行日期'] != 0]
    dates = pd.to_datetime(df['发行日期'].astype('str'))
    dates_2 = pd.to_datetime(df['最早任职日期'].astype('str'))
    df.loc[:, '任职时长'] = (dates - dates_2).map(lambda x: x.days)
    df.loc[:, '发行年份'] = dates.dt.year
    # print(df['基金全称'][df['运作方式.持有期基金'] == 1])
    # df['基金全称'] = df['基金全称'].apply(lambda x: cn2an.transform(x))
    df = df.assign(
        duration=df['基金全称'].apply(lambda x: duration(x))
    )
    df.loc[df['运作方式.持有期基金'] == 0, 'duration'] = df['封闭运作期']
    df = df.assign(
        bank=df['基金托管人'].apply(lambda x: bank_transform(x))
    )
    df.rename(columns={'duration': '封闭运作时间'}, inplace=True)

    df.loc[:, '存在最近一年平均收益率'] = np.where(df['最近一年平均收益率'].isna(), 0, 1)
    df.loc[:, '存在最近三年平均收益率'] = np.where(df['最近三年平均收益率'].isna(), 0, 1)
    df.loc[:, '存在最近一年最高收益率'] = np.where(df['最近一年最高收益率'].isna(), 0, 1)
    df.loc[:, '存在最近三年最高收益率'] = np.where(df['最近三年最高收益率'].isna(), 0, 1)

    df.loc[:, '存在公司最近一年收益率'] = np.where(df['基金公司最近一年收益率'].isna(), 0, 1)
    df.loc[:, '存在公司最近三年收益率'] = np.where(df['基金公司最近三年收益率'].isna(), 0, 1)

    df = pd.get_dummies(df, prefix=['性别', '学历', '托管行', '浮动管理费', '发行年份'],
                        prefix_sep='.', columns=['性别', '学历', 'bank', '是否收取浮动管理费', '发行年份'])
    df.drop(['证券简称', '基金全称', '基金管理人', '基金托管人', '投资类型（三级分类）', '收益类型', '封闭运作期', '发行日期最近一个季度末', '发行日期', '发行最近一个月末',
             '运作方式.持有期基金', '运作方式.封闭运作基金',
             '募集份额上限', '认购份额确认比例', '是否延期募集', '是否提前结束募集', '是否延长募集期', '基金转型说明', '擅长领域', '国籍', '学历.本科', '性别.女', '托管行.证券公司',
             '浮动管理费.否'], axis=1, inplace=True)
    print(df.info())
    manager_col_list = ['基金经理（成立）', '在任基金数', '在管基金总规模(亿元)', '在任公司经理年限', '最近一年平均收益率', '最近一年最高收益率',
                        '最近三年平均收益率', '最近三年最高收益率', '任职时长', '性别.男',
                        '学历.博士', '学历.硕士', '存在最近一年平均收益率', '存在最近三年平均收益率', '存在最近一年最高收益率', '存在最近三年最高收益率']
    company_col_list = ['基金管理人资产净值合计（非货）', '基金公司最近一年收益率', '基金公司最近三年收益率', '存在公司最近一年收益率', '存在公司最近三年收益率']
    product_col_list = ['证券代码', '基金分类', '管理费率', '托管费率', '浮动管理费.是', '封闭运作时间', '认购天数', '发行总份额']
    market_col_list = ['发行首日前日上证综指收盘价', '发行首日前日上证综指30天涨跌幅', '发行首日前日上证综指30天成交量', '发行首日前日跟踪指数30天涨跌幅']
    bank_col_list = ['托管行.其他银行', '托管行.国有银行', '托管行.股份制商业银行']
    date_col_list = ['发行年份.2018', '发行年份.2019', '发行年份.2020', '发行年份.2021', '发行年份.2022']
    tmp1 = df[manager_col_list]
    tmp1.columns = pd.MultiIndex.from_product([['基金经理个人信息'], manager_col_list])
    tmp2 = df[company_col_list]
    tmp2.columns = pd.MultiIndex.from_product([['基金公司情况'], company_col_list])
    tmp3 = df[product_col_list]
    tmp3.columns = pd.MultiIndex.from_product([['产品特征'], product_col_list])
    tmp4 = df[market_col_list]
    tmp4.columns = pd.MultiIndex.from_product([['市场环境变量'], market_col_list])
    tmp5 = df[bank_col_list]
    tmp5.columns = pd.MultiIndex.from_product([['托管行'], bank_col_list])
    tmp6 = df[date_col_list]
    tmp6.columns = pd.MultiIndex.from_product([['发行年份'], date_col_list])
    df = pd.concat([tmp3, tmp1, tmp2, tmp5, tmp4, tmp6], axis=1)
    print(df.info())
    # df = df.reindex(columns=manager_col_list+company_col_list+product_col_list+market_col_list+bank_col_list)
    # print(df.info())
    grouped = df.groupby(by=('产品特征', '基金分类'))
    for value, group in grouped:
        filename = './grouped/' + str(value) + '.' + 'xlsx'
        if str(value) != '被动权益类':
            group.drop(('市场环境变量', '发行首日前日跟踪指数30天涨跌幅'), axis=1, inplace=True)
        try:
            f = open(filename, 'w')
            if f:
                # 清空文件
                f.truncate()
                # 写入新文件
                group.to_excel(filename)
        except Exception as e:
            print(e)
    df.to_excel('./useful/merged3.xlsx')


data_transform()
fund_process()
merge_all()
manager_info_merge()
dummy_process()

