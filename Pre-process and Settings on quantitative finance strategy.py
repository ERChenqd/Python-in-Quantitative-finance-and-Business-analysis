# -*- coding: utf-8 -*-
"""
Created on Sun Jul 13 20:05:53 2025

@author: 24333
"""
import warnings
import os 
import ast
import numpy as np

from Functions import * 

warnings.filterwarnings('ignore') #作用：忽略所有的警告信息，让程序运行时不显示任何警告提示。
pd.set_option('expand_frame_repr', False)  # 当列太多时不换行，可用可不用
pd.set_option('display.max_rows', 5000)  # 最多显示数据的行数，可用可不用

_= os.path.abspath(os.path.dirname(__file__))  # 返回当前文件路径
root_path = os.path.abspath(os.path.join(_, '.'))  # 返回根目录文件夹

c_rate = 1 / 10000  # 手续费，注意调节
t_rate = 1 / 1000  # 印花税，注意调节
period = 'M' #持仓周期：月，一月一换

df = pd.read_pickle(r'******.pkl') #需要自己填写*区域
file_name = '******'
df['交易日期'] = pd.to_datetime(df['交易日期'])

# 读取沪深300指数数据作为策略收益比较基准
index_path = r'sh000300.csv' #只是定义了文件路径
index_data = import_index_data(index_path, back_trader_start='2009-01-01', back_trader_end='2023-07-19')#还需要这个来读取数据
#import_index_data是Functions.py里面用户提前编写好的公式
index_data.head(20)

# =====选股
# ===过滤股票
# 剔除不能交易的异常情况
df = df[df['下日_是否交易'] == 1]
df = df[df['下日_开盘涨停'] == False]
# 剔除交易日期不足的
df = df[df['交易天数'] / df['市场交易天数'] >= 0.8]
# 剔除ST股/退市股
df = df[df['下日_是否ST'] == False]
df = df[df['下日_是否退市'] == False]
df = df[df['交易日期'] <= pd.to_datetime('2023-07-01')]

"""===按照策略选股：不同策略填充在一下即可"""





"""===按照策略选股：不同策略填充在以上即可"""

# =====计算选中股票每天的资金曲线
# 计算每日资金曲线
equity = pd.merge(left=index_data, right=select_stock[['交易日期', '买入股票代码']], on=['交易日期'],
                  how='left', sort=True)  # 将选股结果和大盘指数合并

equity['持有股票代码'] = equity['买入股票代码'].shift()
equity['持有股票代码'].fillna(method='ffill', inplace=True)
equity.dropna(subset=['持有股票代码'], inplace=True)

select_stock.sort_values(by='交易日期', inplace=True)

equity['涨跌幅'] = select_stock['选股下周期每天涨跌幅'].sum()
equity['equity_curve'] = (equity['涨跌幅'] + 1).cumprod()
equity['benchmark'] = (equity['指数涨跌幅'] + 1).cumprod()
equity.to_csv('策略结果.csv', encoding='gbk')

# =====计算策略评价指标
rtn = strategy_evaluate(equity, select_stock)
print(rtn, '\n')
draw_equity_curve_mat2(equity)