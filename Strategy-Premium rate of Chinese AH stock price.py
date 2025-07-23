import warnings
import os
from Functions import *
import ast
import numpy as np

warnings.filterwarnings('ignore')
pd.set_option('expand_frame_repr', False)  # 当列太多时不换行
pd.set_option('display.max_rows', 5000)  # 最多显示数据的行数

_ = os.path.abspath(os.path.dirname(__file__))  # 返回当前文件路径
root_path = os.path.abspath(os.path.join(_, '.'))  # 返回根目录文件夹

c_rate = 1 / 10000  # 手续费
t_rate = 1 / 1000  # 印花税


period = 'M' #每一次的持仓为“月”
df = pd.read_pickle(r'邢不行-AH溢价率策略数据.pkl')
file_name = 'AH溢价率选股_月频'
df['交易日期'] = pd.to_datetime(df['交易日期'])

# 读取沪深300指数数据
index_path = r'sh000300.csv'
index_data = import_index_data(index_path, back_trader_start='2009-01-01', back_trader_end='2023-07-19')

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


"""===按照策略选股：不同策略修改此处即可"""

# 因子1
df['排名'] = df.groupby('交易日期')['AH溢价率'].rank(pct=False, ascending=True, method='first')  # 根据选股因子对股票进行排名,可以调整参数。
df = df[df['排名'] <= 10]  # 选取排名靠前的股票

# 选取排名靠前的股票
print(df.head(6))

# 删除数据
df.dropna(subset=['下周期每天涨跌幅'], inplace=True)

# ===挑选出选中股票
df['股票代码'] += ' '
df['股票名称'] += ' '
group = df.groupby('交易日期')
select_stock = pd.DataFrame()
select_stock['股票数量'] = group['股票名称'].size()
select_stock['买入股票代码'] = group['股票代码'].sum()
select_stock['买入股票名称'] = group['股票名称'].sum()
select_stock.sort_values(by='交易日期', inplace=True)

select_stock['选股下周期每天资金曲线'] = group['选股下周期每天资金曲线'].apply(
    lambda x: list(np.mean([ast.literal_eval(i) for i in list(x)], axis=0)))

# 扣除买入手续费
select_stock['选股下周期每天资金曲线'] = [[item * (1 - c_rate) for item in x] for x in
                                          select_stock['选股下周期每天资金曲线']]

# 扣除卖出手续费、印花税。最后一天的资金曲线值，扣除印花税、手续费
select_stock['选股下周期每天资金曲线'] = select_stock['选股下周期每天资金曲线'].apply(
    lambda x: list(x[:-1]) + [x[-1] * (1 - c_rate - t_rate)])

select_stock['选股下周期每天涨跌幅'] = select_stock['选股下周期每天资金曲线'].apply(
    lambda x: list(pd.DataFrame([1] + x).pct_change()[0].iloc[1:]))


# 为了防止有的周期没有选出股票，创造一个空的df，用于填充不选股的周期
empty_df = create_empty_data_week(index_data, period, 0)
empty_df.update(select_stock)  # 将选股结果更新到empty_df上
select_stock = empty_df

# 计算整体资金曲线
select_stock.reset_index(inplace=True)
select_stock['选股下周期涨跌幅'] = select_stock['选股下周期每天涨跌幅'].apply(lambda x: np.cumprod(np.array(x)+1)[-1]-1)
select_stock['资金曲线'] = (select_stock['选股下周期涨跌幅'] + 1).cumprod()
select_stock.sort_values(by='交易日期', inplace=True)
print(select_stock.tail())

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

