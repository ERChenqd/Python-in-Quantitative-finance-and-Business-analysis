#%%
import pandas as pd

import os


from datetime import datetime, timedelta

# Get the current date
now = datetime.now()

current_year = now.year

# Calculate the start of the current week (Monday)
start_of_current_week = now - timedelta(days=now.weekday())

# Calculate the start of the last work week (previous Monday)
start_of_last_work_week = start_of_current_week - timedelta(days=7)

# Find the week number of the last work week
last_work_week_number = start_of_last_work_week.isocalendar()[1]

names_mappings = {
    "DH1": "东恒1号",
    "HT7": "海天7号",
    "HT3": "海天3号",
    "QF1": "青峰1号",
    "XY1": "喜悦1号",
    "YJ3": "盈玖3号",
    "ZS1": "中盛1号",
    "LH1": "量化1号",
    "HR1": "合瑞1号",
    "QY1": "青银1号",
    "LS1": "岭盛1号",
    "XJ1": "兴建1号",
    "FH1": "福慧1号",
    "FH5": "福慧5号",
    "FH2": "福慧2号",
    "FH3": "福慧3号",
    "FH10": "福慧10号",
    "CX1": "春晓1号",

}

names_mappings["CY1"]="晨跃1号"





# Initialize a list to hold all new rows
all_rows = []

records_holder_path = r"C:\\Users\\19063\\Desktop\\转债交易记录处理\\上周交易记录文件"



if not os.path.exists(records_holder_path):
    print(f"未找到交易记录文件路径：{records_holder_path}，请创建此路径并将交易记录文件放入。按回车键退出并重新运行此小程序：")
    input()
    exit()


print("上周交易记录文件中的文件有：")
for file_name in os.listdir(records_holder_path):
    print(file_name)
confirmation = input("是否继续处理？ (Y/N): ").strip().upper()

if confirmation != "Y":
    print("请按回车键退出程序:")
    input()
    exit()


# Process each file
for file_name in os.listdir(records_holder_path):
    # Load the Excel file
    file_path = os.path.join(records_holder_path,file_name)

    #base_string = file_name.split('\\')[-1]
    date_string = file_name.split('.')[0]
    try:
        file_date = datetime.strptime(date_string, '%Y%m%d').strftime('%Y%m%d')

    except ValueError:
        print(f"已忽略文件‘{file_name}’因为名称的格式并不是YYYYMMDD.xlsx,无法匹配日期。")
        continue

    df = pd.read_excel(file_path)

    # Filter rows where the 'direction' column is either 'BUY' or 'SELL'
    df = df[df['direction'].isin(['BUY', 'SELL'])]

    # Add '成交金额' column as the product of columns F and G
    df['成交金额'] = df.iloc[:, 4] * df.iloc[:, 5]  # Adjust column indices as needed

    # Get unique product IDs
    unique_products = df['product_id'].unique()

    # Process each product
    for product in unique_products:
        # Filter rows for the current product
        product_rows = df[df['product_id'] == product]

        # Calculate sums
        sum_market_value_dif = product_rows['market_value_dif'].sum()
        sum_refer_value_dif = product_rows['refer_value_dif'].sum()
        sum_chengjiao = product_rows['成交金额'].sum()

        refer_value_percentage = (sum_refer_value_dif / sum_chengjiao) * 100
        market_value_percentage = (sum_market_value_dif / sum_chengjiao) * 100

        # Create a new row with sums and transformed name
        new_row = {
            '日期': pd.to_datetime(file_date, format='%Y%m%d').strftime('%Y/%m/%d'),
            '产品名称': names_mappings.get(product, product),
            #'交易员': traders_mappings.get(product,product),
            'market_value_dif': sum_market_value_dif,
            'refer_value_dif': sum_refer_value_dif,
            '成交金额': sum_chengjiao,
            '参考价差异比例': f'{refer_value_percentage:.2f}%',
            '市场价差异比例': f'{market_value_percentage:.2f}%'
        }
        all_rows.append(new_row)

     

# Create a DataFrame from all new rows
all_rows_df = pd.DataFrame(all_rows)



# Group by '产品名称' and calculate sums
grouped = all_rows_df.groupby('产品名称')
summary = grouped[['market_value_dif', 'refer_value_dif', '成交金额']].sum().reset_index()

# Calculate percentages for the summary
summary['参考价差异比例'] = summary.apply(lambda row: f"{(row['refer_value_dif'] / row['成交金额'] * 100) if row['成交金额'] != 0 else 0:.2f}%", axis=1)
summary['市场价差异比例'] = summary.apply(lambda row: f"{(row['market_value_dif'] / row['成交金额'] * 100) if row['成交金额'] != 0 else 0:.2f}%", axis=1)

summary['日期'] = '汇总部分'

# Reorder columns to match the original DataFrame
summary = summary[['日期', '产品名称', 'market_value_dif', 'refer_value_dif', '成交金额', '参考价差异比例', '市场价差异比例']]

all_rows_df = all_rows_df.sort_values(by=['日期'], ascending=[True])

# Create two empty rows as DataFrame
separator = pd.DataFrame([['']*len(all_rows_df.columns)]*2, columns=all_rows_df.columns)

# Combine the original, separator, and summary DataFrames
final_df = pd.concat([all_rows_df, separator, summary], ignore_index=True, sort=True)
final_df = final_df[['日期', '产品名称', 'market_value_dif', 'refer_value_dif', '成交金额', '参考价差异比例', '市场价差异比例']]

# Specify the folder path where you want to save the file
output_path = r"C:\\Users\\19063\\Desktop\\转债交易记录处理\\转债交易汇总output"  #mac测试的路径 # Change this to your desired folder path

#folder_path = r"\\yangtong\share\风控\交易核对\转债\交易汇总output"  # Change this to your desired folder path

#下面这句是新加的


# Check if the folder path exists
if not os.path.exists(output_path):
    print(f"\n以下路径不存在: {output_path}")
    input("请按回车退出程序:")
    exit()

# Save the new rows to a new Excel file in the specified folder
file_name = f"{current_year}转债交易汇总第{last_work_week_number}周.xlsx"
try:
    final_df.to_excel(os.path.join(output_path, file_name), index=False)
    print("转债交易汇总文件已成功生成。")
    input("请按回车键退出程序:")
    exit()

    #new_rows_df.to_excel(os.path.join(folder_path, file_name), index=False)
except Exception as e:
    # Handle exceptions that might occur during file saving
    print(f"\n保存文件时出现未知错误: {e}")
    input("请按回车退出程序:")
    exit()





# ...

# %%
