import pandas as pd

# 读取Excel文件
df = pd.read_excel('test.xlsx', engine='openpyxl')

# 显示前5行数据
print(df.head())
